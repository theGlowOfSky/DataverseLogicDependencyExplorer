// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using DataverseLogicDependencyExplorer.Models;

namespace DataverseLogicDependencyExplorer.Services.Scanners
{
    /// <summary>
    /// Reads plugin registration metadata only (steps, filtering attributes, images, plugin types, assemblies). Compiled plugin
    /// code is never inspected, and step configuration strings (which may hold secrets) are deliberately not retrieved.
    /// </summary>
    public class PluginRegistrationScanner : ScannerBase
    {
        public override ScannerKind Kind { get { return ScannerKind.Plugins; } }

        private class StepImage
        {
            public string Name, Alias, Attributes, TypeText;
            public bool AllAttributes { get { return string.IsNullOrWhiteSpace(Attributes); } }
        }

        public static string StageName(int? stage)
        {
            switch (stage)
            {
                case 10: return "PreValidation (10)";
                case 20: return "PreOperation (20)";
                case 30: return "MainOperation (30)";
                case 40: return "PostOperation (40)";
                default: return stage.HasValue ? "Stage " + stage.Value : "Unknown stage";
            }
        }

        public static string ImageTypeName(int? t)
        {
            return t == 0 ? "PreImage" : t == 1 ? "PostImage" : t == 2 ? "PreImage and PostImage" : "Image";
        }

        protected override void Execute(ScanContext ctx, ScanOutcome outcome, CancellationToken ct)
        {
            ScanRequest req = ctx.Request;
            ScanDiagnostic diag = outcome.Diagnostic;
            diag.Limit("Only plugin REGISTRATION metadata is inspected. Dataverse does not expose which columns a compiled plugin reads or writes in code, so a plugin may depend on this table/column without appearing here.");
            diag.Limit("Step configuration strings are not retrieved (they can contain secrets).");

            ctx.Report("Retrieving plugin steps registered on " + req.TableLogicalName + "...");
            List<Entity> steps = RetrieveSteps(ctx, ct);
            diag.RecordsExamined = steps.Count;
            if (steps.Count == 0) return;

            Dictionary<Guid, List<StepImage>> images = RetrieveImages(ctx, steps, ct, diag);
            Dictionary<Guid, string[]> handlers = RetrieveHandlers(ctx, steps, ct, diag);

            foreach (Entity step in steps)
            {
                ct.ThrowIfCancellationRequested();
                List<StepImage> stepImages;
                if (!images.TryGetValue(step.Id, out stepImages)) stepImages = new List<StepImage>();
                string[] handler;
                handlers.TryGetValue(step.Id, out handler);

                if (req.Mode == SearchMode.Table) outcome.Results.Add(TableResult(ctx, step, stepImages, handler));
                else ColumnResults(ctx, outcome, step, stepImages, handler);
            }
        }

        private List<Entity> RetrieveSteps(ScanContext ctx, CancellationToken ct)
        {
            var q = new QueryExpression("sdkmessageprocessingstep")
            {
                ColumnSet = new ColumnSet("sdkmessageprocessingstepid", "name", "stage", "mode", "rank", "statecode",
                                          "filteringattributes", "eventhandler", "supporteddeployment", "ismanaged")
            };
            LinkEntity filter = q.AddLink("sdkmessagefilter", "sdkmessagefilterid", "sdkmessagefilterid", JoinOperator.Inner);
            filter.EntityAlias = "f";
            filter.Columns = new ColumnSet("primaryobjecttypecode");
            filter.LinkCriteria.AddCondition("primaryobjecttypecode", ConditionOperator.Equal, ctx.Request.TableLogicalName);
            LinkEntity message = q.AddLink("sdkmessage", "sdkmessageid", "sdkmessageid", JoinOperator.Inner);
            message.EntityAlias = "m";
            message.Columns = new ColumnSet("name");
            q.AddOrder("name", OrderType.Ascending);
            return DataverseHelper.RetrieveAll(ctx.Service, q, ct, 100).ToList();
        }

        private Dictionary<Guid, List<StepImage>> RetrieveImages(ScanContext ctx, List<Entity> steps, CancellationToken ct, ScanDiagnostic diag)
        {
            var result = new Dictionary<Guid, List<StepImage>>();
            try
            {
                foreach (List<Guid> chunk in DataverseHelper.Chunk(steps.Select(s => s.Id), 50))
                {
                    var q = new QueryExpression("sdkmessageprocessingstepimage")
                    {
                        ColumnSet = new ColumnSet("sdkmessageprocessingstepid", "name", "entityalias", "imagetype", "attributes")
                    };
                    q.Criteria.AddCondition("sdkmessageprocessingstepid", ConditionOperator.In, chunk.Cast<object>().ToArray());
                    foreach (Entity e in DataverseHelper.RetrieveAll(ctx.Service, q, ct, 100))
                    {
                        EntityReference stepRef = DataverseHelper.Ref(e, "sdkmessageprocessingstepid");
                        if (stepRef == null) continue;
                        List<StepImage> list;
                        if (!result.TryGetValue(stepRef.Id, out list)) result[stepRef.Id] = list = new List<StepImage>();
                        list.Add(new StepImage
                        {
                            Name = DataverseHelper.Str(e, "name"),
                            Alias = DataverseHelper.Str(e, "entityalias"),
                            Attributes = DataverseHelper.Str(e, "attributes"),
                            TypeText = ImageTypeName(DataverseHelper.Opt(e, "imagetype"))
                        });
                    }
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                diag.Status = ScanStatus.PartiallySupported;
                diag.PermissionLimited = DataverseHelper.IsPermissionDenied(ex);
                diag.Warn("Step images could not be read (" + DataverseHelper.Describe(ex) + "); image attributes were not inspected.");
            }
            return result;
        }

        /// <summary>Step id to [plugin type name, assembly description].</summary>
        private Dictionary<Guid, string[]> RetrieveHandlers(ScanContext ctx, List<Entity> steps, CancellationToken ct, ScanDiagnostic diag)
        {
            var map = new Dictionary<Guid, string[]>();
            var typeIds = new HashSet<Guid>();
            foreach (Entity s in steps)
            {
                EntityReference h = DataverseHelper.Ref(s, "eventhandler");
                if (h == null) continue;
                if (h.LogicalName == "plugintype") typeIds.Add(h.Id);
                else map[s.Id] = new[] { h.LogicalName + " (" + (h.Name ?? h.Id.ToString()) + ")", null };
            }

            var types = new Dictionary<Guid, string[]>();
            try
            {
                foreach (List<Guid> chunk in DataverseHelper.Chunk(typeIds, 50))
                {
                    var q = new QueryExpression("plugintype") { ColumnSet = new ColumnSet("plugintypeid", "typename", "name") };
                    q.Criteria.AddCondition("plugintypeid", ConditionOperator.In, chunk.Cast<object>().ToArray());
                    LinkEntity asm = q.AddLink("pluginassembly", "pluginassemblyid", "pluginassemblyid", JoinOperator.LeftOuter);
                    asm.EntityAlias = "a";
                    asm.Columns = new ColumnSet("name", "version", "isolationmode");
                    foreach (Entity e in DataverseHelper.RetrieveAll(ctx.Service, q, ct, 100))
                    {
                        string iso = DataverseHelper.Opt(e, "a.isolationmode") == 2 ? "Sandbox" : DataverseHelper.Opt(e, "a.isolationmode") == 1 ? "None" : null;
                        string asmText = DataverseHelper.Str(e, "a.name") == null ? null
                            : "Assembly " + DataverseHelper.Str(e, "a.name") + " v" + DataverseHelper.Str(e, "a.version") + (iso == null ? string.Empty : ", isolation: " + iso);
                        types[e.Id] = new[] { DataverseHelper.Str(e, "typename") ?? DataverseHelper.Str(e, "name"), asmText };
                    }
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                diag.Status = ScanStatus.PartiallySupported;
                diag.PermissionLimited = diag.PermissionLimited || DataverseHelper.IsPermissionDenied(ex);
                diag.Warn("Plugin types/assemblies could not be read (" + DataverseHelper.Describe(ex) + "); plugin class names are missing from the findings.");
            }

            foreach (Entity s in steps)
            {
                EntityReference h = DataverseHelper.Ref(s, "eventhandler");
                string[] t;
                if (h != null && h.LogicalName == "plugintype" && types.TryGetValue(h.Id, out t)) map[s.Id] = t;
            }
            return map;
        }

        private DependencyResult Base(ScanContext ctx, Entity step, string[] handler, ConfidenceLevel confidence)
        {
            DependencyResult r = ctx.NewResult(ScannerKind.Plugins, ComponentType.PluginStep, confidence);
            r.ComponentName = DataverseHelper.Str(step, "name") ?? "(unnamed step)";
            r.ComponentId = step.Id.ToString();
            r.ComponentUniqueName = handler == null ? null : handler[0];
            r.ParentTable = DataverseHelper.Str(step, "f.primaryobjecttypecode");
            r.SourceResource = handler == null ? "sdkmessageprocessingstep" : handler[0];
            r.AdditionalInfo = handler == null ? null : handler[1];
            int? state = DataverseHelper.Opt(step, "statecode");
            bool? managed = DataverseHelper.Bool(step, "ismanaged");
            var parts = new List<string> { state == 0 ? "Enabled" : state == 1 ? "Disabled" : "State " + state };
            if (managed.HasValue) parts.Add(managed.Value ? "Managed" : "Unmanaged");
            r.ComponentState = string.Join(", ", parts);
            r.AnalysisMethod = "Plugin registration metadata (SDK message processing step)";
            r.Limitation = "Registration metadata only: columns used inside the compiled plugin code are not visible.";
            return r;
        }

        private static string StepSummary(Entity step, List<StepImage> images)
        {
            string filtering = DataverseHelper.Str(step, "filteringattributes");
            string text = "Message: " + DataverseHelper.Str(step, "m.name")
                + "; Stage: " + StageName(DataverseHelper.Opt(step, "stage"))
                + "; Mode: " + (DataverseHelper.Opt(step, "mode") == 1 ? "Asynchronous" : "Synchronous")
                + "; Rank: " + DataverseHelper.Str(step, "rank")
                + "; Filtering attributes: " + (string.IsNullOrWhiteSpace(filtering) ? "(none - all attributes)" : filtering);
            if (images.Count > 0)
                text += "; Images: " + string.Join(" | ", images.Select(i => i.TypeText + " '" + (i.Name ?? i.Alias) + "' [" + (i.AllAttributes ? "all attributes" : i.Attributes) + "]"));
            return text;
        }

        private DependencyResult TableResult(ScanContext ctx, Entity step, List<StepImage> images, string[] handler)
        {
            DependencyResult r = Base(ctx, step, handler, ConfidenceLevel.ConfirmedMetadataReference);
            r.MatchType = "Step registered on table";
            r.MatchedValue = ctx.Request.TableLogicalName;
            r.Evidence = StepSummary(step, images);
            r.Location = "sdkmessagefilter.primaryobjecttypecode";
            r.Explanation = "The plugin step is registered on the selected table.";
            return r;
        }

        private void ColumnResults(ScanContext ctx, ScanOutcome outcome, Entity step, List<StepImage> images, string[] handler)
        {
            string column = ctx.Request.ColumnLogicalName;
            string filtering = DataverseHelper.Str(step, "filteringattributes");
            string message = DataverseHelper.Str(step, "m.name");
            string summary = StepSummary(step, images);

            if (TextLocator.ListContains(filtering, column))
            {
                DependencyResult r = Base(ctx, step, handler, ConfidenceLevel.ConfirmedMetadataReference);
                r.MatchType = "Filtering attribute";
                r.MatchedValue = column;
                r.Evidence = summary;
                r.Location = "sdkmessageprocessingstep.filteringattributes";
                r.Explanation = "The column is listed in the step's filtering attributes, so a change to it triggers the step.";
                outcome.Results.Add(r);
            }
            else if (string.IsNullOrWhiteSpace(filtering) && string.Equals(message, "Update", StringComparison.OrdinalIgnoreCase))
            {
                DependencyResult r = Base(ctx, step, handler, ConfidenceLevel.PossibleOrDynamic);
                r.MatchType = "Update step without filtering attributes";
                r.MatchedValue = column;
                r.Evidence = summary;
                r.Location = "sdkmessageprocessingstep.filteringattributes (empty)";
                r.Explanation = "The Update step has no filtering attributes, so it runs when any column changes, including this one.";
                outcome.Results.Add(r);
            }

            foreach (StepImage img in images)
            {
                if (!img.AllAttributes && TextLocator.ListContains(img.Attributes, column))
                {
                    DependencyResult r = Base(ctx, step, handler, ConfidenceLevel.ConfirmedMetadataReference);
                    r.MatchType = "Image attribute (" + img.TypeText + ")";
                    r.MatchedValue = column;
                    r.Evidence = img.TypeText + " '" + (img.Name ?? img.Alias) + "' attributes: " + img.Attributes + "  --  " + summary;
                    r.Location = "sdkmessageprocessingstepimage.attributes";
                    r.Explanation = "The column is selected in the step's " + img.TypeText + ", so the plugin can read its value from the image.";
                    outcome.Results.Add(r);
                }
                else if (img.AllAttributes)
                {
                    DependencyResult r = Base(ctx, step, handler, ConfidenceLevel.PossibleOrDynamic);
                    r.MatchType = "Image includes all attributes (" + img.TypeText + ")";
                    r.MatchedValue = column;
                    r.Evidence = img.TypeText + " '" + (img.Name ?? img.Alias) + "' has no attribute list (all attributes)  --  " + summary;
                    r.Location = "sdkmessageprocessingstepimage.attributes (empty)";
                    r.Explanation = "The image has no attribute list, so it includes this column.";
                    outcome.Results.Add(r);
                }
            }
        }
    }
}
