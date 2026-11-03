// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using DataverseLogicDependencyExplorer.Models;

namespace DataverseLogicDependencyExplorer.Services.Scanners
{
    /// <summary>
    /// Classic processes: workflows (category 0), dialogs (1), actions (3) and business process flows (4). Business rules (2) and
    /// cloud flows (5) are handled by their own scanners; desktop flows and other categories are deliberately not included.
    /// </summary>
    public class ClassicWorkflowScanner : ScannerBase
    {
        private static readonly XNamespace X = "http://schemas.microsoft.com/winfx/2006/xaml";

        public override ScannerKind Kind { get { return ScannerKind.ClassicWorkflows; } }

        public static string CategoryName(int? category)
        {
            switch (category)
            {
                case 0: return "Workflow";
                case 1: return "Dialog";
                case 3: return "Action";
                case 4: return "Business Process Flow";
                default: return category.HasValue ? "Category " + category.Value : "Process";
            }
        }

        protected override void Execute(ScanContext ctx, ScanOutcome outcome, CancellationToken ct)
        {
            ScanRequest req = ctx.Request;
            ScanDiagnostic diag = outcome.Diagnostic;
            diag.Limit("Inspects process definition rows (type 1) of categories Workflow, Dialog, Action and Business Process Flow. Business rules, cloud flows and desktop flows are excluded from this scanner.");
            diag.Limit("Definition formats differ by category; XAML is parsed when it is well-formed. Not every process record exposes a readable XAML definition.");
            if (!req.DeepProcessScan)
                diag.Limit("Only processes whose primary table is '" + req.TableLogicalName + "' were retrieved. Processes of other tables (and unbound actions) that update or read this table are only found with the cross-table (deep) scan option.");

            ctx.Report("Retrieving process definitions...");
            List<Entity> processes = Retrieve(ctx, ct, diag);
            diag.RecordsExamined = processes.Count;

            int noDefinition = 0, otherTableSkipped = 0;
            foreach (Entity p in processes)
            {
                ct.ThrowIfCancellationRequested();
                string primary = DataverseHelper.Str(p, "primaryentity");
                bool onTable = string.Equals(primary, req.TableLogicalName, StringComparison.OrdinalIgnoreCase);
                string xaml = DataverseHelper.Str(p, "xaml");
                string clientData = DataverseHelper.Str(p, "clientdata");
                if (string.IsNullOrWhiteSpace(xaml) && string.IsNullOrWhiteSpace(clientData)) noDefinition++;

                if (req.Mode == SearchMode.Table)
                {
                    if (onTable) outcome.Results.Add(TablePrimary(ctx, p));
                    else ScanTableCrossReference(ctx, outcome, p, xaml);
                }
                else
                {
                    otherTableSkipped += ScanColumn(ctx, outcome, p, onTable, xaml, clientData);
                }
            }

            if (noDefinition > 0)
                diag.Warn(noDefinition + " process(es) expose no readable definition (XAML/clientdata empty); they could not be searched for references.");
            if (otherTableSkipped > 0)
                diag.Limit(otherTableSkipped + " match(es) of the column name were ignored because the XAML step targets a different table.");
        }

        private List<Entity> Retrieve(ScanContext ctx, CancellationToken ct, ScanDiagnostic diag)
        {
            Func<bool, QueryExpression> build = full =>
            {
                var cols = new List<string> { "workflowid", "name", "uniquename", "category", "primaryentity", "statecode", "statuscode", "ismanaged", "type" };
                if (full) { cols.Add("xaml"); cols.Add("clientdata"); }
                var q = new QueryExpression("workflow") { ColumnSet = new ColumnSet(cols.ToArray()) };
                q.Criteria.AddCondition("category", ConditionOperator.In, new object[] { 0, 1, 3, 4 });
                q.Criteria.AddCondition("type", ConditionOperator.Equal, 1);
                if (!ctx.Request.DeepProcessScan)
                    q.Criteria.AddCondition("primaryentity", ConditionOperator.Equal, ctx.Request.TableLogicalName);
                q.AddOrder("name", OrderType.Ascending);
                return q;
            };

            try
            {
                return DataverseHelper.RetrieveAll(ctx.Service, build(true), ct, 25).ToList();
            }
            catch (FaultException<OrganizationServiceFault> ex) when (DataverseHelper.IsMissingMetadata(ex))
            {
                diag.Status = ScanStatus.PartiallySupported;
                diag.Warn("The 'xaml'/'clientdata' columns could not be read in this environment; process definitions were not inspected, only their metadata.");
                return DataverseHelper.RetrieveAll(ctx.Service, build(false), ct, 25).ToList();
            }
        }

        private DependencyResult Base(ScanContext ctx, Entity p, ConfidenceLevel confidence)
        {
            int? category = DataverseHelper.Opt(p, "category");
            DependencyResult r = ctx.NewResult(ScannerKind.ClassicWorkflows, ComponentType.ClassicWorkflow, confidence);
            r.ComponentName = DataverseHelper.Str(p, "name") ?? "(unnamed)";
            r.ComponentId = p.Id.ToString();
            r.ComponentUniqueName = DataverseHelper.Str(p, "uniquename");
            r.ParentTable = DataverseHelper.Str(p, "primaryentity");
            r.ComponentState = CategoryName(category) + "; " + BusinessRuleScanner.StateText(p);
            return r;
        }

        private DependencyResult TablePrimary(ScanContext ctx, Entity p)
        {
            DependencyResult r = Base(ctx, p, ConfidenceLevel.ConfirmedMetadataReference);
            r.MatchType = CategoryName(DataverseHelper.Opt(p, "category")) + " primary table";
            r.MatchedValue = ctx.Request.TableLogicalName;
            r.Evidence = "workflow.primaryentity = '" + DataverseHelper.Str(p, "primaryentity") + "'";
            r.SourceResource = "workflow record";
            r.Location = "workflow.primaryentity";
            r.AnalysisMethod = "Dataverse process metadata (workflow table)";
            r.Explanation = "The process is defined on the selected table.";
            return r;
        }

        private void ScanTableCrossReference(ScanContext ctx, ScanOutcome outcome, Entity p, string xaml)
        {
            if (string.IsNullOrWhiteSpace(xaml)) return;
            XDocument doc;
            try { doc = XDocument.Parse(xaml); }
            catch (XmlException ex)
            {
                outcome.Diagnostic.MalformedDefinitions++;
                outcome.Diagnostic.Warn("The XAML of process '" + DataverseHelper.Str(p, "name") + "' is malformed (" + ex.Message + ").");
                return;
            }
            List<XmlValueMatch> matches = XmlPath.FindExactValues(doc, ctx.Request.TableLogicalName)
                .Where(m => IsEntityNameContext(m)).ToList();
            foreach (XmlValueMatch m in matches.Take(10))
            {
                DependencyResult r = Base(ctx, p, ConfidenceLevel.StrongDefinitionReference);
                r.MatchType = "Cross-table reference in process XAML";
                r.MatchedValue = ctx.Request.TableLogicalName;
                r.Evidence = XmlPath.Render(m) + StepSuffix(m);
                r.SourceResource = "xaml";
                r.Location = m.Path;
                r.AnalysisMethod = "Parsed XAML";
                r.Explanation = "A step of a process on '" + DataverseHelper.Str(p, "primaryentity") + "' names the selected table as its target entity.";
                outcome.Results.Add(r);
            }
        }

        /// <returns>Number of matches ignored because they belong to another table.</returns>
        private int ScanColumn(ScanContext ctx, ScanOutcome outcome, Entity p, bool onTable, string xaml, string clientData)
        {
            string column = ctx.Request.ColumnLogicalName;
            string table = ctx.Request.TableLogicalName;
            int skipped = 0;
            bool parsed = false;

            if (!string.IsNullOrWhiteSpace(xaml))
            {
                try
                {
                    XDocument doc = XDocument.Parse(xaml);
                    parsed = true;
                    foreach (XmlValueMatch m in XmlPath.FindExactValues(doc, column))
                    {
                        string entity = EntityContext(m);
                        bool targetsTable = entity != null && string.Equals(entity, table, StringComparison.OrdinalIgnoreCase);
                        bool targetsOther = entity != null && !targetsTable;
                        if (targetsOther) { skipped++; continue; }

                        ConfidenceLevel conf = (targetsTable || onTable) ? ConfidenceLevel.StrongDefinitionReference : ConfidenceLevel.PossibleOrDynamic;
                        DependencyResult r = Base(ctx, p, conf);
                        r.MatchType = "Column reference in process XAML (" + m.Kind + ")";
                        r.MatchedValue = column;
                        r.Evidence = XmlPath.Render(m) + StepSuffix(m) + (entity != null ? "  [step entity: " + entity + "]" : string.Empty);
                        r.SourceResource = "xaml";
                        r.Location = m.Path;
                        r.AnalysisMethod = "Parsed XAML";
                        r.Explanation = targetsTable
                            ? "A step explicitly targets table '" + table + "' and names column '" + column + "'."
                            : onTable
                                ? "The process is defined on '" + table + "' and a step names column '" + column + "'."
                                : "A step of a process on another table names '" + column + "' but its target entity could not be determined.";
                        if (!targetsTable && !onTable)
                            r.Limitation = "The column name matched, but the step's target table is not stated, so it may belong to a different table.";
                        else if (entity == null)
                            r.Limitation = "The step's target entity is not stated in the XAML; the column is assumed to belong to the process's primary table.";
                        outcome.Results.Add(r);
                    }
                }
                catch (XmlException ex)
                {
                    outcome.Diagnostic.MalformedDefinitions++;
                    outcome.Diagnostic.Warn("The XAML of process '" + DataverseHelper.Str(p, "name") + "' is malformed (" + ex.Message + "); a text search was used instead.");
                    if (TextLocator.ContainsWord(xaml, column)) outcome.Results.Add(TextFallback(ctx, p, "xaml", xaml, column, onTable));
                }
            }

            // Business process flows and processes without XAML may keep their definition in clientdata (format not guaranteed).
            if (!parsed && !string.IsNullOrWhiteSpace(clientData) && TextLocator.ContainsWord(clientData, column))
                outcome.Results.Add(TextFallback(ctx, p, "clientdata", clientData, column, onTable));

            return skipped;
        }

        private DependencyResult TextFallback(ScanContext ctx, Entity p, string source, string text, string column, bool onTable)
        {
            int idx = TextLocator.IndexOfWord(text, column, 0);
            var map = new LineMap(text);
            DependencyResult r = Base(ctx, p, ConfidenceLevel.TextMatchReviewRequired);
            r.MatchType = "Text match in process " + source;
            r.MatchedValue = column;
            r.Evidence = map.SnippetAt(idx, column.Length);
            r.LineNumber = map.LineOf(idx);
            r.SourceResource = source;
            r.Location = source + " line " + r.LineNumber;
            r.AnalysisMethod = "Text search (definition could not be parsed as XAML)";
            r.Explanation = "The column name appears in the process definition text.";
            r.Limitation = "The definition format could not be parsed structurally, so this is a text match and may not be a real dependency." + (onTable ? string.Empty : " The process belongs to another table.");
            return r;
        }

        // ---------------------------------------------------------------- XAML entity context

        private static bool IsEntityNameContext(XmlValueMatch m)
        {
            if (m.Attribute != null && m.Attribute.Name.LocalName.Equals("EntityName", StringComparison.OrdinalIgnoreCase)) return true;
            XAttribute key = m.Element.Attribute(X + "Key");
            return key != null && key.Value.Equals("EntityName", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// The table a step's attribute belongs to: an EntityName attribute on the same element, or a sibling
        /// &lt;InArgument x:Key="EntityName"&gt;. Null when the step does not state one.
        /// </summary>
        public static string EntityContext(XmlValueMatch m)
        {
            XElement el = m.Element;
            XAttribute own = el.Attribute("EntityName");
            if (own != null && !(m.Attribute != null && m.Attribute.Name.LocalName == "EntityName")) return XmlPath.Normalize(own.Value);

            XElement container = el.Parent;
            for (int depth = 0; container != null && depth < 3; depth++, container = container.Parent)
            {
                foreach (XElement sibling in container.Elements())
                {
                    XAttribute key = sibling.Attribute(X + "Key");
                    if (key != null && key.Value.Equals("EntityName", StringComparison.OrdinalIgnoreCase) && !sibling.HasElements)
                        return XmlPath.Normalize(sibling.Value);
                }
                XAttribute viaAttr = container.Attribute("EntityName");
                if (viaAttr != null) return XmlPath.Normalize(viaAttr.Value);
            }
            return null;
        }

        private static string StepSuffix(XmlValueMatch m)
        {
            string step = XmlPath.NearestDisplayName(m.Element);
            return step == null ? string.Empty : "  [step: " + step + "]";
        }
    }
}
