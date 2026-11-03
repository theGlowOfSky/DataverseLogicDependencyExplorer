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
    /// Business rules are workflow records with category 2 (definition rows, type 1) scoped to a primary table. Their generated
    /// client code (clientdata) and XAML are inspected; classic workflows, actions and business process flows are not mixed in.
    /// </summary>
    public class BusinessRuleScanner : ScannerBase
    {
        public override ScannerKind Kind { get { return ScannerKind.BusinessRules; } }

        protected override void Execute(ScanContext ctx, ScanOutcome outcome, CancellationToken ct)
        {
            ScanRequest req = ctx.Request;
            ScanDiagnostic diag = outcome.Diagnostic;
            diag.Limit("Only business rules whose primary table is '" + req.TableLogicalName + "' exist for this table; they are read from the workflow table (category 2, definition rows).");
            diag.Limit("Conditions and actions are taken from the rule's generated client code (clientdata) and, when present, its XAML. Form/scope association of a rule is not exposed in a form definition and is not reported.");

            ctx.Report("Retrieving business rules of " + req.TableLogicalName + "...");
            List<Entity> rules = Retrieve(ctx, ct, diag);
            diag.RecordsExamined = rules.Count;

            int unreadable = 0;
            foreach (Entity rule in rules)
            {
                ct.ThrowIfCancellationRequested();
                string name = DataverseHelper.Str(rule, "name") ?? "(unnamed)";
                string clientData = DataverseHelper.Str(rule, "clientdata");
                string xaml = DataverseHelper.Str(rule, "xaml");
                string js = BusinessRuleDefinition.ExtractClientCode(clientData);
                if (string.IsNullOrWhiteSpace(js) && string.IsNullOrWhiteSpace(xaml)) unreadable++;

                if (req.Mode == SearchMode.Table)
                    outcome.Results.Add(BuildTableResult(ctx, rule, name, js));
                else
                    ScanColumn(ctx, outcome, rule, name, js, xaml);
            }

            if (unreadable > 0)
                diag.Warn(unreadable + " business rule(s) have no readable definition (no client code or XAML); they could not be searched for column references.");
            if (req.Mode == SearchMode.Column && rules.Count == 0)
                diag.Limit("No business rules exist for this table, so there is nothing to reference the column.");
        }

        private List<Entity> Retrieve(ScanContext ctx, CancellationToken ct, ScanDiagnostic diag)
        {
            Func<bool, QueryExpression> build = withXaml =>
            {
                var cols = new List<string> { "workflowid", "name", "uniquename", "primaryentity", "statecode", "statuscode", "clientdata", "category", "type", "ismanaged" };
                if (withXaml) cols.Add("xaml");
                var q = new QueryExpression("workflow") { ColumnSet = new ColumnSet(cols.ToArray()) };
                q.Criteria.AddCondition("category", ConditionOperator.Equal, 2);
                q.Criteria.AddCondition("type", ConditionOperator.Equal, 1);
                q.Criteria.AddCondition("primaryentity", ConditionOperator.Equal, ctx.Request.TableLogicalName);
                q.AddOrder("name", OrderType.Ascending);
                return q;
            };

            try
            {
                return DataverseHelper.RetrieveAll(ctx.Service, build(true), ct, 50).ToList();
            }
            catch (FaultException<OrganizationServiceFault> ex) when (DataverseHelper.IsMissingMetadata(ex))
            {
                diag.Warn("The 'xaml' column could not be read in this environment; only the generated client code was inspected.");
                return DataverseHelper.RetrieveAll(ctx.Service, build(false), ct, 50).ToList();
            }
        }

        private DependencyResult BuildTableResult(ScanContext ctx, Entity rule, string name, string js)
        {
            DependencyResult r = Base(ctx, rule, name, ConfidenceLevel.ConfirmedMetadataReference);
            r.MatchType = "Business rule primary table";
            r.MatchedValue = ctx.Request.TableLogicalName;
            r.Evidence = "workflow.primaryentity = '" + DataverseHelper.Str(rule, "primaryentity") + "' (category 2: business rule)";
            r.SourceResource = "workflow record";
            r.Location = "workflow.primaryentity";
            r.AnalysisMethod = "Dataverse process metadata (workflow table)";
            r.Explanation = "The business rule is defined on the selected table.";
            List<string> cols = BusinessRuleDefinition.ReferencedColumns(js);
            r.AdditionalInfo = cols.Count == 0 ? null : "Columns used by the rule: " + DataverseHelper.Truncate(string.Join(", ", cols), 600);
            return r;
        }

        private void ScanColumn(ScanContext ctx, ScanOutcome outcome, Entity rule, string name, string js, string xaml)
        {
            string column = ctx.Request.ColumnLogicalName;
            List<BusinessRuleUsage> usages = BusinessRuleDefinition.FindUsage(js, column);

            var xamlMatches = new List<XmlValueMatch>();
            bool xamlTextFallback = false;
            if (!string.IsNullOrWhiteSpace(xaml))
            {
                try { xamlMatches = XmlPath.FindExactValues(XDocument.Parse(xaml), column); }
                catch (XmlException ex)
                {
                    outcome.Diagnostic.MalformedDefinitions++;
                    outcome.Diagnostic.Warn("The XAML of business rule '" + name + "' is malformed (" + ex.Message + "); a text search was used instead.");
                    xamlTextFallback = TextLocator.ContainsWord(xaml, column);
                }
            }

            if (usages.Count == 0 && xamlMatches.Count == 0 && !xamlTextFallback) return;

            bool structured = usages.Count > 0 || xamlMatches.Count > 0;
            DependencyResult r = Base(ctx, rule, name, structured ? ConfidenceLevel.StrongDefinitionReference : ConfidenceLevel.TextMatchReviewRequired);
            var methods = new List<string>();
            var evidence = new List<string>();

            if (usages.Count > 0)
            {
                r.MatchType = BusinessRuleDefinition.DescribeRoles(usages);
                r.LineNumber = usages[0].Line;
                r.SourceResource = "clientdata (generated client code)";
                r.Location = "clientcode line " + usages[0].Line;
                evidence.AddRange(usages.Take(6).Select(u => u.Role + " (line " + u.Line + "): " + DataverseHelper.Truncate(u.Text, 300)));
                if (usages.Count > 6) evidence.Add("... " + (usages.Count - 6) + " more line(s) in the generated client code.");
                methods.Add("Parsed generated client code");
            }
            if (xamlMatches.Count > 0)
            {
                if (usages.Count == 0)
                {
                    r.MatchType = "XAML reference";
                    r.SourceResource = "xaml";
                    r.Location = xamlMatches[0].Path;
                }
                evidence.AddRange(xamlMatches.Take(4).Select(m => "XAML " + m.Path + ": " + XmlPath.Render(m, 250)));
                methods.Add("Parsed XAML");
            }
            if (xamlTextFallback)
            {
                if (!structured) { r.MatchType = "Text match in malformed XAML"; r.SourceResource = "xaml"; }
                evidence.Add("The column name appears in the (malformed) XAML text.");
                methods.Add("Text search of malformed XAML");
                r.Limitation = "The XAML could not be parsed, so this is a text match only.";
            }

            r.MatchedValue = column;
            r.Evidence = string.Join("\n", evidence);
            r.AnalysisMethod = string.Join(" + ", methods);
            r.Explanation = "The business rule's definition references column '" + column + "'.";
            if (r.Limitation == null && usages.Count > 0)
                r.Limitation = "Line numbers refer to the generated client code extracted from clientdata, not to the raw record.";
            outcome.Results.Add(r);
        }

        private DependencyResult Base(ScanContext ctx, Entity rule, string name, ConfidenceLevel confidence)
        {
            DependencyResult r = ctx.NewResult(ScannerKind.BusinessRules, ComponentType.BusinessRule, confidence);
            r.ComponentName = name;
            r.ComponentId = rule.Id.ToString();
            r.ComponentUniqueName = DataverseHelper.Str(rule, "uniquename");
            r.ParentTable = DataverseHelper.Str(rule, "primaryentity");
            r.ComponentState = StateText(rule);
            return r;
        }

        internal static string StateText(Entity rule)
        {
            int? state = DataverseHelper.Opt(rule, "statecode");
            bool? managed = DataverseHelper.Bool(rule, "ismanaged");
            var parts = new List<string>();
            if (state.HasValue) parts.Add(state.Value == 1 ? "Activated" : state.Value == 0 ? "Draft" : "State " + state.Value);
            if (managed.HasValue) parts.Add(managed.Value ? "Managed" : "Unmanaged");
            return parts.Count == 0 ? null : string.Join(", ", parts);
        }
    }
}
