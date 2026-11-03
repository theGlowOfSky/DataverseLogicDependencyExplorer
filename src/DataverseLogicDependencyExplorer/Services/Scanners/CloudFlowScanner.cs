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
    /// Cloud flows that are stored in Dataverse (workflow records, category 5) are read with the existing connection. No Power
    /// Automate Management API call and no separate authentication is used, so flows that exist only in the flow service are
    /// invisible. The scanner therefore reports "Partially supported" even when it succeeds.
    /// </summary>
    public class CloudFlowScanner : ScannerBase
    {
        public override ScannerKind Kind { get { return ScannerKind.CloudFlows; } }

        protected override void Execute(ScanContext ctx, ScanOutcome outcome, CancellationToken ct)
        {
            ScanRequest req = ctx.Request;
            ScanDiagnostic diag = outcome.Diagnostic;
            diag.Status = ScanStatus.PartiallySupported;
            diag.Limit("Only cloud flows stored in Dataverse (solution-aware flows) are visible through this connection. Flows that exist only in the Power Automate service require the Power Automate Management API, which this plugin does not call (no separate sign-in is requested).");
            diag.Limit("A text match in a flow is not proof of a runtime dependency: flows may build table/column names dynamically, use environment variables, call child flows or use custom connectors.");

            ctx.Report("Retrieving cloud flows stored in Dataverse...");
            var q = new QueryExpression("workflow")
            {
                ColumnSet = new ColumnSet("workflowid", "name", "uniquename", "statecode", "statuscode", "clientdata", "primaryentity", "ismanaged")
            };
            q.Criteria.AddCondition("category", ConditionOperator.Equal, 5);
            q.Criteria.AddCondition("type", ConditionOperator.Equal, 1);
            q.AddOrder("name", OrderType.Ascending);

            List<Entity> flows = DataverseHelper.RetrieveAll(ctx.Service, q, ct, 25).ToList();
            diag.RecordsExamined = flows.Count;
            if (flows.Count == 0)
                diag.Warn("No cloud flow definitions are stored in Dataverse for this user/environment. This does NOT mean no flow uses the table or column.");

            int malformed = 0, unresolved = 0, noDefinition = 0;
            foreach (Entity flow in flows)
            {
                ct.ThrowIfCancellationRequested();
                string name = DataverseHelper.Str(flow, "name") ?? "(unnamed flow)";
                string json = DataverseHelper.Str(flow, "clientdata");
                if (string.IsNullOrWhiteSpace(json)) { noDefinition++; continue; }

                FlowAnalysis a = CloudFlowDefinition.Analyze(json, req.TableLogicalName, req.TableEntitySetName, req.Mode == SearchMode.Column ? req.ColumnLogicalName : null);
                if (a.Error != null)
                {
                    malformed++;
                    diag.MalformedDefinitions++;
                    diag.Warn("Flow '" + name + "': " + a.Error);
                    continue;
                }
                unresolved += a.UnresolvedTableSteps;

                IEnumerable<FlowReference> refs = req.Mode == SearchMode.Table ? a.TableReferences : a.ColumnReferences;
                foreach (FlowReference fr in refs.Take(25))
                    outcome.Results.Add(Build(ctx, flow, name, fr, a));
            }

            if (noDefinition > 0) diag.Warn(noDefinition + " flow(s) have no readable definition (clientdata empty) and could not be searched.");
            if (unresolved > 0)
                diag.Warn(unresolved + " Dataverse step(s) use an expression or omit the table name, so their target table cannot be resolved statically.");
        }

        private DependencyResult Build(ScanContext ctx, Entity flow, string name, FlowReference fr, FlowAnalysis a)
        {
            ScanRequest req = ctx.Request;
            ConfidenceLevel conf;
            if (!fr.Structured) conf = ConfidenceLevel.TextMatchReviewRequired;
            else if (fr.TableUnresolved) conf = ConfidenceLevel.PossibleOrDynamic;
            else conf = ConfidenceLevel.StrongDefinitionReference;

            DependencyResult r = ctx.NewResult(ScannerKind.CloudFlows, ComponentType.CloudFlow, conf);
            r.ComponentName = name;
            r.ComponentId = flow.Id.ToString();
            r.ComponentUniqueName = DataverseHelper.Str(flow, "uniquename");
            r.ParentTable = DataverseHelper.Str(flow, "primaryentity");
            r.ComponentState = BusinessRuleScanner.StateText(flow);
            r.MatchType = (fr.IsTrigger ? "Trigger: " : fr.ActionName != null ? "Action: " : string.Empty) + fr.Role;
            r.MatchedValue = req.Mode == SearchMode.Column ? req.ColumnLogicalName : req.TableLogicalName;
            r.Evidence = (fr.ActionName == null ? string.Empty : "Step '" + fr.ActionName + "': ") + fr.Expression;
            r.SourceResource = "workflow.clientdata (flow definition JSON)";
            r.Location = fr.Path;
            r.AnalysisMethod = fr.Structured ? "Parsed flow JSON (Dataverse connector parameters)" : "Expression text search in flow JSON";
            r.AdditionalInfo = a.ConnectionReferences.Count == 0 ? null : "Connection references: " + string.Join(", ", a.ConnectionReferences);

            if (!fr.Structured)
            {
                r.Explanation = "A dynamic-content expression in the flow names the column. The table it belongs to is not established.";
                r.Limitation = "Expression text match only; the dynamic content may come from a different table.";
            }
            else if (fr.TableUnresolved)
            {
                r.Explanation = "A Dataverse step names the column, but its table is given by an expression, so it may or may not be '" + req.TableLogicalName + "'.";
                r.Limitation = "The step's table cannot be resolved statically.";
            }
            else
            {
                r.Explanation = req.Mode == SearchMode.Table
                    ? "A Dataverse trigger or action in the flow targets the selected table."
                    : "A Dataverse step that targets '" + req.TableLogicalName + "' names the column.";
                r.Limitation = "Static definition analysis; environment variables, child flows and custom connectors are not resolved.";
            }
            return r;
        }
    }
}
