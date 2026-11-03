// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using DataverseLogicDependencyExplorer.Models;

namespace DataverseLogicDependencyExplorer.Services.Scanners
{
    /// <summary>
    /// Forms of the selected table: field controls, registered libraries and event handlers (from the parsed form XML) and,
    /// for column searches, whether a registered handler's source actually references the column.
    /// </summary>
    public class FormScanner : ScannerBase
    {
        public override ScannerKind Kind { get { return ScannerKind.Forms; } }

        protected override void Execute(ScanContext ctx, ScanOutcome outcome, CancellationToken ct)
        {
            ScanRequest req = ctx.Request;
            ScanDiagnostic diag = outcome.Diagnostic;
            diag.Limit("Only forms of '" + req.TableLogicalName + "' are inspected. Forms of other tables that embed it (subgrids, quick view forms) are not searched.");
            diag.Limit("Business rules associated with forms are not exposed in form XML; see the Business Rules scanner (rules are read per table).");

            ctx.Report("Retrieving forms of " + req.TableLogicalName + "...");
            List<FormRecord> forms = ctx.Forms.GetForms(req.TableLogicalName, ct);
            foreach (string w in ctx.Forms.Warnings.Distinct()) diag.Warn(w);
            diag.RecordsExamined = forms.Count;

            // Libraries are needed to check whether handler code references the column.
            var analyses = new Dictionary<string, JsAnalysis>(StringComparer.OrdinalIgnoreCase);
            if (req.Mode == SearchMode.Column)
            {
                var libs = forms.Where(f => f.Definition.IsValid).SelectMany(f => f.Definition.Handlers.Where(h => h.Enabled)).Select(h => h.LibraryName).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                var warnings = new List<string>();
                ctx.WebResources.EnsureNames(libs, ct, warnings);
                foreach (string w in warnings.Distinct()) diag.Warn(w);
                foreach (string lib in libs)
                {
                    WebResourceContent wr = ctx.WebResources.TryGet(lib);
                    if (wr == null) { diag.Warn("Library '" + lib + "' is registered on a form but was not found as a script web resource; its handlers could not be checked."); continue; }
                    if (wr.Text == null) continue;
                    analyses[lib] = JsSourceAnalyzer.FindColumn(wr.Text, req.ColumnLogicalName, false);
                }
            }

            foreach (FormRecord form in forms)
            {
                ct.ThrowIfCancellationRequested();
                FormDefinition def = form.Definition;
                if (!def.IsValid)
                {
                    diag.MalformedDefinitions++;
                    diag.Warn("Form '" + form.Name + "' could not be analysed: " + def.Error);
                    continue;
                }
                if (req.Mode == SearchMode.Table) TableMode(ctx, outcome, form, def);
                else ColumnMode(ctx, outcome, form, def, analyses);
            }
        }

        private DependencyResult Base(ScanContext ctx, FormRecord form, ComponentType type, ConfidenceLevel confidence)
        {
            DependencyResult r = ctx.NewResult(ScannerKind.Forms, type, confidence);
            r.ComponentName = form.Name;
            r.ComponentId = form.Id.ToString();
            r.ParentTable = form.Table;
            r.ComponentState = form.StateText;
            r.SourceResource = "form XML of '" + form.Name + "' (" + FormXmlReader.DescribeFormType(form.Type) + ")";
            r.AnalysisMethod = "Parsed form XML";
            return r;
        }

        // ---------------------------------------------------------------- table mode

        private void TableMode(ScanContext ctx, ScanOutcome outcome, FormRecord form, FormDefinition def)
        {
            string table = ctx.Request.TableLogicalName;

            DependencyResult f = Base(ctx, form, ComponentType.Form, ConfidenceLevel.ConfirmedMetadataReference);
            f.MatchType = FormXmlReader.DescribeFormType(form.Type) + " form of the table";
            f.MatchedValue = table;
            f.Evidence = "systemform.objecttypecode = '" + table + "'; " + def.Libraries.Count + " librar" + (def.Libraries.Count == 1 ? "y" : "ies") + ", " + def.Handlers.Count + " event handler(s), " + def.Controls.Count + " field control(s)";
            f.Location = "systemform.objecttypecode";
            f.AnalysisMethod = "Dataverse form metadata";
            f.Explanation = "The form belongs to the selected table.";
            outcome.Results.Add(f);

            foreach (string lib in def.Libraries.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                DependencyResult r = Base(ctx, form, ComponentType.FormEventRegistration, ConfidenceLevel.StrongDefinitionReference);
                r.MatchType = "JavaScript library registered on form";
                r.MatchedValue = table;
                r.Evidence = "formLibraries/Library name=\"" + lib + "\"";
                r.Location = "/form/formLibraries/Library[@name='" + lib + "']";
                r.ComponentUniqueName = lib;
                r.Explanation = "The library is loaded by a form of the selected table.";
                outcome.Results.Add(r);
            }

            foreach (FormHandler h in def.Handlers)
            {
                DependencyResult r = Base(ctx, form, ComponentType.FormEventRegistration, ConfidenceLevel.StrongDefinitionReference);
                r.MatchType = DescribeEvent(h);
                r.MatchedValue = table;
                r.ComponentUniqueName = h.FunctionName;
                r.Evidence = HandlerText(h);
                r.Location = h.Location;
                r.Explanation = "An event handler is registered on a form of the selected table.";
                if (!h.Enabled) r.ComponentState = (r.ComponentState == null ? string.Empty : r.ComponentState + "; ") + "Handler disabled";
                outcome.Results.Add(r);
            }

            foreach (XmlValueMatch m in FormXmlReader.FindOtherReferences(def, table).Take(20))
            {
                DependencyResult r = Base(ctx, form, ComponentType.Form, ConfidenceLevel.StrongDefinitionReference);
                r.MatchType = "General reference to the table in form XML";
                r.MatchedValue = table;
                r.Evidence = XmlPath.Render(m);
                r.Location = m.Path;
                r.Explanation = "The form XML names the selected table (for example a subgrid or quick view target).";
                outcome.Results.Add(r);
            }
        }

        // ---------------------------------------------------------------- column mode

        private void ColumnMode(ScanContext ctx, ScanOutcome outcome, FormRecord form, FormDefinition def, Dictionary<string, JsAnalysis> analyses)
        {
            string column = ctx.Request.ColumnLogicalName;

            foreach (FormControl c in def.Controls.Where(c => string.Equals(c.DataFieldName, column, StringComparison.OrdinalIgnoreCase)))
            {
                DependencyResult r = Base(ctx, form, ComponentType.Form, ConfidenceLevel.StrongDefinitionReference);
                r.MatchType = "Field control on form";
                r.MatchedValue = column;
                r.Evidence = "<control id=\"" + c.Id + "\" datafieldname=\"" + c.DataFieldName + "\" />";
                r.Location = c.Location;
                r.Explanation = "The form displays column '" + column + "' in a field control.";
                outcome.Results.Add(r);
            }

            foreach (FormHandler h in def.Handlers.Where(h => string.Equals(h.Attribute, column, StringComparison.OrdinalIgnoreCase)))
            {
                DependencyResult r = Base(ctx, form, ComponentType.FormEventRegistration, ConfidenceLevel.StrongDefinitionReference);
                r.MatchType = "Event registered against the column (" + (string.IsNullOrEmpty(h.EventName) ? "event" : h.EventName) + ")";
                r.MatchedValue = column;
                r.ComponentUniqueName = h.FunctionName;
                r.Evidence = HandlerText(h);
                r.Location = h.Location;
                r.Explanation = "A handler is registered for the " + h.EventName + " event of column '" + column + "' (event element attribute=\"" + column + "\").";
                if (!h.Enabled) r.ComponentState = (r.ComponentState == null ? string.Empty : r.ComponentState + "; ") + "Handler disabled";
                outcome.Results.Add(r);
            }

            // Registered handlers whose code references the column (distinct from the registration being on the column).
            var reported = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (FormHandler h in def.Handlers.Where(h => h.Enabled))
            {
                JsAnalysis analysis;
                if (string.IsNullOrEmpty(h.LibraryName) || !analyses.TryGetValue(h.LibraryName, out analysis) || analysis.Matches.Count == 0) continue;
                if (!reported.Add(h.LibraryName + "::" + h.FunctionName)) continue;

                WebResourceContent wr = ctx.WebResources.TryGet(h.LibraryName);
                int start, end;
                bool located = wr != null && JsSourceAnalyzer.TryFindFunctionBody(wr.Text, h.FunctionName, out start, out end);
                List<JsMatch> inside;
                if (located)
                {
                    JsSourceAnalyzer.TryFindFunctionBody(wr.Text, h.FunctionName, out start, out end);
                    int s = start, e = end;
                    inside = analysis.Matches.Where(m => m.Index >= s && m.Index <= e).ToList();
                    if (inside.Count == 0) continue;
                }
                else inside = analysis.Matches;

                JsMatch first = inside[0];
                DependencyResult r = Base(ctx, form, ComponentType.FormEventRegistration, ConfidenceLevel.TextMatchReviewRequired);
                r.MatchType = located ? "Handler function source references the column" : "Registered library references the column (function body not located)";
                r.MatchedValue = column;
                r.ComponentUniqueName = h.FunctionName;
                r.LineNumber = first.Line;
                r.Location = h.Location;
                r.Evidence = HandlerText(h) + "\nSource line " + first.Line + " of " + h.LibraryName + ": " + first.Snippet + (inside.Count > 1 ? "\n(" + (inside.Count - 1) + " more match(es))" : string.Empty);
                r.AnalysisMethod = "Parsed form XML + static JavaScript text analysis";
                r.Explanation = located
                    ? "The handler registered on this form calls a function whose source contains '" + column + "'."
                    : "The handler's library contains '" + column + "', but the function body could not be isolated, so the match may belong to another function in the same library.";
                r.Limitation = "Static text match; runtime behaviour is not verified.";
                outcome.Results.Add(r);
            }

            foreach (XmlValueMatch m in FormXmlReader.FindOtherReferences(def, column).Take(20))
            {
                DependencyResult r = Base(ctx, form, ComponentType.Form, ConfidenceLevel.StrongDefinitionReference);
                r.MatchType = "General reference to the column in form XML";
                r.MatchedValue = column;
                r.Evidence = XmlPath.Render(m);
                r.Location = m.Path;
                r.Explanation = "The form XML contains the column name outside a field control or event registration.";
                outcome.Results.Add(r);
            }
        }

        private static string DescribeEvent(FormHandler h)
        {
            string e = (h.EventName ?? string.Empty).ToLowerInvariant();
            string name = e == "onload" ? "OnLoad handler" : e == "onsave" ? "OnSave handler" : e == "onchange" ? "OnChange handler" : (string.IsNullOrEmpty(h.EventName) ? "Event" : h.EventName) + " handler";
            return name;
        }

        private static string HandlerText(FormHandler h)
        {
            return (h.EventName ?? "event") + (h.Attribute != null ? " of '" + h.Attribute + "'" : string.Empty)
                + (h.ControlId != null && h.Attribute == null ? " on control '" + h.ControlId + "'" : string.Empty)
                + " -> " + h.LibraryName + " :: " + h.FunctionName
                + " (" + (h.Enabled ? "enabled" : "disabled") + (h.PassExecutionContext ? ", passes execution context" : string.Empty) + ")"
                + (string.IsNullOrEmpty(h.Parameters) ? string.Empty : "; parameters: " + DataverseHelper.Truncate(h.Parameters, 200));
        }
    }
}
