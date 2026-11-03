// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DataverseLogicDependencyExplorer.Models;

namespace DataverseLogicDependencyExplorer.Reporting
{
    public class ExportOutcome
    {
        public string Path { get; set; }
        public int TruncatedCells { get; set; }
    }

    /// <summary>Builds the Impact Report workbook: Summary, Dependencies, Scan Diagnostics and Scan Configuration.</summary>
    public static class ExcelReportExporter
    {
        public const string ReportVersion = ProductInfo.Version;

        public static string DefaultFileName(DateTime now)
        {
            return "DataverseLogicDependencyReport_" + now.ToString("yyyyMMdd_HHmmss") + ".xlsx";
        }

        public static ExportOutcome Export(AnalysisResult analysis, string path, string pluginVersion)
        {
            if (analysis == null) throw new ArgumentNullException("analysis");
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A file path is required.", "path");

            XlsxWorkbookWriter wb = Build(analysis, pluginVersion);
            wb.Save(path);
            return new ExportOutcome { Path = path, TruncatedCells = wb.TruncatedCells };
        }

        public static XlsxWorkbookWriter Build(AnalysisResult a, string pluginVersion)
        {
            var wb = new XlsxWorkbookWriter();
            BuildDependencies(wb, a);   // first, so the truncation count is known when the summary is written
            XlsxSheet deps = wb.Sheets[0];
            BuildDiagnostics(wb, a);
            BuildConfiguration(wb, a, pluginVersion);
            BuildSummary(wb, a, pluginVersion);

            // Required order: Summary, Dependencies, Scan Diagnostics, Scan Configuration.
            var ordered = new XlsxWorkbookWriter();
            foreach (string name in new[] { "Summary", "Dependencies", "Scan Diagnostics", "Scan Configuration" })
            {
                XlsxSheet src = wb.Sheets.First(s => s.Name == name);
                XlsxSheet dst = ordered.AddSheet(name);
                dst.FreezeRows = src.FreezeRows;
                dst.AutoFilterHeaderRow = src.AutoFilterHeaderRow;
                foreach (var kv in src.ColumnWidths) dst.ColumnWidths[kv.Key] = kv.Value;
                foreach (var row in src.Rows) dst.AddRow(row.Key, row.Value);
            }
            return ordered;
        }

        private static string Join(IEnumerable<string> items) { return string.Join("\n", items.Where(s => !string.IsNullOrWhiteSpace(s))); }

        // ------------------------------------------------------------------ Dependencies

        private static void BuildDependencies(XlsxWorkbookWriter wb, AnalysisResult a)
        {
            XlsxSheet s = wb.AddSheet("Dependencies");
            s.FreezeRows = 1;
            s.AutoFilterHeaderRow = 1;
            s.AddRow(new object[]
            {
                "Component Type", "Component Name", "Component ID", "Component Unique Name", "Parent Table", "Component State", "Solution",
                "Target Table", "Target Column", "Match Type", "Matched Value", "Confidence", "Evidence", "Source Resource",
                "Location (XML node / JSON path / attribute)", "Line Number", "Analysis Method", "Why Included", "Limitation",
                "Additional Info", "Scan Time (UTC)", "Environment"
            }, XlsxStyle.Header);

            foreach (DependencyResult r in a.Results)
            {
                s.AddRow(new object[]
                {
                    r.ComponentType.Display(), r.ComponentName, r.ComponentId, r.ComponentUniqueName, r.ParentTable, r.ComponentState, r.SolutionName,
                    r.TargetTable, r.TargetColumn, r.MatchType, r.MatchedValue, r.Confidence.Display(), r.Evidence, r.SourceResource,
                    r.Location, r.LineNumber.HasValue ? (object)r.LineNumber.Value : null, r.AnalysisMethod, r.Explanation, r.Limitation,
                    r.AdditionalInfo, r.ScanTimestampUtc, r.Environment
                });
            }
            double[] widths = { 24, 34, 38, 28, 18, 22, 28, 18, 20, 36, 22, 28, 70, 36, 44, 10, 36, 60, 50, 40, 20, 40 };
            for (int i = 0; i < widths.Length; i++) s.ColumnWidths[i] = widths[i];
        }

        // ------------------------------------------------------------------ Scan Diagnostics

        private static void BuildDiagnostics(XlsxWorkbookWriter wb, AnalysisResult a)
        {
            XlsxSheet s = wb.AddSheet("Scan Diagnostics");
            s.FreezeRows = 1;
            s.AutoFilterHeaderRow = 1;
            s.AddRow(new object[]
            {
                "Component Category", "Scan Status", "Inspected?", "Records / Definitions Examined", "Matches", "Malformed Definitions",
                "Permission Limited", "Duration (ms)", "Errors", "Warnings", "Limitations (what this scan cannot see)"
            }, XlsxStyle.Header);

            foreach (ScanDiagnostic d in a.Diagnostics)
            {
                s.AddRow(new object[]
                {
                    d.Category.Display(), d.Status.Display(), d.WasInspected ? "Yes" : "No",
                    d.WasInspected ? (object)d.RecordsExamined : null, d.MatchCount, d.MalformedDefinitions,
                    d.PermissionLimited ? "Yes" : "No", d.ElapsedMilliseconds,
                    Join(d.Errors), Join(d.Warnings), Join(d.Limitations)
                });
            }
            foreach (string w in a.GeneralWarnings)
                s.AddRow(new object[] { "(run-wide)", "Warning", null, null, null, null, null, null, null, w, null });

            double[] widths = { 34, 24, 12, 18, 10, 14, 14, 14, 60, 70, 80 };
            for (int i = 0; i < widths.Length; i++) s.ColumnWidths[i] = widths[i];
        }

        // ------------------------------------------------------------------ Scan Configuration

        private static void BuildConfiguration(XlsxWorkbookWriter wb, AnalysisResult a, string pluginVersion)
        {
            ScanRequest q = a.Request;
            XlsxSheet s = wb.AddSheet("Scan Configuration");
            s.FreezeRows = 1;
            s.AutoFilterHeaderRow = 1;
            s.AddRow(new object[] { "Setting", "Value" }, XlsxStyle.Header);

            Action<string, object> add = (k, v) => s.AddRow(new object[] { k, v });
            add("Tool", ProductInfo.Name);
            add("Report version", ExcelReportExporter.ReportVersion);
            add("Plugin version", pluginVersion);
            add("Maintainer", ProductInfo.Maintainer);
            add("LinkedIn", ProductInfo.LinkedInUrl);
            add("Project URL", ProductInfo.RepositoryUrl);
            add("Host", ProductInfo.Host);
            add("Data platform", ProductInfo.DataPlatform);
            add("Export format", ProductInfo.ExportFormat);
            add("Scan timestamp (UTC)", a.StartedUtc);
            add("Search mode", q.Mode.Display());
            add("Target table (logical name)", q.TableLogicalName);
            add("Target table (display name)", q.TableDisplayName);
            add("Target table entity set name", q.TableEntitySetName);
            add("Target column (logical name)", q.Mode == SearchMode.Column ? q.ColumnLogicalName : null);
            add("Target column (display name)", q.Mode == SearchMode.Column ? q.ColumnDisplayName : null);
            add("Component categories selected", string.Join("\n", ScanRequest.AllScopes.Where(k => q.Scopes.Contains(k)).Select(k => k.Display())));
            add("Component categories NOT selected", string.Join("\n", ScanRequest.AllScopes.Where(k => !q.Scopes.Contains(k)).Select(k => k.Display())));
            add("JavaScript scope", q.JavaScriptScope == JavaScriptScope.AllScriptWebResources ? "All script web resources" : "Only libraries registered on forms of the selected table");
            add("Include comment-only JavaScript matches", q.IncludeCommentMatches);
            add("Cross-table (deep) process scan", q.DeepProcessScan);
            add("Environment name", q.EnvironmentName);
            add("Environment URL", q.EnvironmentUrl);
            add("Environment identifier", q.EnvironmentId);
            add("Data access", "Read-only. Late-bound Dataverse SDK queries with paging; no data was modified. No credentials or tokens are stored in this report.");
            add("Scan cancelled", a.WasCancelled);
            foreach (ConfidenceLevel c in Enum.GetValues(typeof(ConfidenceLevel)).Cast<ConfidenceLevel>())
                add("Confidence: " + c.Display(), c.Explain());
            s.ColumnWidths[0] = 44;
            s.ColumnWidths[1] = 100;
        }

        // ------------------------------------------------------------------ Summary

        private static void BuildSummary(XlsxWorkbookWriter wb, AnalysisResult a, string pluginVersion)
        {
            ScanRequest q = a.Request;
            XlsxSheet s = wb.AddSheet("Summary");
            s.ColumnWidths[0] = 46;
            s.ColumnWidths[1] = 36;
            s.ColumnWidths[2] = 100;

            s.AddRow(new object[] { "Dataverse Logic Dependency Impact Report" }, XlsxStyle.Title);
            s.AddRow(new object[] { "Read-only analysis. Findings show what the plugin was able to inspect; they are not a guarantee that every dependency was found." });
            s.AddRow(new object[] { ProductInfo.Name + " v" + pluginVersion + "  |  Maintainer: " + ProductInfo.Maintainer + "  |  LinkedIn: " + ProductInfo.LinkedInUrl });
            s.AddRow(new object[] { "Host: " + ProductInfo.Host + "  |  Data platform: " + ProductInfo.DataPlatform + "  |  Export format: " + ProductInfo.ExportFormat });
            s.AddRow();

            Action<string, object> kv = (k, v) => s.AddRow(new object[] { k, v });

            kv("Environment", !string.IsNullOrEmpty(q.EnvironmentUrl) ? q.EnvironmentUrl : q.EnvironmentName);
            kv("Environment identifier", string.IsNullOrEmpty(q.EnvironmentId) ? "(not available)" : q.EnvironmentId);
            kv("Scan date and time (UTC)", a.StartedUtc);
            kv("Scan duration (seconds)", Math.Round((a.FinishedUtc - a.StartedUtc).TotalSeconds, 1));
            kv("Search mode", q.Mode.Display());
            kv("Selected table", q.TableDisplayName == null ? q.TableLogicalName : q.TableDisplayName + " (" + q.TableLogicalName + ")");
            kv("Selected column", q.Mode == SearchMode.Column ? (q.ColumnDisplayName == null ? q.ColumnLogicalName : q.ColumnDisplayName + " (" + q.ColumnLogicalName + ")") : "(table search)");
            kv("Component categories selected", string.Join(", ", ScanRequest.AllScopes.Where(k => q.Scopes.Contains(k)).Select(k => k.Display())));
            kv("Total matches", a.MatchCount);
            kv("Unique components", a.UniqueComponents);
            kv("Components / definitions examined", a.ComponentsExamined);
            kv("Scans completed (with or without warnings)", a.CompletedScans);
            kv("Scans failed (incl. permission denied)", a.FailedScans);
            kv("Scans partially supported / not inspected / cancelled", a.UnsupportedOrPartialScans);
            kv("Report version / plugin version", ExcelReportExporter.ReportVersion + " / " + pluginVersion);
            kv("Overall result", OverallText(a));
            s.AddRow();

            s.AddRow(new object[] { "Matches by component type", "Matches" }, XlsxStyle.Header);
            foreach (ComponentType t in Enum.GetValues(typeof(ComponentType)).Cast<ComponentType>())
                s.AddRow(new object[] { t.Display(), a.Results.Count(r => r.ComponentType == t) });
            s.AddRow();

            s.AddRow(new object[] { "Matches by confidence classification", "Matches", "Meaning" }, XlsxStyle.Header);
            foreach (ConfidenceLevel c in Enum.GetValues(typeof(ConfidenceLevel)).Cast<ConfidenceLevel>())
                s.AddRow(new object[] { c.Display(), a.Results.Count(r => r.Confidence == c), c.Explain() });
            s.AddRow();

            s.AddRow(new object[] { "Scan status by category", "Status", "Matches / notes" }, XlsxStyle.Header);
            foreach (ScanDiagnostic d in a.Diagnostics)
                s.AddRow(new object[] { d.Category.Display(), d.Status.Display(), d.WasInspected ? d.MatchCount + " match(es)" : "Not inspected - zero matches would be meaningless" });
            s.AddRow();

            s.AddRow(new object[] { "Important limitations and warnings" }, XlsxStyle.Header);
            var notes = new List<string>
            {
                "A scan with no matches is not proof that no dependency exists: only the categories and definitions listed as inspected were examined.",
                "Compiled plugin code is never inspected; plugin findings come from registration metadata only.",
                "JavaScript and flow-expression matches are static text matches and must be reviewed; dynamic or computed names cannot be resolved.",
                "Cloud flows are read only when stored in Dataverse; the Power Automate Management API is not used."
            };
            notes.AddRange(a.GeneralWarnings);
            foreach (ScanDiagnostic d in a.Diagnostics.Where(d => d.Status != ScanStatus.Completed))
            {
                foreach (string e in d.Errors) notes.Add(d.Category.Display() + " [" + d.Status.Display() + "]: " + e);
                foreach (string w in d.Warnings) notes.Add(d.Category.Display() + ": " + w);
            }
            foreach (string n in notes) s.AddRow(new object[] { n });
            if (wb.TruncatedCells > 0)
                s.AddRow(new object[] { wb.TruncatedCells + " cell(s) exceeded Excel's 32,767 character limit and were cut with a visible [TRUNCATED] marker." });
        }

        public static string OverallText(AnalysisResult a)
        {
            if (a.WasCancelled) return "CANCELLED - results are incomplete.";
            if (a.Results.Count > 0)
                return a.Results.Count + " finding(s)." + (a.Diagnostics.Any(d => d.Status != ScanStatus.Completed && d.Status != ScanStatus.CompletedWithWarnings && a.Request.Scopes.Contains(d.Category))
                    ? " Some selected categories were not fully inspected - see Scan Diagnostics." : string.Empty);
            return a.IsCleanEmptyResult
                ? "No references were found. Every selected category was inspected."
                : "No matches were found, BUT not every selected category was fully inspected (failed, permission denied or partially supported) - see Scan Diagnostics. This is not a clean result.";
        }
    }
}
