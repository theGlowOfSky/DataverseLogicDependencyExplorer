// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Linq;
using System.Text;
using DataverseLogicDependencyExplorer.Models;

namespace DataverseLogicDependencyExplorer.Services
{
    /// <summary>Builds the evidence text shown in the Impact Report detail panel (kept out of the UI so it can be tested).</summary>
    public static class ResultDetailFormatter
    {
        public static string Format(DependencyResult r, ScanDiagnostic categoryDiagnostic)
        {
            if (r == null) return string.Empty;
            var sb = new StringBuilder();
            Action<string, string> line = (label, value) =>
            {
                if (!string.IsNullOrWhiteSpace(value)) sb.Append(label).Append(": ").AppendLine(value);
            };

            line("Component", r.ComponentName);
            line("Component type", r.ComponentType.Display());
            line("Component ID", r.ComponentId);
            line("Logical / unique name", r.ComponentUniqueName);
            line("Parent table", r.ParentTable);
            line("State", r.ComponentState);
            line("Solution", r.SolutionName);
            sb.AppendLine();
            line("Searched table", r.TargetTable);
            line("Matched column", r.TargetColumn);
            line("Matched value", r.MatchedValue);
            line("Match type", r.MatchType);
            line("Reference source", r.SourceResource);
            line("File / resource name", r.ComponentType == ComponentType.JavaScriptWebResource ? r.SourceResource : null);
            line("Location (XML node / JSON path / attribute)", r.Location);
            if (r.LineNumber.HasValue) line("Line number", r.LineNumber.Value.ToString());
            line("Analysis method", r.AnalysisMethod);
            sb.AppendLine();
            line("Confidence", r.Confidence.Display());
            line("Meaning of this confidence", r.Confidence.Explain());
            line("Why this result was included", r.Explanation);
            line("Technical limitation of this finding", r.Limitation);
            line("Additional information", r.AdditionalInfo);

            if (!string.IsNullOrWhiteSpace(r.Evidence))
            {
                sb.AppendLine();
                sb.AppendLine("Evidence:");
                sb.AppendLine(r.Evidence);
            }

            if (categoryDiagnostic != null)
            {
                sb.AppendLine();
                line("Scan status of this category", categoryDiagnostic.Status.Display());
                if (categoryDiagnostic.Limitations.Count > 0)
                {
                    sb.AppendLine("Limitations of this category:");
                    foreach (string l in categoryDiagnostic.Limitations) sb.Append("  - ").AppendLine(l);
                }
                if (categoryDiagnostic.Warnings.Count > 0)
                {
                    sb.AppendLine("Warnings during this category:");
                    foreach (string w in categoryDiagnostic.Warnings.Take(10)) sb.Append("  - ").AppendLine(w);
                }
            }
            line("Scan time (UTC)", r.ScanTimestampUtc == default(DateTime) ? null : r.ScanTimestampUtc.ToString("yyyy-MM-dd HH:mm:ss"));
            line("Environment", r.Environment);
            return sb.ToString().Replace("\r\n", "\n").Replace("\n", Environment.NewLine);
        }
    }

    public static class AnalysisSummary
    {
        /// <summary>One-line summary: components examined, matches found and scan failures.</summary>
        public static string Line(AnalysisResult a)
        {
            if (a == null) return "No scan has been run.";
            string s = "Examined " + a.ComponentsExamined + " component(s)  |  Matches found: " + a.MatchCount + "  |  Scan failures: " + a.FailedScans;
            int partial = a.UnsupportedOrPartialScans;
            if (partial > 0) s += "  |  Partial / not inspected: " + partial;
            if (a.WasCancelled) s += "  |  CANCELLED";
            return s;
        }
    }
}
