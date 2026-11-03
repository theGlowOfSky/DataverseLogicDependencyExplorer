// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using DataverseLogicDependencyExplorer.Models;

namespace DataverseLogicDependencyExplorer.ViewModels
{
    /// <summary>One line of the Impact Report grid (display values only).</summary>
    public class ResultRow
    {
        public DependencyResult Result { get; set; }
        public string ComponentType { get; set; }
        public string ComponentName { get; set; }
        public string ParentTable { get; set; }
        public string MatchedValue { get; set; }
        public string MatchType { get; set; }
        public string Confidence { get; set; }
        public int ConfidenceOrder { get; set; }
        public string ScanStatus { get; set; }
        public string Solution { get; set; }
        public string Location { get; set; }
        /// <summary>Full (possibly multi-line) text of the collapsed cells, keyed by property name; shown in the tooltip.</summary>
        public System.Collections.Generic.Dictionary<string, string> Full { get; } = new System.Collections.Generic.Dictionary<string, string>();

        public static ResultRow From(DependencyResult r, ScanDiagnostic d)
        {
            var row = new ResultRow
            {
                Result = r,
                ComponentType = r.ComponentType.Display(),
                ComponentName = UiLayout.Collapse(r.ComponentName),
                ParentTable = r.ParentTable,
                MatchedValue = UiLayout.Collapse(r.TargetColumn ?? r.MatchedValue),
                MatchType = UiLayout.Collapse(r.MatchType),
                Confidence = r.Confidence.Display(),
                ConfidenceOrder = (int)r.Confidence,
                ScanStatus = d == null ? string.Empty : d.Status.Display(),
                Solution = r.SolutionName,
                Location = UiLayout.Collapse(r.LineNumber.HasValue ? "Line " + r.LineNumber : r.Location)
            };
            Remember(row, "ComponentName", r.ComponentName);
            Remember(row, "MatchedValue", r.TargetColumn ?? r.MatchedValue);
            Remember(row, "MatchType", r.MatchType);
            Remember(row, "Location", r.LineNumber.HasValue ? "Line " + r.LineNumber + (string.IsNullOrEmpty(r.Location) ? string.Empty : "\n" + r.Location) : r.Location);
            return row;
        }

        private static void Remember(ResultRow row, string property, string full)
        {
            if (!string.IsNullOrEmpty(full) && (full.IndexOf('\n') >= 0 || full.IndexOf('\r') >= 0)) row.Full[property] = full;
        }
    }
}
