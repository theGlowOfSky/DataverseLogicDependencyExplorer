// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;
using DataverseLogicDependencyExplorer.Models;

namespace DataverseLogicDependencyExplorer.Services
{
    /// <summary>Filter values of the Impact Report tab. Empty/null means "no restriction".</summary>
    public class FilterCriteria
    {
        public ComponentType? ComponentType { get; set; }
        public string ComponentName { get; set; }
        public string MatchedTableOrColumn { get; set; }
        public ConfidenceLevel? Confidence { get; set; }
        public ScanStatus? ScanStatus { get; set; }
        public string Solution { get; set; }
        /// <summary>Free text searched across every visible field.</summary>
        public string SearchText { get; set; }
    }

    public static class ResultFilter
    {
        public static List<DependencyResult> Apply(AnalysisResult analysis, FilterCriteria c)
        {
            if (analysis == null) return new List<DependencyResult>();
            IEnumerable<DependencyResult> q = analysis.Results;
            if (c == null) return q.ToList();

            if (c.ComponentType.HasValue) q = q.Where(r => r.ComponentType == c.ComponentType.Value);
            if (c.Confidence.HasValue) q = q.Where(r => r.Confidence == c.Confidence.Value);
            if (c.ScanStatus.HasValue)
                q = q.Where(r => { ScanDiagnostic d = analysis.DiagnosticFor(r.Source); return d != null && d.Status == c.ScanStatus.Value; });
            if (!string.IsNullOrWhiteSpace(c.ComponentName)) q = q.Where(r => Has(r.ComponentName, c.ComponentName) || Has(r.ComponentUniqueName, c.ComponentName));
            if (!string.IsNullOrWhiteSpace(c.MatchedTableOrColumn))
                q = q.Where(r => Has(r.TargetTable, c.MatchedTableOrColumn) || Has(r.TargetColumn, c.MatchedTableOrColumn) || Has(r.MatchedValue, c.MatchedTableOrColumn));
            if (!string.IsNullOrWhiteSpace(c.Solution)) q = q.Where(r => Has(r.SolutionName, c.Solution));
            if (!string.IsNullOrWhiteSpace(c.SearchText))
            {
                string t = c.SearchText.Trim();
                q = q.Where(r => Has(r.ComponentName, t) || Has(r.ComponentUniqueName, t) || Has(r.ComponentId, t) || Has(r.ParentTable, t)
                                 || Has(r.MatchType, t) || Has(r.MatchedValue, t) || Has(r.Evidence, t) || Has(r.SourceResource, t)
                                 || Has(r.Location, t) || Has(r.SolutionName, t) || Has(r.Explanation, t) || Has(r.ComponentType.Display(), t));
            }
            return q.ToList();
        }

        private static bool Has(string value, string fragment)
        {
            return value != null && value.IndexOf(fragment, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static List<string> DistinctSolutions(AnalysisResult analysis)
        {
            if (analysis == null) return new List<string>();
            return analysis.Results.Where(r => !string.IsNullOrEmpty(r.SolutionName))
                           .SelectMany(r => r.SolutionName.Split(new[] { ", " }, StringSplitOptions.RemoveEmptyEntries))
                           .Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(s => s, StringComparer.OrdinalIgnoreCase).ToList();
        }
    }
}
