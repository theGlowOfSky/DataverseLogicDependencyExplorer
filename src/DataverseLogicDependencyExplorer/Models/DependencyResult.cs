// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;

namespace DataverseLogicDependencyExplorer.Models
{
    /// <summary>
    /// One finding produced by a scanner. The model is independent of the UI and of the Excel output so that every
    /// scanner can fill it and every consumer (grid, workbook, tests) can read it.
    /// </summary>
    public class DependencyResult
    {
        public string TargetTable { get; set; }
        public string TargetColumn { get; set; }

        public ComponentType ComponentType { get; set; }
        public string ComponentName { get; set; }
        public string ComponentId { get; set; }
        /// <summary>Unique name / logical name of the component, when it has one.</summary>
        public string ComponentUniqueName { get; set; }
        /// <summary>The table the component belongs to (form table, business rule primary table, step primary table...).</summary>
        public string ParentTable { get; set; }
        public string ComponentState { get; set; }
        public string SolutionName { get; set; }

        public string MatchType { get; set; }
        /// <summary>The table/column value that matched.</summary>
        public string MatchedValue { get; set; }
        /// <summary>Matching expression, XML node, JSON expression, registered attribute or source-code snippet.</summary>
        public string Evidence { get; set; }
        /// <summary>File/resource/definition the evidence came from (web resource name, form name, "workflow.xaml"...).</summary>
        public string SourceResource { get; set; }
        /// <summary>XML node path, JSON path or attribute name where the match was found.</summary>
        public string Location { get; set; }
        /// <summary>1-based line number; only set when it is accurately known.</summary>
        public int? LineNumber { get; set; }

        public ConfidenceLevel Confidence { get; set; }
        public string AnalysisMethod { get; set; }
        /// <summary>Why this result was included.</summary>
        public string Explanation { get; set; }
        /// <summary>Technical limitation affecting this finding.</summary>
        public string Limitation { get; set; }
        /// <summary>Extra information such as connection references or plugin assembly details.</summary>
        public string AdditionalInfo { get; set; }

        public DateTime ScanTimestampUtc { get; set; }
        public string Environment { get; set; }

        /// <summary>Which scanner produced the finding (used for per-category counts).</summary>
        public ScannerKind Source { get; set; }

        /// <summary>Key that identifies "the same reference" so overlapping analysis paths do not create duplicates.</summary>
        public string DuplicateKey()
        {
            return string.Join("|", new[]
            {
                ((int)ComponentType).ToString(),
                (ComponentId ?? string.Empty).ToLowerInvariant(),
                (MatchType ?? string.Empty).ToLowerInvariant(),
                (MatchedValue ?? string.Empty).ToLowerInvariant(),
                (Location ?? string.Empty),
                LineNumber.HasValue ? LineNumber.Value.ToString() : string.Empty,
                (SourceResource ?? string.Empty).ToLowerInvariant()
            });
        }
    }
}
