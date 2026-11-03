// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;

namespace DataverseLogicDependencyExplorer.Models
{
    /// <summary>The complete outcome of one analysis run.</summary>
    public class AnalysisResult
    {
        public ScanRequest Request { get; set; }
        public DateTime StartedUtc { get; set; }
        public DateTime FinishedUtc { get; set; }
        public bool WasCancelled { get; set; }
        public List<DependencyResult> Results { get; } = new List<DependencyResult>();
        public List<ScanDiagnostic> Diagnostics { get; } = new List<ScanDiagnostic>();
        /// <summary>Run-wide warnings that are not tied to one category.</summary>
        public List<string> GeneralWarnings { get; } = new List<string>();

        public int ComponentsExamined
        {
            get { return Diagnostics.Where(d => d.RecordsExamined > 0).Sum(d => d.RecordsExamined); }
        }

        public int MatchCount { get { return Results.Count; } }

        public int UniqueComponents
        {
            get
            {
                return Results.Select(r => (int)r.ComponentType + ":" + (r.ComponentId ?? r.ComponentName ?? string.Empty).ToLowerInvariant())
                              .Distinct().Count();
            }
        }

        public int FailedScans
        {
            get { return Diagnostics.Count(d => d.Status == ScanStatus.Failed || d.Status == ScanStatus.PermissionDenied); }
        }

        public int CompletedScans
        {
            get { return Diagnostics.Count(d => d.Status == ScanStatus.Completed || d.Status == ScanStatus.CompletedWithWarnings); }
        }

        /// <summary>Scans that could only inspect part of what exists (or were not inspected / cancelled).</summary>
        public int UnsupportedOrPartialScans
        {
            get { return Diagnostics.Count(d => d.Status == ScanStatus.PartiallySupported || d.Status == ScanStatus.NotInspected || d.Status == ScanStatus.Cancelled); }
        }

        public ScanDiagnostic DiagnosticFor(ScannerKind kind)
        {
            return Diagnostics.FirstOrDefault(d => d.Category == kind);
        }

        /// <summary>
        /// A scan with no matches is only a "clean" result when every selected category was actually inspected.
        /// Used by the UI to avoid presenting a failed or unsupported scan as "no dependencies".
        /// </summary>
        public bool IsCleanEmptyResult
        {
            get
            {
                return Results.Count == 0 && !WasCancelled && Request != null &&
                       Diagnostics.Where(d => Request.Scopes.Contains(d.Category))
                                  .All(d => d.Status == ScanStatus.Completed || d.Status == ScanStatus.CompletedWithWarnings);
            }
        }
    }
}
