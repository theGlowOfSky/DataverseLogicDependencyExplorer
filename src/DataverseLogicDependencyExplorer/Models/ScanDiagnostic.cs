// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;

namespace DataverseLogicDependencyExplorer.Models
{
    /// <summary>What happened while one component category was scanned.</summary>
    public class ScanDiagnostic
    {
        public ScannerKind Category { get; set; }
        public ScanStatus Status { get; set; } = ScanStatus.NotInspected;
        /// <summary>Records or definitions examined, where measurable.</summary>
        public int RecordsExamined { get; set; }
        public int MatchCount { get; set; }
        public int MalformedDefinitions { get; set; }
        public bool PermissionLimited { get; set; }
        public long ElapsedMilliseconds { get; set; }
        public List<string> Warnings { get; } = new List<string>();
        public List<string> Errors { get; } = new List<string>();
        /// <summary>Capability notes: what this scanner can and cannot see (shown even when the scan succeeded).</summary>
        public List<string> Limitations { get; } = new List<string>();

        public void Warn(string message) { if (!string.IsNullOrWhiteSpace(message)) Warnings.Add(message); }
        public void Error(string message) { if (!string.IsNullOrWhiteSpace(message)) Errors.Add(message); }
        public void Limit(string message) { if (!string.IsNullOrWhiteSpace(message)) Limitations.Add(message); }

        /// <summary>True when the category was actually inspected (even partially) - used to tell "no matches" from "not looked at".</summary>
        public bool WasInspected
        {
            get
            {
                return Status == ScanStatus.Completed || Status == ScanStatus.CompletedWithWarnings
                    || Status == ScanStatus.PartiallySupported || (Status == ScanStatus.Cancelled && RecordsExamined > 0);
            }
        }

        public string Message
        {
            get
            {
                var parts = new List<string>();
                parts.AddRange(Errors);
                parts.AddRange(Warnings);
                return string.Join(" | ", parts);
            }
        }
    }

    /// <summary>The scanner output: findings plus the diagnostic for that category.</summary>
    public class ScanOutcome
    {
        public List<DependencyResult> Results { get; } = new List<DependencyResult>();
        public ScanDiagnostic Diagnostic { get; set; }
        /// <summary>True when the scanner already streamed its findings through <c>ScanContext.Emit</c> while it was running.</summary>
        public bool ResultsAlreadyEmitted { get; set; }
    }

    public class ProgressInfo
    {
        public int Percent { get; set; }
        public string Message { get; set; }
        /// <summary>Findings discovered since the previous report (provisional: duplicates and solutions are resolved at the end).</summary>
        public List<DependencyResult> NewResults { get; set; }
    }
}
