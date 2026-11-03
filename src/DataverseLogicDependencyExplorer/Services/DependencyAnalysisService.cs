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
using DataverseLogicDependencyExplorer.Models;
using DataverseLogicDependencyExplorer.Services.Scanners;

namespace DataverseLogicDependencyExplorer.Services
{
    /// <summary>
    /// Runs the selected scanners one after another, isolating failures per scanner, and assembles the combined result. Contains
    /// no UI code: progress is reported through <see cref="IProgress{T}"/>.
    /// </summary>
    public class DependencyAnalysisService
    {
        private readonly Func<ScannerKind, IDependencyScanner> scannerFactory;

        public DependencyAnalysisService() : this(DefaultScanner) { }

        /// <summary>Allows tests to substitute scanners.</summary>
        public DependencyAnalysisService(Func<ScannerKind, IDependencyScanner> scannerFactory)
        {
            this.scannerFactory = scannerFactory;
        }

        public static IDependencyScanner DefaultScanner(ScannerKind kind)
        {
            switch (kind)
            {
                case ScannerKind.BusinessRules: return new BusinessRuleScanner();
                case ScannerKind.JavaScript: return new JavaScriptScanner();
                case ScannerKind.Forms: return new FormScanner();
                case ScannerKind.Plugins: return new PluginRegistrationScanner();
                case ScannerKind.ClassicWorkflows: return new ClassicWorkflowScanner();
                case ScannerKind.CloudFlows: return new CloudFlowScanner();
                default: throw new ArgumentOutOfRangeException("kind");
            }
        }

        public AnalysisResult Analyze(IOrganizationService service, ScanRequest request, IProgress<ProgressInfo> progress, CancellationToken ct)
        {
            if (service == null) throw new ArgumentNullException("service");
            if (request == null) throw new ArgumentNullException("request");
            string problem = request.Validate();
            if (problem != null) throw new ArgumentException(problem);

            var analysis = new AnalysisResult { Request = request, StartedUtc = DateTime.UtcNow };
            var ctx = new ScanContext
            {
                Service = service,
                Request = request,
                TimestampUtc = analysis.StartedUtc,
                WebResources = new WebResourceRepository(service),
                Forms = new FormRepository(service)
            };

            List<ScannerKind> selected = ScanRequest.AllScopes.Where(k => request.Scopes.Contains(k)).ToList();
            int done = 0;
            var all = new List<DependencyResult>();

            foreach (ScannerKind kind in ScanRequest.AllScopes)
            {
                if (!request.Scopes.Contains(kind))
                {
                    var skipped = new ScanDiagnostic { Category = kind, Status = ScanStatus.NotInspected };
                    skipped.Warn("Not inspected: this category was not selected for the scan. This is NOT a 'no dependencies' result.");
                    analysis.Diagnostics.Add(skipped);
                    continue;
                }

                if (ct.IsCancellationRequested)
                {
                    analysis.WasCancelled = true;
                    var notStarted = new ScanDiagnostic { Category = kind, Status = ScanStatus.NotInspected };
                    notStarted.Warn("Not inspected: the scan was cancelled before this category started.");
                    analysis.Diagnostics.Add(notStarted);
                    continue;
                }

                int selectedCount = Math.Max(1, selected.Count);
                int percent = done * 100 / selectedCount;
                Report(progress, percent, "Scanning " + kind.Display() + "...");
                int baseDone = done;
                // A scanner reports its own 0..1 fraction; it is mapped into this scanner's share of the overall bar.
                ctx.Report = m => Report(progress, percent, m);
                ctx.ReportProgress = (fraction, m) => Report(progress, (int)((baseDone + Math.Max(0, Math.Min(1, fraction))) * 100 / selectedCount), m);
                ctx.Emit = found => ReportFound(progress, (int)(baseDone * 100 / selectedCount), found);

                ScanOutcome outcome;
                try { outcome = scannerFactory(kind).Scan(ctx, ct); }
                catch (Exception ex) // a scanner must never take down the run
                {
                    outcome = new ScanOutcome { Diagnostic = new ScanDiagnostic { Category = kind, Status = ScanStatus.Failed } };
                    outcome.Diagnostic.Error("The scanner crashed: " + ex.Message);
                }
                outcome.Diagnostic.Category = kind;
                if (outcome.Diagnostic.Status == ScanStatus.Cancelled) analysis.WasCancelled = true;
                analysis.Diagnostics.Add(outcome.Diagnostic);
                all.AddRange(outcome.Results);
                done++;
                // Scanners that did not stream their findings while running hand them over now (still provisional).
                if (!outcome.ResultsAlreadyEmitted && outcome.Results.Count > 0)
                    ReportFound(progress, done * 100 / selectedCount, outcome.Results);
            }

            if (ct.IsCancellationRequested) analysis.WasCancelled = true;

            Report(progress, 95, "Finishing: removing duplicates and resolving solutions...");
            List<DependencyResult> unique = ResultDeduplicator.Deduplicate(all);
            foreach (ScanDiagnostic d in analysis.Diagnostics)
                d.MatchCount = unique.Count(r => r.Source == d.Category);

            if (unique.Count > 0)
            {
                var solutions = new SolutionLookup(service);
                try { solutions.Enrich(unique, ct); }
                catch (OperationCanceledException) { analysis.WasCancelled = true; }
                catch (Exception ex) { analysis.GeneralWarnings.Add("Solution information could not be resolved: " + DataverseHelper.Describe(ex)); }
                analysis.GeneralWarnings.AddRange(solutions.Warnings);
            }

            analysis.Results.AddRange(ResultDeduplicator.Sort(unique));
            if (analysis.WasCancelled) analysis.GeneralWarnings.Add("The scan was cancelled. Results are incomplete and categories that did not finish are marked accordingly.");
            analysis.FinishedUtc = DateTime.UtcNow;
            Report(progress, 100, analysis.WasCancelled ? "Scan cancelled." : "Scan finished.");
            return analysis;
        }

        private static void Report(IProgress<ProgressInfo> progress, int percent, string message)
        {
            if (progress != null) progress.Report(new ProgressInfo { Percent = Math.Max(0, Math.Min(100, percent)), Message = message });
        }

        private static void ReportFound(IProgress<ProgressInfo> progress, int percent, IList<DependencyResult> found)
        {
            if (progress == null || found == null || found.Count == 0) return;
            // A snapshot: the scanner keeps adding to its own list while the UI reads this one.
            progress.Report(new ProgressInfo { Percent = Math.Max(0, Math.Min(100, percent)), NewResults = found.ToList() });
        }
    }
}
