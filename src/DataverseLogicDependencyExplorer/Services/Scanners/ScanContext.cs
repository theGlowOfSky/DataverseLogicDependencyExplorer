// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using Microsoft.Xrm.Sdk;
using DataverseLogicDependencyExplorer.Models;

namespace DataverseLogicDependencyExplorer.Services.Scanners
{
    /// <summary>Shared state of one analysis run: the request, the connection and the caches that scanners reuse.</summary>
    public class ScanContext
    {
        public IOrganizationService Service { get; set; }
        public ScanRequest Request { get; set; }
        public DateTime TimestampUtc { get; set; }
        public WebResourceRepository WebResources { get; set; }
        public FormRepository Forms { get; set; }
        public Action<string> Report { get; set; }
        /// <summary>Reports progress inside the current scanner: fraction 0..1 of the scanner's own work plus a message. Safe to call from any thread.</summary>
        public Action<double, string> ReportProgress { get; set; }
        /// <summary>Streams findings to the UI while the scanner is still running. Safe to call from any thread.</summary>
        public Action<IList<DependencyResult>> Emit { get; set; }
        /// <summary>Creates the connections used for parallel downloads (clones of the main connection when possible).</summary>
        public Func<IOrganizationService, int, IList<IOrganizationService>> WorkerServiceFactory { get; set; }
        /// <summary>Number of parallel download connections requested.</summary>
        public int DownloadWorkers { get; set; }

        public ScanContext()
        {
            Report = m => { };
            ReportProgress = (f, m) => { };
            Emit = r => { };
            WorkerServiceFactory = ServiceCloner.CreateWorkers;
            DownloadWorkers = 4;
        }

        public DependencyResult NewResult(ScannerKind source, ComponentType type, ConfidenceLevel confidence)
        {
            return new DependencyResult
            {
                Source = source,
                ComponentType = type,
                Confidence = confidence,
                TargetTable = Request.TableLogicalName,
                TargetColumn = Request.Mode == SearchMode.Column ? Request.ColumnLogicalName : null,
                ScanTimestampUtc = TimestampUtc,
                Environment = !string.IsNullOrEmpty(Request.EnvironmentUrl) ? Request.EnvironmentUrl : Request.EnvironmentName
            };
        }
    }

    public interface IDependencyScanner
    {
        ScannerKind Kind { get; }
        ScanOutcome Scan(ScanContext context, CancellationToken ct);
    }

    /// <summary>
    /// Template for scanners: runs <see cref="Execute"/> and converts failures into an honest diagnostic (permission denied,
    /// failed, cancelled) so one failing category never stops the analysis and never looks like "no dependencies".
    /// </summary>
    public abstract class ScannerBase : IDependencyScanner
    {
        public abstract ScannerKind Kind { get; }

        public ScanOutcome Scan(ScanContext context, CancellationToken ct)
        {
            var outcome = new ScanOutcome { Diagnostic = new ScanDiagnostic { Category = Kind } };
            var sw = Stopwatch.StartNew();
            try
            {
                Execute(context, outcome, ct);
                if (outcome.Diagnostic.Status == ScanStatus.NotInspected)
                    outcome.Diagnostic.Status = outcome.Diagnostic.Warnings.Count > 0 ? ScanStatus.CompletedWithWarnings : ScanStatus.Completed;
            }
            catch (OperationCanceledException)
            {
                outcome.Diagnostic.Status = ScanStatus.Cancelled;
                outcome.Diagnostic.Warn("The scan was cancelled; results for this category are incomplete.");
            }
            catch (Exception ex)
            {
                Trace.TraceError(DateTime.Now.ToString("s") + "  " + Kind + " scan failed: " + ex);
                if (DataverseHelper.IsPermissionDenied(ex))
                {
                    outcome.Diagnostic.Status = ScanStatus.PermissionDenied;
                    outcome.Diagnostic.PermissionLimited = true;
                    outcome.Diagnostic.Error("Permission denied: the connected user cannot read the data this scanner needs (" + DataverseHelper.Describe(ex) + ").");
                }
                else
                {
                    outcome.Diagnostic.Status = ScanStatus.Failed;
                    outcome.Diagnostic.Error("The scan failed: " + DataverseHelper.Describe(ex));
                }
            }
            sw.Stop();
            outcome.Diagnostic.ElapsedMilliseconds = sw.ElapsedMilliseconds;
            outcome.Diagnostic.MatchCount = outcome.Results.Count;
            return outcome;
        }

        protected abstract void Execute(ScanContext context, ScanOutcome outcome, CancellationToken ct);

        protected static string EntityStateText(int? state, int? status)
        {
            return state.HasValue ? (state.Value == 1 ? "Activated" : state.Value == 0 ? "Draft" : "State " + state.Value) : null;
        }
    }
}
