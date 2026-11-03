// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using DataverseLogicDependencyExplorer.Models;

namespace DataverseLogicDependencyExplorer.Services.Scanners
{
    /// <summary>
    /// Static text analysis of JavaScript web resources. Every finding is a source-text match (never "confirmed"); the match kind
    /// (API call argument, query token, string literal, identifier) shows how strong the textual evidence is.
    ///
    /// Large environments are handled with a pipeline: web resources are listed cheaply (no content), the content is downloaded in
    /// parallel chunks (one cloned connection per worker where the connection type supports cloning) and analysed by consumer
    /// threads while later chunks are still downloading. Progress and findings are streamed to the UI as they appear. Files that
    /// cannot contain the searched name are skipped with a plain substring test before any tokenizing.
    /// </summary>
    public class JavaScriptScanner : ScannerBase
    {
        private const int MaxMatchesPerResource = 100;
        private const int DownloadChunkSize = 10;
        private const int QueueCapacity = 24;

        public override ScannerKind Kind { get { return ScannerKind.JavaScript; } }

        public static string DescribeKind(JsMatch m)
        {
            switch (m.Kind)
            {
                case JsMatchKind.ApiCallArgument: return "API call argument (" + m.ApiName + ")";
                case JsMatchKind.QueryOrFetchToken: return "Query / FetchXML / URL token";
                case JsMatchKind.StringLiteral: return "String literal";
                case JsMatchKind.Identifier: return "Identifier / property access";
                default: return "Comment (not a runtime reference)";
            }
        }

        /// <summary>Thread-safe counters shared by the analysis workers.</summary>
        private class Totals
        {
            public int Examined, Problems, CommentOnly, Truncated, Downloaded, Analysed;
            public readonly List<string> Warnings = new List<string>();
            public readonly object Gate = new object();
        }

        protected override void Execute(ScanContext ctx, ScanOutcome outcome, CancellationToken ct)
        {
            ScanRequest req = ctx.Request;
            ScanDiagnostic diag = outcome.Diagnostic;
            outcome.ResultsAlreadyEmitted = true; // findings are streamed through ctx.Emit while scanning

            diag.Limit("Static text analysis only: dynamic references, aliases, string concatenation, minified or obfuscated code cannot be resolved and may be missed or reported as uncertain.");
            diag.Limit("A source-text match is not proof of a runtime dependency; the match kind shows how the name was used. Unpublished/managed layers are not distinguished.");
            if (req.Mode == SearchMode.Column && req.ColumnLogicalName.IndexOf('_') < 0)
                diag.Warn("'" + req.ColumnLogicalName + "' is a generic (non-prefixed) column name. Only API-call arguments, query tokens and exact string literals are matched; the same name on other tables cannot be told apart.");

            // Libraries registered on the selected table's forms: explains relevance and enables dynamic-access detection.
            ctx.ReportProgress(0.0, "JavaScript: reading forms of " + req.TableLogicalName + "...");
            var registered = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            try
            {
                foreach (FormRecord f in ctx.Forms.GetForms(req.TableLogicalName, ct))
                {
                    if (!f.Definition.IsValid) continue;
                    foreach (string lib in f.Definition.Libraries)
                    {
                        List<string> forms;
                        if (!registered.TryGetValue(lib, out forms)) registered[lib] = forms = new List<string>();
                        forms.Add(f.Name);
                    }
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                diag.Warn("Forms of '" + req.TableLogicalName + "' could not be read (" + DataverseHelper.Describe(ex) + "); form registration context is missing from the findings.");
                if (req.JavaScriptScope == JavaScriptScope.RegisteredOnSelectedTableForms) throw;
            }

            string[] needles = req.Mode == SearchMode.Column
                ? new[] { req.ColumnLogicalName }
                : new[] { req.TableLogicalName, req.TableEntitySetName }.Where(n => !string.IsNullOrEmpty(n)).ToArray();

            var totals = new Totals();
            if (req.JavaScriptScope == JavaScriptScope.AllScriptWebResources)
                RunPipeline(ctx, outcome, registered, needles, totals, ct);
            else
                RunTargeted(ctx, outcome, registered, needles, totals, ct);

            foreach (string w in totals.Warnings.Distinct()) diag.Warn(w);
            diag.RecordsExamined = totals.Examined;
            if (totals.Problems > 0) diag.Warn(totals.Problems + " script web resource(s) had empty or undecodable content and were skipped.");
            if (totals.CommentOnly > 0) diag.Limit(totals.CommentOnly + " match(es) occurred only inside comments and were not listed (enable 'include comment matches' to list them).");
            if (totals.Truncated > 0) diag.Warn(totals.Truncated + " script(s) had more than " + MaxMatchesPerResource + " matches; only the first " + MaxMatchesPerResource + " of each are listed.");
        }

        // ------------------------------------------------------------------ targeted scope (registered libraries only)

        private void RunTargeted(ScanContext ctx, ScanOutcome outcome, Dictionary<string, List<string>> registered, string[] needles, Totals totals, CancellationToken ct)
        {
            ScanRequest req = ctx.Request;
            ScanDiagnostic diag = outcome.Diagnostic;
            ctx.ReportProgress(0.1, "JavaScript: downloading libraries registered on the table's forms...");

            var warnings = new List<string>();
            ctx.WebResources.EnsureNames(registered.Keys, ct, warnings);
            totals.Warnings.AddRange(warnings);
            diag.Limit("Only libraries registered on forms of '" + req.TableLogicalName + "' were analysed (targeted scope); scripts that are not registered on those forms were not searched.");

            var scripts = registered.Keys.Select(n => ctx.WebResources.TryGet(n)).Where(w => w != null).ToList();
            foreach (string lib in registered.Keys.Where(n => ctx.WebResources.TryGet(n) == null))
                diag.Warn("Library '" + lib + "' is registered on a form but was not found as a script web resource.");

            int i = 0;
            foreach (WebResourceContent script in scripts)
            {
                ct.ThrowIfCancellationRequested();
                ProcessScript(ctx, outcome, registered, needles, totals, script);
                ctx.ReportProgress(0.1 + 0.85 * (++i) / Math.Max(1, scripts.Count), "JavaScript: analysed " + i + " of " + scripts.Count + " libraries...");
            }
        }

        // ------------------------------------------------------------------ pipeline (all script web resources)

        private void RunPipeline(ScanContext ctx, ScanOutcome outcome, Dictionary<string, List<string>> registered, string[] needles, Totals totals, CancellationToken ct)
        {
            ScanRequest req = ctx.Request;
            ScanDiagnostic diag = outcome.Diagnostic;

            ctx.ReportProgress(0.02, "JavaScript: listing script web resources...");
            var listWarnings = new List<string>();
            List<WebResourceHeader> headers = ctx.WebResources.ListScriptHeaders(ct, req.SkipManagedWebResources, listWarnings);
            totals.Warnings.AddRange(listWarnings);
            if (req.SkipManagedWebResources)
                diag.Limit("Managed (solution-installed) web resources were skipped by option; references inside them were not searched.");

            int total = headers.Count;
            if (total == 0) { ctx.ReportProgress(1.0, "JavaScript: no script web resources found."); return; }

            var chunks = new ConcurrentQueue<List<Guid>>(DataverseHelper.Chunk(headers.Select(h => h.Id), DownloadChunkSize));
            IList<IOrganizationService> workers = ctx.WorkerServiceFactory(ctx.Service, Math.Max(1, ctx.DownloadWorkers));
            int analysers = Math.Max(1, Math.Min(Environment.ProcessorCount, 4));
            var queue = new BlockingCollection<WebResourceContent>(QueueCapacity);
            var clock = Stopwatch.StartNew();
            long lastReportMs = -1000;
            Exception fault = null;

            Action<bool> report = force =>
            {
                long now = clock.ElapsedMilliseconds;
                if (!force && now - lastReportMs < 200) return; // throttle: the UI does not need more than a few updates per second
                lastReportMs = now;
                int downloaded = Volatile.Read(ref totals.Downloaded), analysed = Volatile.Read(ref totals.Analysed);
                int found;
                lock (outcome) found = outcome.Results.Count;
                ctx.ReportProgress(0.03 + 0.95 * analysed / total,
                    "JavaScript: downloaded " + downloaded + "/" + total + ", analysed " + analysed + "/" + total + ", " + found + " match(es) so far  (" +
                    TimeSpan.FromMilliseconds(now).ToString(@"mm\:ss") + ", " + workers.Count + " download connection(s))");
            };

            try
            {
                using (CancellationTokenSource linked = CancellationTokenSource.CreateLinkedTokenSource(ct))
                {
                    CancellationToken token = linked.Token;

                    Task[] downloaders = workers.Select(connection => Task.Run(() =>
                    {
                        try
                        {
                            List<Guid> chunk;
                            while (!token.IsCancellationRequested && chunks.TryDequeue(out chunk))
                            {
                                foreach (WebResourceContent item in ctx.WebResources.DownloadByIds(connection, chunk, token))
                                    queue.Add(item, token);
                                Interlocked.Add(ref totals.Downloaded, chunk.Count);
                                report(false);
                            }
                        }
                        catch (OperationCanceledException) { }
                        catch (Exception ex)
                        {
                            Interlocked.CompareExchange(ref fault, ex, null);
                            linked.Cancel();
                        }
                    })).ToArray();

                    Task closer = Task.Run(() =>
                    {
                        try { Task.WaitAll(downloaders); }
                        finally { queue.CompleteAdding(); }
                    });

                    Task[] consumers = Enumerable.Range(0, analysers).Select(n => Task.Run(() =>
                    {
                        try
                        {
                            foreach (WebResourceContent item in queue.GetConsumingEnumerable(token))
                            {
                                ProcessScript(ctx, outcome, registered, needles, totals, item);
                                Interlocked.Increment(ref totals.Analysed);
                                report(false);
                            }
                        }
                        catch (OperationCanceledException) { }
                        catch (Exception ex)
                        {
                            Interlocked.CompareExchange(ref fault, ex, null);
                            linked.Cancel();
                        }
                    })).ToArray();

                    Task.WaitAll(consumers);
                    closer.Wait();
                }

                ct.ThrowIfCancellationRequested();
                if (fault != null) ExceptionDispatchInfo.Capture(fault).Throw();
                report(true);
            }
            finally
            {
                queue.Dispose();
                ServiceCloner.Release(ctx.Service, workers);
            }
        }

        // ------------------------------------------------------------------ analysis of one script (thread-safe)

        private void ProcessScript(ScanContext ctx, ScanOutcome outcome, Dictionary<string, List<string>> registered, string[] needles, Totals totals, WebResourceContent script)
        {
            ScanRequest req = ctx.Request;

            List<string> forms;
            bool isRegistered = registered.TryGetValue(script.Name, out forms);
            if (isRegistered) ctx.WebResources.Retain(script); // other scanners (forms) reuse registered libraries without downloading again

            if (script.Text == null)
            {
                if (script.Problem != null) Interlocked.Increment(ref totals.Problems);
                return;
            }
            Interlocked.Increment(ref totals.Examined);

            // Cheap pre-filter: a file that does not even contain the name cannot reference it, so skip tokenizing it.
            bool mayMatch = needles.Any(n => script.Text.IndexOf(n, StringComparison.OrdinalIgnoreCase) >= 0);
            bool needsDynamicCheck = req.Mode == SearchMode.Column && isRegistered;
            if (!mayMatch && !needsDynamicCheck) return;

            var found = new List<DependencyResult>();
            string context = isRegistered
                ? "Registered on " + forms.Distinct().Count() + " form(s) of '" + req.TableLogicalName + "': " + DataverseHelper.Truncate(string.Join(", ", forms.Distinct().Take(5)), 200)
                : "Not registered on any form of '" + req.TableLogicalName + "' - the reference may concern another table or be used indirectly.";

            if (mayMatch)
            {
                JsAnalysis analysis = req.Mode == SearchMode.Column
                    ? JsSourceAnalyzer.FindColumn(script.Text, req.ColumnLogicalName, req.IncludeCommentMatches)
                    : JsSourceAnalyzer.FindTable(script.Text, req.TableLogicalName, req.TableEntitySetName, req.IncludeCommentMatches);
                Interlocked.Add(ref totals.CommentOnly, analysis.CommentOnlyMatches);

                foreach (JsMatch m in analysis.Matches.Take(MaxMatchesPerResource))
                {
                    DependencyResult r = ctx.NewResult(ScannerKind.JavaScript, ComponentType.JavaScriptWebResource, ConfidenceLevel.TextMatchReviewRequired);
                    r.ComponentName = script.Name;
                    r.ComponentId = script.Id.ToString();
                    r.ComponentUniqueName = script.Name;
                    r.ComponentState = script.IsManaged == true ? "Managed" : script.IsManaged == false ? "Unmanaged" : null;
                    r.ParentTable = isRegistered ? req.TableLogicalName : null;
                    r.MatchType = DescribeKind(m);
                    r.MatchedValue = m.MatchedText;
                    r.Evidence = m.Snippet;
                    r.SourceResource = script.Name;
                    r.LineNumber = m.Line;
                    r.Location = "line " + m.Line;
                    r.AnalysisMethod = "Identifier-aware static text analysis (comments and strings separated)";
                    r.Explanation = "The JavaScript source contains '" + m.MatchedText + "' as: " + DescribeKind(m) + ". " + context;
                    r.Limitation = "Static text match; runtime behaviour is not verified.";
                    r.AdditionalInfo = "Encoding: " + script.EncodingName;
                    found.Add(r);
                }
                if (analysis.Matches.Count > MaxMatchesPerResource) Interlocked.Increment(ref totals.Truncated);
            }

            if (needsDynamicCheck)
            {
                DynamicAccessInfo dyn = JsSourceAnalyzer.FindDynamicAccess(script.Text);
                if (dyn.Count > 0)
                {
                    DependencyResult r = ctx.NewResult(ScannerKind.JavaScript, ComponentType.JavaScriptWebResource, ConfidenceLevel.PossibleOrDynamic);
                    r.ComponentName = script.Name;
                    r.ComponentId = script.Id.ToString();
                    r.ComponentUniqueName = script.Name;
                    r.ParentTable = req.TableLogicalName;
                    r.MatchType = "Dynamic attribute access";
                    r.MatchedValue = req.ColumnLogicalName;
                    r.Evidence = dyn.Count + " attribute/control lookup(s) use a variable or concatenated name; first at line " + dyn.FirstLine + ": " + dyn.FirstSnippet;
                    r.SourceResource = script.Name;
                    r.LineNumber = dyn.FirstLine;
                    r.Location = "line " + dyn.FirstLine;
                    r.AnalysisMethod = "Static detection of non-literal getAttribute/getControl arguments";
                    r.Explanation = "This library runs on forms of '" + req.TableLogicalName + "' and builds attribute names dynamically; it may read or write '" + req.ColumnLogicalName + "'.";
                    r.Limitation = "Cannot be resolved statically; the column may or may not be used.";
                    found.Add(r);
                }
            }

            if (found.Count > 0)
            {
                lock (outcome) outcome.Results.AddRange(found);
                ctx.Emit(found); // live: the UI shows these findings while the scan continues
            }
        }
    }
}
