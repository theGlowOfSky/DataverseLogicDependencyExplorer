// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using DataverseLogicDependencyExplorer.Models;
using DataverseLogicDependencyExplorer.Services;
using DataverseLogicDependencyExplorer.Services.Scanners;
using F = DataverseLogicDependencyExplorer.Tests.FakeOrganizationService;

namespace DataverseLogicDependencyExplorer.Tests
{
    [TestClass]
    public class JavaScriptPipelineTests
    {
        private static Entity Script(string name, string js, bool managed = false)
        {
            return F.Row("webresource", null, "name", name, "ismanaged", managed,
                         "content", Convert.ToBase64String(new UTF8Encoding(false).GetBytes(js)));
        }

        private static List<Entity> Scripts(int count, int withMatchEvery = 5)
        {
            var list = new List<Entity>();
            for (int i = 0; i < count; i++)
                list.Add(Script("new_/js/file" + i.ToString("000") + ".js",
                    i % withMatchEvery == 0 ? "var a = f.getAttribute('new_status'); // file " + i : "var unrelated = " + i + ";"));
            return list;
        }

        private static ScanContext Ctx(FakeOrganizationService svc, Action<ScanRequest> tweak = null)
        {
            var req = new ScanRequest
            {
                Mode = SearchMode.Column, TableLogicalName = "account", ColumnLogicalName = "new_status",
                Scopes = new HashSet<ScannerKind> { ScannerKind.JavaScript }, JavaScriptScope = JavaScriptScope.AllScriptWebResources
            };
            if (tweak != null) tweak(req);
            return new ScanContext { Service = svc, Request = req, TimestampUtc = DateTime.UtcNow, WebResources = new WebResourceRepository(svc), Forms = new FormRepository(svc) };
        }

        private static FakeOrganizationService Env(List<Entity> scripts)
        {
            return new FakeOrganizationService().On("systemform").OnWebResources(scripts.ToArray());
        }

        [TestMethod]
        public void Pipeline_ListsWithoutContent_ThenDownloadsInSmallChunks()
        {
            var scripts = Scripts(25);
            var svc = Env(scripts);
            ScanOutcome o = new JavaScriptScanner().Scan(Ctx(svc), CancellationToken.None);

            Assert.AreEqual(ScanStatus.Completed, o.Diagnostic.Status);
            Assert.AreEqual(25, o.Diagnostic.RecordsExamined);
            Assert.AreEqual(5, o.Results.Count);

            List<QueryExpression> wr = svc.Queries.Where(q => q.EntityName == "webresource").ToList();
            QueryExpression listing = wr.First(q => !q.ColumnSet.Columns.Contains("content"));
            Assert.AreEqual(1, wr.Count(q => !q.ColumnSet.Columns.Contains("content")), "one cheap listing query");
            Assert.AreEqual(5000, listing.PageInfo.Count, "listing uses the maximum page size");
            List<QueryExpression> downloads = wr.Where(q => q.ColumnSet.Columns.Contains("content")).ToList();
            Assert.AreEqual(3, downloads.Count, "25 scripts in chunks of 10");
            Assert.IsTrue(downloads.All(q => q.Criteria.Conditions.Any(c => c.AttributeName == "webresourceid" && c.Values.Count <= 10)));
        }

        [TestMethod]
        public void Pipeline_UsesAllWorkerConnections_AndFindsEverything()
        {
            var scripts = Scripts(60);
            var main = Env(scripts);
            var workers = new List<FakeOrganizationService>();
            for (int i = 0; i < 3; i++) workers.Add(new FakeOrganizationService { DelayMs = 30 }.OnWebResources(scripts.ToArray()));
            ScanContext ctx = Ctx(main);
            ctx.DownloadWorkers = 3;
            ctx.WorkerServiceFactory = (svc, n) => workers.Cast<IOrganizationService>().ToList();

            ScanOutcome o = new JavaScriptScanner().Scan(ctx, CancellationToken.None);

            Assert.AreEqual(ScanStatus.Completed, o.Diagnostic.Status);
            Assert.AreEqual(60, o.Diagnostic.RecordsExamined);
            Assert.AreEqual(12, o.Results.Count);
            Assert.IsTrue(workers.Count(w => w.Queries.Count > 0) >= 2, "downloads were spread over several connections");
            Assert.AreEqual(6, workers.Sum(w => w.Queries.Count), "every chunk downloaded exactly once");
        }

        [TestMethod]
        public void Pipeline_StreamsFindingsAndProgressWhileRunning()
        {
            var svc = Env(Scripts(40));
            ScanContext ctx = Ctx(svc);
            var emitted = new List<DependencyResult>(); var progress = new List<KeyValuePair<double, string>>();
            ctx.Emit = found => { lock (emitted) emitted.AddRange(found); };
            ctx.ReportProgress = (f, m) => { lock (progress) progress.Add(new KeyValuePair<double, string>(f, m)); };

            ScanOutcome o = new JavaScriptScanner().Scan(ctx, CancellationToken.None);

            Assert.IsTrue(o.ResultsAlreadyEmitted, "the scanner streams, so the service must not announce its results twice");
            Assert.AreEqual(o.Results.Count, emitted.Count, "every finding was emitted live");
            Assert.IsTrue(progress.Any(p => p.Value.Contains("downloaded") && p.Value.Contains("analysed")));
            Assert.IsTrue(progress.All(p => p.Key >= 0 && p.Key <= 1));
            Assert.AreEqual(1.0, progress.Max(p => p.Key) > 0.9 ? 1.0 : 0.0, "progress reaches the end of the scanner's range");
        }

        [TestMethod]
        public void Pipeline_SkipManaged_AddsServerFilter_AndStatesTheLimit()
        {
            var scripts = Scripts(10);
            scripts.Add(Script("msdyn_/managed.js", "f.getAttribute('new_status');", managed: true));
            var svc = Env(scripts);
            ScanOutcome o = new JavaScriptScanner().Scan(Ctx(svc, r => r.SkipManagedWebResources = true), CancellationToken.None);
            QueryExpression listing = svc.Queries.First(q => q.EntityName == "webresource" && !q.ColumnSet.Columns.Contains("content"));
            Assert.IsTrue(listing.Criteria.Conditions.Any(c => c.AttributeName == "ismanaged"));
            Assert.AreEqual(10, o.Diagnostic.RecordsExamined);
            Assert.IsFalse(o.Results.Any(r => r.ComponentName.StartsWith("msdyn_")));
            StringAssert.Contains(string.Join(" ", o.Diagnostic.Limitations), "Managed");

            ScanOutcome all = new JavaScriptScanner().Scan(Ctx(Env(scripts)), CancellationToken.None);
            Assert.IsTrue(all.Results.Any(r => r.ComponentName.StartsWith("msdyn_")));
        }

        [TestMethod]
        public void Pipeline_FilesWithoutTheName_AreCountedButProduceNothing()
        {
            var svc = Env(Scripts(30, withMatchEvery: 1000).Skip(1).ToList());
            ScanOutcome o = new JavaScriptScanner().Scan(Ctx(svc), CancellationToken.None);
            Assert.AreEqual(29, o.Diagnostic.RecordsExamined);
            Assert.AreEqual(0, o.Results.Count);
            Assert.AreEqual(ScanStatus.Completed, o.Diagnostic.Status);
        }

        [TestMethod]
        public void Pipeline_RegisteredLibraries_AreRetainedForTheFormScanner_AndNothingElseIsKept()
        {
            var scripts = Scripts(15);
            scripts.Add(Script("new_/js/account.js", "var NS={}; NS.Account={onLoad:function(c){ c.getFormContext().getAttribute('new_status'); }};"));
            var form = F.Row("systemform", null, "name", "Main", "objecttypecode", "account", "type", new OptionSetValue(2), "formxml", ParserTests.FormXml);
            var svc = new FakeOrganizationService().On("systemform", form).OnWebResources(scripts.ToArray());
            ScanContext ctx = Ctx(svc);
            new JavaScriptScanner().Scan(ctx, CancellationToken.None);
            Assert.IsNotNull(ctx.WebResources.TryGet("new_/js/account.js"), "registered library retained");
            Assert.IsNull(ctx.WebResources.TryGet("new_/js/file000.js"), "unregistered scripts are released after analysis");
            Assert.AreEqual(1, ctx.WebResources.RetainedCount);
        }

        [TestMethod]
        public void Pipeline_DownloadFailure_IsReportedAsFailed_AndDoesNotHang()
        {
            var scripts = Scripts(30);
            var broken = new FakeOrganizationService().OnThrow("webresource", new InvalidOperationException("download exploded"));
            ScanContext ctx = Ctx(Env(scripts));
            ctx.WorkerServiceFactory = (s, n) => new List<IOrganizationService> { broken };
            ScanOutcome o = new JavaScriptScanner().Scan(ctx, CancellationToken.None);
            Assert.AreEqual(ScanStatus.Failed, o.Diagnostic.Status);
            StringAssert.Contains(o.Diagnostic.Message, "download exploded");
        }

        [TestMethod]
        public void Pipeline_PermissionDeniedWhileDownloading_IsReportedAsPermissionDenied()
        {
            var scripts = Scripts(30);
            var denied = new FakeOrganizationService().OnThrow("webresource", F.PermissionFault());
            ScanContext ctx = Ctx(Env(scripts));
            ctx.WorkerServiceFactory = (s, n) => new List<IOrganizationService> { denied };
            ScanOutcome o = new JavaScriptScanner().Scan(ctx, CancellationToken.None);
            Assert.AreEqual(ScanStatus.PermissionDenied, o.Diagnostic.Status);
        }

        [TestMethod]
        public void Pipeline_Cancellation_StopsPromptly_KeepsPartialResults()
        {
            var cts = new CancellationTokenSource();
            var svc = Env(Scripts(200, 1));
            ScanContext ctx = Ctx(svc);
            ctx.Emit = found => cts.Cancel();   // cancel as soon as the first finding is streamed
            var sw = System.Diagnostics.Stopwatch.StartNew();
            ScanOutcome o = new JavaScriptScanner().Scan(ctx, cts.Token);
            Assert.AreEqual(ScanStatus.Cancelled, o.Diagnostic.Status);
            Assert.IsTrue(o.Results.Count >= 1 && o.Results.Count < 200, "partial results kept");
            Assert.IsTrue(sw.ElapsedMilliseconds < 5000, "cancellation is prompt");
        }

        [TestMethod]
        public void Pipeline_EmptyEnvironment_CompletesWithZeroExamined()
        {
            ScanOutcome o = new JavaScriptScanner().Scan(Ctx(Env(new List<Entity>())), CancellationToken.None);
            Assert.AreEqual(ScanStatus.Completed, o.Diagnostic.Status);
            Assert.AreEqual(0, o.Diagnostic.RecordsExamined);
        }

        // ------------------------------------------------------------------ cloning

        public class Cloneable : IOrganizationService, IDisposable
        {
            public static int Created, Disposed;
            public Cloneable Clone() { Created++; return new Cloneable(); }
            public void Dispose() { Disposed++; }
            public Guid Create(Entity e) { throw new NotSupportedException(); }
            public Entity Retrieve(string n, Guid id, ColumnSet c) { throw new NotSupportedException(); }
            public void Update(Entity e) { throw new NotSupportedException(); }
            public void Delete(string n, Guid id) { throw new NotSupportedException(); }
            public OrganizationResponse Execute(OrganizationRequest r) { throw new NotSupportedException(); }
            public void Associate(string n, Guid id, Relationship r, EntityReferenceCollection c) { throw new NotSupportedException(); }
            public void Disassociate(string n, Guid id, Relationship r, EntityReferenceCollection c) { throw new NotSupportedException(); }
            public EntityCollection RetrieveMultiple(QueryBase q) { throw new NotSupportedException(); }
        }

        [TestMethod]
        public void ServiceCloner_ClonesWhenSupported_AndNeverDisposesTheOriginal()
        {
            Cloneable.Created = Cloneable.Disposed = 0;
            var original = new Cloneable();
            IList<IOrganizationService> workers = ServiceCloner.CreateWorkers(original, 4);
            Assert.AreEqual(4, workers.Count);
            Assert.IsFalse(workers.Any(w => ReferenceEquals(w, original)));
            ServiceCloner.Release(original, workers);
            Assert.AreEqual(4, Cloneable.Disposed);
        }

        [TestMethod]
        public void ServiceCloner_FallsBackToTheOriginalConnection()
        {
            var plain = new FakeOrganizationService();
            IList<IOrganizationService> workers = ServiceCloner.CreateWorkers(plain, 4);
            Assert.AreEqual(1, workers.Count);
            Assert.IsTrue(ReferenceEquals(plain, workers[0]));
            ServiceCloner.Release(plain, workers); // must not throw or dispose the original
            Assert.AreEqual(1, ServiceCloner.CreateWorkers(plain, 1).Count);
        }

        // ------------------------------------------------------------------ service-level live reporting

        private class Collect : IProgress<ProgressInfo>
        {
            public readonly List<ProgressInfo> Items = new List<ProgressInfo>();
            public void Report(ProgressInfo p) { lock (Items) Items.Add(p); }
        }

        [TestMethod]
        public void AnalysisService_ForwardsStreamedFindings_AndMapsScannerProgressIntoTheOverallBar()
        {
            var streaming = new StubScanner(ScannerKind.JavaScript, (c, t) =>
            {
                c.ReportProgress(0.5, "halfway");
                var r = Make.Result(ScannerKind.JavaScript, "live.js");
                c.Emit(new List<DependencyResult> { r });
                var o = Make.Outcome(ScannerKind.JavaScript, ScanStatus.Completed, r);
                o.ResultsAlreadyEmitted = true;
                return o;
            });
            var plain = new StubScanner(ScannerKind.BusinessRules, (c, t) => Make.Outcome(ScannerKind.BusinessRules, ScanStatus.Completed, Make.Result(ScannerKind.BusinessRules, "rule")));

            var progress = new Collect();
            var svc = new DependencyAnalysisService(k => k == ScannerKind.JavaScript ? (IDependencyScanner)streaming : plain);
            ScanRequest req = Make.Request(SearchMode.Table, ScannerKind.BusinessRules, ScannerKind.JavaScript);
            AnalysisResult a = svc.Analyze(new FakeOrganizationService(), req, progress, CancellationToken.None);

            Assert.AreEqual(2, a.MatchCount);
            var live = progress.Items.Where(p => p.NewResults != null).SelectMany(p => p.NewResults).Select(r => r.ComponentName).ToList();
            Assert.AreEqual(1, live.Count(n => n == "live.js"), "streamed findings are forwarded once (not again after the scanner ends)");
            Assert.AreEqual(1, live.Count(n => n == "rule"), "non-streaming scanners hand over their findings when they finish");
            ProgressInfo half = progress.Items.First(p => p.Message == "halfway");
            Assert.AreEqual(75, half.Percent, "scanner 2 of 2 at 50% => 75% overall");
        }

        // ------------------------------------------------------------------ the single warning

        [TestMethod]
        public void Warning_MentionsSkippingManaged_OnlyWhenItIsNotAlreadyTicked()
        {
            string plain = ScanWarnings.AllWebResources(false);
            StringAssert.Contains(plain, "EVERY JavaScript web resource");
            StringAssert.Contains(plain, "Skip managed web resources");
            StringAssert.Contains(plain, "Only libraries registered on the selected table's forms");

            string skipping = ScanWarnings.AllWebResources(true);
            StringAssert.Contains(skipping, "EVERY JavaScript web resource");
            Assert.IsFalse(skipping.Contains("Skip managed"), "no suggestion to tick an option that is already ticked");
            Assert.IsFalse(skipping.ToLowerInvariant().Contains("skip"), "the message does not talk about skipping at all");
            StringAssert.Contains(skipping, "Only libraries registered on the selected table's forms");
        }

        // ------------------------------------------------------------------ product identity

        [TestMethod]
        public void ProductInfo_MatchesTheAssemblyVersion_AndCarriesTheMaintainerDetails()
        {
            Assert.AreEqual("1.2026.11.11", ProductInfo.Version);
            Assert.AreEqual(ProductInfo.Version, typeof(ProductInfo).Assembly.GetName().Version.ToString(), "csproj Version and ProductInfo.Version must agree");
            Assert.AreEqual("Gagan Deep Singh", ProductInfo.Maintainer);
            Assert.AreEqual("https://www.linkedin.com/in/glowofsky/", ProductInfo.LinkedInUrl);
            Assert.AreEqual("https://github.com/theGlowOfSky/DataverseLogicDependencyExplorer", ProductInfo.RepositoryUrl);
            Assert.AreEqual("XrmToolBox", ProductInfo.Host);
            Assert.AreEqual("Microsoft Dataverse", ProductInfo.DataPlatform);
            StringAssert.Contains(ProductInfo.ExportFormat, ".xlsx");
        }
    }
}
