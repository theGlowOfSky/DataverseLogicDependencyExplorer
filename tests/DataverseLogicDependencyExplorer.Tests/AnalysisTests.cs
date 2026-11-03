// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DataverseLogicDependencyExplorer.Models;
using DataverseLogicDependencyExplorer.Services;
using DataverseLogicDependencyExplorer.Services.Scanners;

namespace DataverseLogicDependencyExplorer.Tests
{
    internal class StubScanner : IDependencyScanner
    {
        private readonly Func<ScanContext, CancellationToken, ScanOutcome> body;
        public StubScanner(ScannerKind kind, Func<ScanContext, CancellationToken, ScanOutcome> body) { Kind = kind; this.body = body; }
        public ScannerKind Kind { get; private set; }
        public ScanOutcome Scan(ScanContext c, CancellationToken ct) { return body(c, ct); }
    }

    internal static class Make
    {
        public static DependencyResult Result(ScannerKind src, string name, ConfidenceLevel conf = ConfidenceLevel.StrongDefinitionReference, string id = null, string location = "loc")
        {
            return new DependencyResult
            {
                Source = src, ComponentType = ComponentType.BusinessRule, ComponentName = name, ComponentId = id ?? name, MatchType = "m", MatchedValue = "v",
                Location = location, Confidence = conf, TargetTable = "account", ScanTimestampUtc = DateTime.UtcNow
            };
        }

        public static ScanOutcome Outcome(ScannerKind k, ScanStatus status, params DependencyResult[] results)
        {
            var o = new ScanOutcome { Diagnostic = new ScanDiagnostic { Category = k, Status = status, RecordsExamined = results.Length } };
            o.Results.AddRange(results);
            return o;
        }

        public static ScanRequest Request(SearchMode mode = SearchMode.Table, params ScannerKind[] scopes)
        {
            return new ScanRequest
            {
                Mode = mode, TableLogicalName = "account", TableDisplayName = "Account", TableEntitySetName = "accounts",
                ColumnLogicalName = mode == SearchMode.Column ? "new_status" : null,
                Scopes = new HashSet<ScannerKind>(scopes.Length == 0 ? ScanRequest.AllScopes : scopes),
                EnvironmentName = "Contoso Dev", EnvironmentUrl = "https://contoso.crm.dynamics.com"
            };
        }
    }

    [TestClass]
    public class AnalysisTests
    {
        private static AnalysisResult Run(ScanRequest req, Func<ScannerKind, IDependencyScanner> factory, CancellationToken? ct = null)
        {
            return new DependencyAnalysisService(factory).Analyze(new FakeOrganizationService(), req, null, ct ?? CancellationToken.None);
        }

        [TestMethod]
        public void UnselectedCategories_AreNotInspected_NeverZeroMatches()
        {
            ScanRequest req = Make.Request(SearchMode.Table, ScannerKind.BusinessRules);
            AnalysisResult a = Run(req, k => new StubScanner(k, (c, t) => Make.Outcome(k, ScanStatus.Completed)));
            Assert.AreEqual(6, a.Diagnostics.Count);
            Assert.AreEqual(ScanStatus.Completed, a.DiagnosticFor(ScannerKind.BusinessRules).Status);
            foreach (ScannerKind k in ScanRequest.AllScopes.Where(x => x != ScannerKind.BusinessRules))
            {
                ScanDiagnostic d = a.DiagnosticFor(k);
                Assert.AreEqual(ScanStatus.NotInspected, d.Status);
                Assert.IsFalse(d.WasInspected);
                StringAssert.Contains(d.Message, "NOT a 'no dependencies' result");
            }
        }

        [TestMethod]
        public void OneFailingScanner_DoesNotStopTheOthers_AndEmptyResultIsNotClean()
        {
            AnalysisResult a = Run(Make.Request(), k =>
            {
                if (k == ScannerKind.Plugins) return new StubScanner(k, (c, t) => { throw new InvalidOperationException("kaboom"); });
                return new StubScanner(k, (c, t) => Make.Outcome(k, ScanStatus.Completed));
            });
            Assert.AreEqual(ScanStatus.Failed, a.DiagnosticFor(ScannerKind.Plugins).Status);
            Assert.AreEqual(5, a.Diagnostics.Count(d => d.Status == ScanStatus.Completed));
            Assert.AreEqual(1, a.FailedScans);
            Assert.AreEqual(0, a.MatchCount);
            Assert.IsFalse(a.IsCleanEmptyResult, "a failed category must not look like 'no dependencies'");
        }

        [TestMethod]
        public void AllCategoriesCompleted_WithNoMatches_IsACleanEmptyResult()
        {
            AnalysisResult a = Run(Make.Request(), k => new StubScanner(k, (c, t) => Make.Outcome(k, ScanStatus.Completed)));
            Assert.IsTrue(a.IsCleanEmptyResult);
        }

        [TestMethod]
        public void PartiallySupportedCategory_MakesEmptyResultNotClean()
        {
            AnalysisResult a = Run(Make.Request(), k => new StubScanner(k, (c, t) => Make.Outcome(k, k == ScannerKind.CloudFlows ? ScanStatus.PartiallySupported : ScanStatus.Completed)));
            Assert.IsFalse(a.IsCleanEmptyResult);
            Assert.AreEqual(1, a.UnsupportedOrPartialScans);
        }

        [TestMethod]
        public void Cancellation_MarksRemainingCategoriesNotInspected_AndKeepsPartialResults()
        {
            var cts = new CancellationTokenSource();
            AnalysisResult a = Run(Make.Request(), k => new StubScanner(k, (c, t) =>
            {
                if (k == ScannerKind.JavaScript) cts.Cancel();
                return Make.Outcome(k, ScanStatus.Completed, Make.Result(k, "R-" + k));
            }), cts.Token);
            Assert.IsTrue(a.WasCancelled);
            Assert.AreEqual(ScanStatus.Completed, a.DiagnosticFor(ScannerKind.BusinessRules).Status);
            Assert.AreEqual(ScanStatus.NotInspected, a.DiagnosticFor(ScannerKind.CloudFlows).Status);
            Assert.IsTrue(a.MatchCount >= 2, "results found before cancelling are kept");
            Assert.IsFalse(a.IsCleanEmptyResult);
            Assert.IsTrue(a.GeneralWarnings.Any(w => w.Contains("cancelled")));
        }

        [TestMethod]
        public void OverlappingFindings_AreDeduplicated_KeepingStrongestConfidence()
        {
            var weak = Make.Result(ScannerKind.BusinessRules, "Rule", ConfidenceLevel.TextMatchReviewRequired, "id1");
            var strong = Make.Result(ScannerKind.BusinessRules, "Rule", ConfidenceLevel.ConfirmedMetadataReference, "id1");
            var other = Make.Result(ScannerKind.BusinessRules, "Rule", ConfidenceLevel.StrongDefinitionReference, "id1", location: "different");
            List<DependencyResult> u = ResultDeduplicator.Deduplicate(new[] { weak, strong, other });
            Assert.AreEqual(2, u.Count);
            Assert.AreEqual(ConfidenceLevel.ConfirmedMetadataReference, u.First(r => r.Location == "loc").Confidence);
        }

        [TestMethod]
        public void Analyze_DuplicatesAcrossScanners_AreRemovedFromTheFinalResult()
        {
            AnalysisResult a = Run(Make.Request(SearchMode.Table, ScannerKind.BusinessRules, ScannerKind.ClassicWorkflows),
                k => new StubScanner(k, (c, t) => Make.Outcome(k, ScanStatus.Completed, Make.Result(ScannerKind.BusinessRules, "Same", id: "same"))));
            Assert.AreEqual(1, a.MatchCount);
            Assert.AreEqual(1, a.UniqueComponents);
        }

        [TestMethod]
        public void InvalidRequest_IsRejectedBeforeAnyScanRuns()
        {
            ScanRequest req = Make.Request(SearchMode.Column);
            req.ColumnLogicalName = null;
            Assert.ThrowsException<ArgumentException>(() => Run(req, k => new StubScanner(k, (c, t) => Make.Outcome(k, ScanStatus.Completed))));
        }

        [TestMethod]
        public void Progress_IsReportedAndEndsAt100()
        {
            var seen = new List<ProgressInfo>();
            new DependencyAnalysisService(k => new StubScanner(k, (c, t) => Make.Outcome(k, ScanStatus.Completed)))
                .Analyze(new FakeOrganizationService(), Make.Request(), new SyncProgress(seen), CancellationToken.None);
            Assert.IsTrue(seen.Count >= 6);
            Assert.AreEqual(100, seen.Last().Percent);
        }

        private class SyncProgress : IProgress<ProgressInfo>
        {
            private readonly List<ProgressInfo> list;
            public SyncProgress(List<ProgressInfo> list) { this.list = list; }
            public void Report(ProgressInfo v) { list.Add(v); }
        }

        // ---------------------------------------------------------------- request validation

        [TestMethod]
        public void Validation_GivesActionableMessagesForInvalidCombinations()
        {
            StringAssert.Contains(ScanRequest.Validate(SearchMode.Column, "account", null, 3), "needs a column");
            StringAssert.Contains(ScanRequest.Validate(SearchMode.Column, null, "name", 3), "without its parent table");
            StringAssert.Contains(ScanRequest.Validate(SearchMode.Column, null, null, 3), "table and a column");
            StringAssert.Contains(ScanRequest.Validate(SearchMode.Table, null, null, 3), "Select a table");
            StringAssert.Contains(ScanRequest.Validate(SearchMode.Table, "account", null, 0), "component scope");
            Assert.IsNull(ScanRequest.Validate(SearchMode.Table, "account", null, 1));
            Assert.IsNull(ScanRequest.Validate(SearchMode.Table, "account", "name", 1), "a column is ignored in table mode");
            Assert.IsNull(ScanRequest.Validate(SearchMode.Column, "account", "name", 1));
        }

        // ---------------------------------------------------------------- metadata mapping / filtering

        [TestMethod]
        public void MetadataMapping_ExcludesInternalTables_AndFlagsCustomOnes()
        {
            Assert.IsNull(EntityMetadataService.MapTable(null, "x", false, false, false, false, null));
            Assert.IsNull(EntityMetadataService.MapTable("priv", null, false, false, true, false, null), "private");
            Assert.IsNull(EntityMetadataService.MapTable("logic", null, false, false, false, true, null), "logical");
            Assert.IsNull(EntityMetadataService.MapTable("a_b", null, false, true, false, false, null), "intersect");
            TableInfo custom = EntityMetadataService.MapTable("new_project", "Project", true, false, false, false, "new_projects");
            Assert.IsTrue(custom.IsCustom);
            Assert.AreEqual("Project (new_project)", custom.Title);
            Assert.AreEqual("new_projects", custom.EntitySetName);
            TableInfo std = EntityMetadataService.MapTable("account", null, null, null, null, null, "accounts");
            Assert.IsFalse(std.IsCustom);
            Assert.AreEqual("account", std.Title);
        }

        [TestMethod]
        public void MetadataMapping_ExcludesShadowColumns_AndMissingNames()
        {
            Assert.IsNull(EntityMetadataService.MapColumn("owneridname", "Owner", "String", false, "ownerid"));
            Assert.IsNull(EntityMetadataService.MapColumn("", null, null, null, null));
            ColumnInfo c = EntityMetadataService.MapColumn("new_status", "Status", "Picklist", true, null);
            Assert.IsTrue(c.IsCustom);
            Assert.AreEqual("Status (new_status)", c.Title);
        }

        [TestMethod]
        public void TableAndColumnFilters_SearchNamesAndSeparateCustomFromStandard()
        {
            var tables = new List<TableInfo>
            {
                EntityMetadataService.MapTable("account", "Account", false, false, false, false, "accounts"),
                EntityMetadataService.MapTable("new_project", "Project", true, false, false, false, "new_projects"),
                EntityMetadataService.MapTable("contact", "Contact", false, false, false, false, "contacts")
            };
            Assert.AreEqual(3, EntityMetadataService.FilterTables(tables, "", TableKindFilter.All).Count);
            Assert.AreEqual(1, EntityMetadataService.FilterTables(tables, null, TableKindFilter.Custom).Count);
            Assert.AreEqual(2, EntityMetadataService.FilterTables(tables, "", TableKindFilter.Standard).Count);
            Assert.AreEqual("contact", EntityMetadataService.FilterTables(tables, "CONT", TableKindFilter.All).Single().LogicalName);
            Assert.AreEqual("new_project", EntityMetadataService.FilterTables(tables, "project", TableKindFilter.All).Single().LogicalName);
            var cols = new List<ColumnInfo> { EntityMetadataService.MapColumn("name", "Account Name", "String", false, null), EntityMetadataService.MapColumn("new_x", null, "Integer", true, null) };
            Assert.AreEqual("name", EntityMetadataService.FilterColumns(cols, "account").Single().LogicalName);
            Assert.AreEqual(0, EntityMetadataService.FilterColumns(cols, "zzz").Count);
        }

        [TestMethod]
        public void MetadataService_CacheCanBeInvalidated()
        {
            var svc = new EntityMetadataService();
            Assert.IsFalse(svc.HasTables);
            Assert.AreEqual(0, svc.CachedTables.Count);
            svc.Invalidate();
            Assert.IsNull(svc.FindTable("account"));
        }

        // ---------------------------------------------------------------- result filter

        private static AnalysisResult Sample()
        {
            var a = new AnalysisResult { Request = Make.Request(SearchMode.Column) };
            var r1 = Make.Result(ScannerKind.BusinessRules, "Alpha rule", ConfidenceLevel.StrongDefinitionReference, "1"); r1.SolutionName = "CoreSol";
            var r2 = Make.Result(ScannerKind.JavaScript, "new_/js/beta.js", ConfidenceLevel.TextMatchReviewRequired, "2");
            r2.ComponentType = ComponentType.JavaScriptWebResource; r2.Evidence = "getAttribute('new_status')"; r2.TargetColumn = "new_status"; r2.SolutionName = "OtherSol, CoreSol";
            a.Results.AddRange(new[] { r1, r2 });
            a.Diagnostics.Add(new ScanDiagnostic { Category = ScannerKind.BusinessRules, Status = ScanStatus.Completed });
            a.Diagnostics.Add(new ScanDiagnostic { Category = ScannerKind.JavaScript, Status = ScanStatus.CompletedWithWarnings });
            return a;
        }

        [TestMethod]
        public void ResultFilter_FiltersByEveryDimension()
        {
            AnalysisResult a = Sample();
            Assert.AreEqual(2, ResultFilter.Apply(a, new FilterCriteria()).Count);
            Assert.AreEqual(1, ResultFilter.Apply(a, new FilterCriteria { ComponentType = ComponentType.BusinessRule }).Count);
            Assert.AreEqual("Alpha rule", ResultFilter.Apply(a, new FilterCriteria { ComponentName = "alpha" }).Single().ComponentName);
            Assert.AreEqual(1, ResultFilter.Apply(a, new FilterCriteria { MatchedTableOrColumn = "new_status" }).Count);
            Assert.AreEqual(1, ResultFilter.Apply(a, new FilterCriteria { Confidence = ConfidenceLevel.TextMatchReviewRequired }).Count);
            Assert.AreEqual(1, ResultFilter.Apply(a, new FilterCriteria { ScanStatus = ScanStatus.CompletedWithWarnings }).Count);
            Assert.AreEqual(2, ResultFilter.Apply(a, new FilterCriteria { Solution = "CoreSol" }).Count);
            Assert.AreEqual(1, ResultFilter.Apply(a, new FilterCriteria { Solution = "OtherSol" }).Count);
            Assert.AreEqual("new_/js/beta.js", ResultFilter.Apply(a, new FilterCriteria { SearchText = "getattribute" }).Single().ComponentName);
            Assert.AreEqual(0, ResultFilter.Apply(a, new FilterCriteria { SearchText = "nothing-like-this" }).Count);
            Assert.AreEqual("CoreSol,OtherSol", string.Join(",", ResultFilter.DistinctSolutions(a)));
            Assert.AreEqual(0, ResultFilter.Apply(null, null).Count);
        }
    }
}

namespace DataverseLogicDependencyExplorer.Tests
{
    [Microsoft.VisualStudio.TestTools.UnitTesting.TestClass]
    public class PresentationTests
    {
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void DetailFormatter_ShowsEvidenceConfidenceExplanationAndLimits()
        {
            var r = Make.Result(DataverseLogicDependencyExplorer.Models.ScannerKind.JavaScript, "new_/js/a.js", DataverseLogicDependencyExplorer.Models.ConfidenceLevel.TextMatchReviewRequired, "id-1");
            r.ComponentType = DataverseLogicDependencyExplorer.Models.ComponentType.JavaScriptWebResource;
            r.TargetColumn = "new_status"; r.ParentTable = "account"; r.LineNumber = 12; r.Evidence = "f.getAttribute('new_status')"; r.SourceResource = "new_/js/a.js";
            r.Explanation = "because"; r.Limitation = "static only"; r.SolutionName = "Core";
            var d = new DataverseLogicDependencyExplorer.Models.ScanDiagnostic { Category = DataverseLogicDependencyExplorer.Models.ScannerKind.JavaScript, Status = DataverseLogicDependencyExplorer.Models.ScanStatus.CompletedWithWarnings };
            d.Limit("Static text analysis only"); d.Warn("one warning");
            string text = DataverseLogicDependencyExplorer.Services.ResultDetailFormatter.Format(r, d);
            foreach (string needle in new[] { "Component: new_/js/a.js", "Component ID: id-1", "Parent table", "Matched column: new_status", "Line number: 12", "File / resource name: new_/js/a.js",
                "Confidence: Text match - review required", "Why this result was included: because", "Technical limitation of this finding: static only", "Evidence:", "getAttribute('new_status')",
                "Scan status of this category: Completed with warnings", "Static text analysis only", "one warning", "Solution: Core" })
                Microsoft.VisualStudio.TestTools.UnitTesting.StringAssert.Contains(text, needle);
        }

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void SummaryLine_ReportsExaminedMatchesAndFailures()
        {
            var a = new DataverseLogicDependencyExplorer.Models.AnalysisResult { Request = Make.Request() };
            a.Diagnostics.Add(new DataverseLogicDependencyExplorer.Models.ScanDiagnostic { Category = DataverseLogicDependencyExplorer.Models.ScannerKind.Plugins, Status = DataverseLogicDependencyExplorer.Models.ScanStatus.Failed, RecordsExamined = 0 });
            a.Diagnostics.Add(new DataverseLogicDependencyExplorer.Models.ScanDiagnostic { Category = DataverseLogicDependencyExplorer.Models.ScannerKind.Forms, Status = DataverseLogicDependencyExplorer.Models.ScanStatus.Completed, RecordsExamined = 4 });
            a.Results.Add(Make.Result(DataverseLogicDependencyExplorer.Models.ScannerKind.Forms, "F"));
            string s = DataverseLogicDependencyExplorer.Services.AnalysisSummary.Line(a);
            Microsoft.VisualStudio.TestTools.UnitTesting.StringAssert.Contains(s, "Examined 4");
            Microsoft.VisualStudio.TestTools.UnitTesting.StringAssert.Contains(s, "Matches found: 1");
            Microsoft.VisualStudio.TestTools.UnitTesting.StringAssert.Contains(s, "Scan failures: 1");
            Microsoft.VisualStudio.TestTools.UnitTesting.StringAssert.Contains(DataverseLogicDependencyExplorer.Services.AnalysisSummary.Line(null), "No scan");
        }
    }
}


namespace DataverseLogicDependencyExplorer.Tests
{
    [Microsoft.VisualStudio.TestTools.UnitTesting.TestClass]
    public class UiLayoutTests
    {
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void LeftPane_Is27Percent_Clamped280To350_AndOnlyAppliedFrom700px()
        {
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsNull(UiLayout.LeftPaneWidth(699, 5));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(280, UiLayout.LeftPaneWidth(700, 5).Value);   // 27% = 189 -> clamped up
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(297, UiLayout.LeftPaneWidth(1100, 5).Value);  // 27% of 1100
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(350, UiLayout.LeftPaneWidth(2000, 5).Value);  // clamped down
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(UiLayout.LeftPaneWidth(750, 5).Value <= 750 - 400 - 5, "right pane keeps its minimum");
        }

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void HeaderAndCardWidths_FollowTheSpecFormulas()
        {
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(230, UiLayout.HeaderInnerWidth(100));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(276, UiLayout.HeaderInnerWidth(300));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(85, UiLayout.HalfButtonWidth(100));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(134, UiLayout.HalfButtonWidth(276));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(215, UiLayout.CardWidth(200));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(275, UiLayout.CardWidth(300));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(200, UiLayout.StatusLabelWidth(150));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(480, UiLayout.StatusLabelWidth(500));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(176, UiLayout.LeftHeaderHeight);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(78, UiLayout.RightHeaderHeight);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(62, UiLayout.CardHeight);
        }

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void MultiLineText_IsCollapsedToOneLine_WithTheSeparator()
        {
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("one  \u203A  two  \u203A  three", UiLayout.Collapse("one\r\n two \n\n three "));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("single", UiLayout.Collapse("single"));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsNull(UiLayout.Collapse(null));

            var r = Make.Result(DataverseLogicDependencyExplorer.Models.ScannerKind.Plugins, "Step");
            r.MatchType = "Filtering attribute\nImage attribute";
            var row = DataverseLogicDependencyExplorer.ViewModels.ResultRow.From(r, null);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("Filtering attribute  \u203A  Image attribute", row.MatchType);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("Filtering attribute\nImage attribute", row.Full["MatchType"], "full text kept for the tooltip");
        }
    }
}
