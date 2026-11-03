// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DataverseLogicDependencyExplorer.Models;
using DataverseLogicDependencyExplorer.Reporting;

namespace DataverseLogicDependencyExplorer.Tests
{
    [TestClass]
    public class ExcelReportTests
    {
        private static readonly XNamespace M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private string dir;

        [TestInitialize] public void Init() { dir = Path.Combine(Path.GetTempPath(), "dlde_" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(dir); }
        [TestCleanup] public void Clean() { try { Directory.Delete(dir, true); } catch { } }

        // --- tiny xlsx reader (zip + xml only; independent of the writer)
        private class Book
        {
            public ZipArchive Zip; public List<string> Shared = new List<string>(); public List<string> SheetNames = new List<string>();
            public XDocument Sheet(int i) { using (var s = Zip.GetEntry("xl/worksheets/sheet" + (i + 1) + ".xml").Open()) return XDocument.Load(s); }
            public string[][] Rows(int i)
            {
                var rows = new List<string[]>();
                foreach (XElement row in Sheet(i).Descendants(M + "row"))
                {
                    var cells = new List<string>();
                    foreach (XElement c in row.Elements(M + "c"))
                    {
                        int col = 0; foreach (char ch in new string(c.Attribute("r").Value.TakeWhile(char.IsLetter).ToArray())) col = col * 26 + (ch - 'A' + 1);
                        while (cells.Count < col - 1) cells.Add(null);
                        XElement v = c.Element(M + "v");
                        cells.Add(v == null ? null : (string)c.Attribute("t") == "s" ? Shared[int.Parse(v.Value)] : v.Value);
                    }
                    rows.Add(cells.ToArray());
                }
                return rows.ToArray();
            }
        }

        private static Book Open(string path)
        {
            var b = new Book { Zip = new ZipArchive(new MemoryStream(File.ReadAllBytes(path)), ZipArchiveMode.Read) };
            using (var s = b.Zip.GetEntry("xl/sharedStrings.xml").Open())
                b.Shared = XDocument.Load(s).Descendants(M + "t").Select(t => t.Value).ToList();
            using (var s = b.Zip.GetEntry("xl/workbook.xml").Open())
                b.SheetNames = XDocument.Load(s).Descendants(M + "sheet").Select(x => (string)x.Attribute("name")).ToList();
            return b;
        }

        private static AnalysisResult Analysis(bool withResults, bool withProblems)
        {
            var a = new AnalysisResult { Request = Make.Request(SearchMode.Column), StartedUtc = new DateTime(2026, 10, 27, 8, 30, 0, DateTimeKind.Utc), FinishedUtc = new DateTime(2026, 10, 27, 8, 31, 0, DateTimeKind.Utc) };
            foreach (ScannerKind k in ScanRequest.AllScopes)
            {
                var d = new ScanDiagnostic { Category = k, Status = ScanStatus.Completed, RecordsExamined = 3 };
                d.Limit("Limit of " + k);
                if (withProblems && k == ScannerKind.CloudFlows) { d.Status = ScanStatus.PartiallySupported; d.Warn("Cloud flows are partial"); }
                if (withProblems && k == ScannerKind.Plugins) { d.Status = ScanStatus.PermissionDenied; d.PermissionLimited = true; d.Error("no privilege"); d.RecordsExamined = 0; }
                a.Diagnostics.Add(d);
            }
            if (withResults)
            {
                var r = Make.Result(ScannerKind.BusinessRules, "=HYPERLINK(\"http://x\")", ConfidenceLevel.StrongDefinitionReference, Guid.NewGuid().ToString());
                r.Evidence = "line1\nline2 with <xml> & \"quotes\" \u00e9\u4e2d\u6587\u0001"; r.LineNumber = 7; r.SolutionName = "Sol";
                a.Results.Add(r);
                a.Results.Add(Make.Result(ScannerKind.JavaScript, "x.js", ConfidenceLevel.TextMatchReviewRequired));
            }
            if (withProblems) a.GeneralWarnings.Add("Solution information could not be read");
            return a;
        }

        [TestMethod]
        public void Workbook_HasFourSheetsInRequiredOrder()
        {
            string p = Path.Combine(dir, "a.xlsx");
            ExcelReportExporter.Export(Analysis(true, false), p, "1.2026.11.11");
            Book b = Open(p);
            Assert.AreEqual("Summary,Dependencies,Scan Diagnostics,Scan Configuration", string.Join(",", b.SheetNames));
        }

        [TestMethod]
        public void DependenciesSheet_HasHeaderRowFrozenFilteredAndWrapsEvidence()
        {
            string p = Path.Combine(dir, "a.xlsx");
            ExcelReportExporter.Export(Analysis(true, false), p, "1.2026.11.11");
            Book b = Open(p);
            XDocument sheet = b.Sheet(1);
            Assert.AreEqual("frozen", (string)sheet.Descendants(M + "pane").Single().Attribute("state"));
            Assert.AreEqual("1", (string)sheet.Descendants(M + "pane").Single().Attribute("ySplit"));
            Assert.IsTrue(((string)sheet.Descendants(M + "autoFilter").Single().Attribute("ref")).StartsWith("A1:"));
            string[][] rows = b.Rows(1);
            Assert.AreEqual(3, rows.Length);
            Assert.AreEqual("Component Type", rows[0][0]);
            Assert.AreEqual("Confidence", rows[0][11]);
            Assert.AreEqual("Strong definition reference", rows[1][11]);
            Assert.AreEqual("7", rows[1][15]);
            Assert.IsTrue(sheet.Descendants(M + "cols").Single().Elements(M + "col").Count() >= 20, "column widths set");
            Assert.IsTrue(sheet.Descendants(M + "c").Any(c => (string)c.Attribute("s") == "3"), "date cells use the datetime style");
        }

        [TestMethod]
        public void SpecialText_IsStoredAsTextNotFormula_AndInvalidXmlCharsAreRemoved()
        {
            string p = Path.Combine(dir, "a.xlsx");
            ExcelReportExporter.Export(Analysis(true, false), p, "1.2026.11.11");
            Book b = Open(p);
            string[][] rows = b.Rows(1);
            Assert.AreEqual("=HYPERLINK(\"http://x\")", rows[1][1]);
            Assert.IsFalse(b.Sheet(1).Descendants(M + "f").Any(), "no formula elements");
            string evidence = rows[1][12];
            StringAssert.Contains(evidence, "<xml> & \"quotes\" \u00e9\u4e2d\u6587");
            Assert.IsFalse(evidence.Contains("\u0001"));
            StringAssert.Contains(evidence, "\n");
        }

        [TestMethod]
        public void SummarySheet_ContainsRequiredFacts()
        {
            string p = Path.Combine(dir, "a.xlsx");
            ExcelReportExporter.Export(Analysis(true, true), p, "1.2026.11.11");
            string all = string.Join("\n", Open(p).Rows(0).Select(r => string.Join(" | ", r.Select(x => x ?? ""))));
            foreach (string needle in new[] { "https://contoso.crm.dynamics.com", "Search mode", "Selected table", "Selected column", "new_status", "Total matches", "Unique components",
                "Matches by component type", "Matches by confidence classification", "Confirmed metadata reference", "Strong definition reference", "Text match - review required",
                "Possible or dynamic dependency", "Scans completed", "Scans failed", "Important limitations and warnings", "Compiled plugin code is never inspected", "Maintainer: Gagan Deep Singh", "glowofsky",
                "Cloud flows are partial", "no privilege", "Solution information could not be read", "Permission denied" })
                StringAssert.Contains(all, needle);
        }

        [TestMethod]
        public void DiagnosticsSheet_ShowsNotInspectedAndPermissionDetails()
        {
            AnalysisResult a = Analysis(false, true);
            var skipped = a.Diagnostics.First(d => d.Category == ScannerKind.Forms);
            skipped.Status = ScanStatus.NotInspected; skipped.RecordsExamined = 0; skipped.Warn("Not inspected: not selected");
            string p = Path.Combine(dir, "a.xlsx");
            ExcelReportExporter.Export(a, p, "1.2026.11.11");
            string[][] rows = Open(p).Rows(2);
            string[] header = rows[0];
            Assert.AreEqual("Scan Status", header[1]);
            string[] forms = rows.First(r => r[0] == "Forms and Form Event Registrations");
            Assert.AreEqual("Not inspected", forms[1]);
            Assert.AreEqual("No", forms[2]);
            Assert.IsNull(forms[3], "records examined must be blank, not 0, when not inspected");
            string[] plugins = rows.First(r => r[0] == "Plugin Registrations");
            Assert.AreEqual("Permission denied", plugins[1]);
            Assert.AreEqual("Yes", plugins[6]);
            Assert.IsTrue(rows.Any(r => r[0] == "(run-wide)"));
        }

        [TestMethod]
        public void ConfigurationSheet_RecordsTargetScopeAndVersions()
        {
            string p = Path.Combine(dir, "a.xlsx");
            ExcelReportExporter.Export(Analysis(false, false), p, "1.2026.11.11");
            string all = string.Join("\n", Open(p).Rows(3).Select(r => string.Join("=", r)));
            foreach (string needle in new[] { "Report version=1.2026.11.11", "Plugin version=1.2026.11.11", "Target table (logical name)=account", "Target column (logical name)=new_status", "Business Rules", "Scan timestamp (UTC)", "Maintainer=Gagan Deep Singh", "LinkedIn=https://www.linkedin.com/in/glowofsky/", "Project URL=https://github.com/theGlowOfSky/DataverseLogicDependencyExplorer", "Host=XrmToolBox", "Data platform=Microsoft Dataverse", "Export format=Excel workbook (.xlsx)" })
                StringAssert.Contains(all, needle);
        }

        [TestMethod]
        public void Export_WithNoResults_StillProducesAllSheets_AndSaysWhetherItIsClean()
        {
            string p = Path.Combine(dir, "empty.xlsx");
            ExcelReportExporter.Export(Analysis(false, false), p, "1.2026.11.11");
            Book b = Open(p);
            Assert.AreEqual(4, b.SheetNames.Count);
            Assert.AreEqual(1, b.Rows(1).Length, "header only");
            StringAssert.Contains(string.Join("\n", b.Rows(0).Select(r => string.Join(" ", r.Select(x => x ?? "")))), "No references were found. Every selected category was inspected.");

            string p2 = Path.Combine(dir, "empty2.xlsx");
            ExcelReportExporter.Export(Analysis(false, true), p2, "1.2026.11.11");
            StringAssert.Contains(string.Join("\n", Open(p2).Rows(0).Select(r => string.Join(" ", r.Select(x => x ?? "")))), "This is not a clean result");
        }

        [TestMethod]
        public void Export_WithWarningsAndPartialResults_IncludesThem()
        {
            AnalysisResult a = Analysis(true, true);
            a.WasCancelled = true;
            string p = Path.Combine(dir, "partial.xlsx");
            ExcelReportExporter.Export(a, p, "1.2026.11.11");
            Book b = Open(p);
            Assert.AreEqual(3, b.Rows(1).Length);
            StringAssert.Contains(string.Join(" ", b.Rows(0).SelectMany(r => r.Select(x => x ?? ""))), "CANCELLED");
        }

        [TestMethod]
        public void LongText_IsCutWithVisibleMarker_AndCounted()
        {
            AnalysisResult a = Analysis(true, false);
            a.Results[0].Evidence = new string('x', 40000);
            string p = Path.Combine(dir, "long.xlsx");
            ExportOutcome o = ExcelReportExporter.Export(a, p, "1.2026.11.11");
            Assert.AreEqual(1, o.TruncatedCells);
            string ev = Open(p).Rows(1)[1][12];
            Assert.AreEqual(32767, ev.Length);
            StringAssert.Contains(ev, "TRUNCATED");
        }

        [TestMethod]
        public void FileErrors_ThrowCatchableExceptions_AndNeverCorruptState()
        {
            AnalysisResult a = Analysis(true, false);
            Assert.ThrowsException<ArgumentException>(() => ExcelReportExporter.Export(a, " ", "1.2026.11.11"));
            string missing = Path.Combine(dir, "no_such_folder", "x.xlsx");
            Assert.ThrowsException<DirectoryNotFoundException>(() => ExcelReportExporter.Export(a, missing, "1.2026.11.11"));
            string locked = Path.Combine(dir, "locked.xlsx");
            using (new FileStream(locked, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
                Assert.ThrowsException<IOException>(() => ExcelReportExporter.Export(a, locked, "1.2026.11.11"));
        }

        [TestMethod]
        public void SheetNames_AreSanitisedAndUnique_AndColumnLettersExceedZ()
        {
            var wb = new XlsxWorkbookWriter();
            Assert.AreEqual("a-b-c", wb.AddSheet("a/b:c").Name);
            Assert.AreEqual(31, wb.AddSheet(new string('n', 50)).Name.Length);
            string dup1 = wb.AddSheet("Same").Name, dup2 = wb.AddSheet("same").Name;
            Assert.AreNotEqual(dup1, dup2);
            Assert.AreEqual("A", XlsxWorkbookWriter.ColumnName(0));
            Assert.AreEqual("Z", XlsxWorkbookWriter.ColumnName(25));
            Assert.AreEqual("AA", XlsxWorkbookWriter.ColumnName(26));
            Assert.AreEqual("AB", XlsxWorkbookWriter.ColumnName(27));
            Assert.AreEqual("AAA", XlsxWorkbookWriter.ColumnName(702));
        }

        [TestMethod]
        public void DefaultFileName_FollowsTheRequestedPattern()
        {
            Assert.AreEqual("DataverseLogicDependencyReport_20261027_083005.xlsx", ExcelReportExporter.DefaultFileName(new DateTime(2026, 10, 27, 8, 30, 5)));
        }
    }
}
