// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DataverseLogicDependencyExplorer.Services;

namespace DataverseLogicDependencyExplorer.Tests
{
    [TestClass]
    public class JsSourceAnalyzerTests
    {
        [TestMethod]
        public void ApiCallArgument_IsFoundWithAccurateLineNumber()
        {
            string js = "function onLoad(ctx) {\r\n  var f = ctx.getFormContext();\r\n  var a = f.getAttribute(\"new_status\");\r\n}";
            JsAnalysis r = JsSourceAnalyzer.FindColumn(js, "new_status", false);
            Assert.AreEqual(1, r.Matches.Count);
            Assert.AreEqual(JsMatchKind.ApiCallArgument, r.Matches[0].Kind);
            Assert.AreEqual("getAttribute", r.Matches[0].ApiName);
            Assert.AreEqual(3, r.Matches[0].Line);
        }

        [TestMethod]
        public void CommentOnlyMatch_IsNotReportedAsReference_ButIsCounted()
        {
            string js = "// uses new_status somewhere\n/* new_status again */\nvar x = 1;";
            JsAnalysis r = JsSourceAnalyzer.FindColumn(js, "new_status", false);
            Assert.AreEqual(0, r.Matches.Count);
            Assert.AreEqual(2, r.CommentOnlyMatches);
        }

        [TestMethod]
        public void CommentMatches_AreListedWhenRequested()
        {
            JsAnalysis r = JsSourceAnalyzer.FindColumn("// new_status\n", "new_status", true);
            Assert.AreEqual(1, r.Matches.Count);
            Assert.AreEqual(JsMatchKind.Comment, r.Matches[0].Kind);
        }

        [TestMethod]
        public void LongerIdentifier_IsNotAMatch()
        {
            string js = "var new_status_reason = 1; var x_new_status = 2; var a = form.getAttribute('new_status_reason');";
            JsAnalysis r = JsSourceAnalyzer.FindColumn(js, "new_status", false);
            Assert.AreEqual(0, r.Matches.Count);
        }

        [TestMethod]
        public void BareIdentifierInCode_IsAMatchForDistinctiveNames()
        {
            JsAnalysis r = JsSourceAnalyzer.FindColumn("var v = record.new_status;", "new_status", false);
            Assert.AreEqual(1, r.Matches.Count);
            Assert.AreEqual(JsMatchKind.Identifier, r.Matches[0].Kind);
        }

        [TestMethod]
        public void GenericColumnName_InProseIsIgnored_ButApiCallIsFound()
        {
            string js = "alert('Name is required'); var o = {name: 1}; var a = f.getAttribute('name');";
            JsAnalysis r = JsSourceAnalyzer.FindColumn(js, "name", false);
            Assert.AreEqual(1, r.Matches.Count);
            Assert.AreEqual(JsMatchKind.ApiCallArgument, r.Matches[0].Kind);
        }

        [TestMethod]
        public void QueryTokens_AreRecognised()
        {
            string js = "var u = \"?$select=name,new_status&$filter=new_status eq 1\"; var fx = \"<attribute name='new_status' />\"; var l = r._new_status_value;";
            JsAnalysis r = JsSourceAnalyzer.FindColumn(js, "new_status", false);
            Assert.IsTrue(r.Matches.Count(m => m.Kind == JsMatchKind.QueryOrFetchToken) >= 3, "select/filter/fetch tokens");
        }

        [TestMethod]
        public void LookupValueForm_IsMatched()
        {
            JsAnalysis r = JsSourceAnalyzer.FindColumn("var id = rec[\"_new_owner_value\"];", "new_owner", false);
            Assert.AreEqual(1, r.Matches.Count);
        }

        [TestMethod]
        public void RegexLiteralAndTemplates_DoNotCrashOrLeak()
        {
            string js = "var re = /new_status\\/x/g; var t = `a ${b} new_status`; var s = 4 / 2; var c = f.getAttribute('new_status');";
            JsAnalysis r = JsSourceAnalyzer.FindColumn(js, "new_status", false);
            Assert.IsTrue(r.Matches.Any(m => m.Kind == JsMatchKind.ApiCallArgument));
            Assert.IsFalse(r.Matches.Any(m => m.Index < js.IndexOf("var t")), "regex literal content must not match");
        }

        [TestMethod]
        public void MalformedOrEmptySource_ReturnsNoMatchesWithoutThrowing()
        {
            Assert.AreEqual(0, JsSourceAnalyzer.FindColumn(null, "new_x", false).Matches.Count);
            Assert.AreEqual(0, JsSourceAnalyzer.FindColumn("", "new_x", false).Matches.Count);
            JsAnalysis r = JsSourceAnalyzer.FindColumn("var a = 'unterminated new_x\nvar b = f.getAttribute('new_x'", "new_x", false);
            Assert.IsNotNull(r);
        }

        [TestMethod]
        public void TableSearch_FindsApiAndUrlButNotBareWord()
        {
            string js = "Xrm.WebApi.retrieveRecord(\"new_project\", id); var u = \"/api/data/v9.2/new_projects?$top=1\"; var new_project_count = 1; // new_project";
            JsAnalysis r = JsSourceAnalyzer.FindTable(js, "new_project", "new_projects", false);
            Assert.IsTrue(r.Matches.Any(m => m.Kind == JsMatchKind.ApiCallArgument && m.ApiName == "retrieveRecord"));
            Assert.IsTrue(r.Matches.Any(m => m.MatchedText == "new_projects"));
            Assert.IsFalse(r.Matches.Any(m => m.Kind == JsMatchKind.Comment));
        }

        [TestMethod]
        public void TableSearch_GenericWordInCodeIsIgnored()
        {
            JsAnalysis r = JsSourceAnalyzer.FindTable("var account = getAccount(); show(account);", "account", "accounts", false);
            Assert.AreEqual(0, r.Matches.Count);
        }

        [TestMethod]
        public void DynamicAccess_IsDetected_ButLiteralCallsAreNot()
        {
            string js = "f.getAttribute('new_a');\nf.getAttribute(fieldName);\nf.getAttribute('new_' + suffix);\n// f.getAttribute(x)\nvar s = \"getAttribute(y)\";";
            DynamicAccessInfo d = JsSourceAnalyzer.FindDynamicAccess(js);
            Assert.AreEqual(2, d.Count);
            Assert.AreEqual(2, d.FirstLine);
        }

        [TestMethod]
        public void FunctionBody_IsLocated_WithBracesInStringsIgnored()
        {
            string js = "var NS = NS || {};\nNS.Account = {\n  onLoad: function (ctx) {\n    var s = \"}\"; // }\n    var a = f.getAttribute('new_x');\n  },\n  onSave: function () { return 1; }\n};";
            int s, e;
            Assert.IsTrue(JsSourceAnalyzer.TryFindFunctionBody(js, "NS.Account.onLoad", out s, out e));
            string body = js.Substring(s, e - s + 1);
            StringAssert.Contains(body, "new_x");
            Assert.IsFalse(body.Contains("onSave"));
        }

        [TestMethod]
        public void FunctionBody_NotFound_ReturnsFalse()
        {
            int s, e;
            Assert.IsFalse(JsSourceAnalyzer.TryFindFunctionBody("var a = 1;", "missing", out s, out e));
        }

        [TestMethod]
        public void LineMap_HandlesMixedLineEndings()
        {
            var map = new LineMap("a\r\nb\nc\rd");
            Assert.AreEqual(1, map.LineOf(0));
            Assert.AreEqual(2, map.LineOf(3));
            Assert.AreEqual(3, map.LineOf(5));
            Assert.AreEqual(4, map.LineOf(7));
        }
    }
}
