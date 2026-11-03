// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Linq;
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
    public class ScannerTests
    {
        private static ScanContext Ctx(FakeOrganizationService svc, SearchMode mode, string table = "account", string column = "new_status", Action<ScanRequest> tweak = null)
        {
            var req = new ScanRequest
            {
                Mode = mode, TableLogicalName = table, TableEntitySetName = table + "s",
                ColumnLogicalName = mode == SearchMode.Column ? column : null,
                Scopes = new System.Collections.Generic.HashSet<ScannerKind>(ScanRequest.AllScopes)
            };
            if (tweak != null) tweak(req);
            return new ScanContext { Service = svc, Request = req, TimestampUtc = DateTime.UtcNow, WebResources = new WebResourceRepository(svc), Forms = new FormRepository(svc) };
        }

        private static bool HasCondition(QueryExpression q, string attr, object value)
        {
            return q.Criteria.Conditions.Any(c => c.AttributeName == attr && c.Values.Contains(value));
        }

        // ------------------------------------------------------------ status mapping

        [TestMethod]
        public void Scanner_PermissionDenied_IsReportedAsPermissionDenied_NotAsNoMatches()
        {
            var svc = new FakeOrganizationService().OnThrow("workflow", F.PermissionFault());
            ScanOutcome o = new BusinessRuleScanner().Scan(Ctx(svc, SearchMode.Table), CancellationToken.None);
            Assert.AreEqual(ScanStatus.PermissionDenied, o.Diagnostic.Status);
            Assert.IsTrue(o.Diagnostic.PermissionLimited);
            Assert.IsFalse(o.Diagnostic.WasInspected);
            Assert.AreEqual(0, o.Results.Count);
        }

        [TestMethod]
        public void Scanner_GenericFailure_IsReportedAsFailed()
        {
            var svc = new FakeOrganizationService().OnThrow("workflow", new InvalidOperationException("boom"));
            ScanOutcome o = new ClassicWorkflowScanner().Scan(Ctx(svc, SearchMode.Table), CancellationToken.None);
            Assert.AreEqual(ScanStatus.Failed, o.Diagnostic.Status);
            StringAssert.Contains(o.Diagnostic.Message, "boom");
        }

        [TestMethod]
        public void Scanner_Cancellation_IsReportedAsCancelled()
        {
            var cts = new CancellationTokenSource();
            var svc = new FakeOrganizationService().On("workflow", F.Row("workflow", null, "name", "R"));
            cts.Cancel();
            ScanOutcome o = new BusinessRuleScanner().Scan(Ctx(svc, SearchMode.Table), cts.Token);
            Assert.AreEqual(ScanStatus.Cancelled, o.Diagnostic.Status);
        }

        // ------------------------------------------------------------ business rules

        private static Entity Rule(string name, string js, string xaml = null, string table = "account")
        {
            var e = F.Row("workflow", null, "name", name, "primaryentity", table, "statecode", new OptionSetValue(1), "ismanaged", false,
                          "clientdata", "<clientdata><clientcode><![CDATA[" + js + "]]></clientcode></clientdata>");
            if (xaml != null) e["xaml"] = xaml;
            return e;
        }

        [TestMethod]
        public void BusinessRule_TableSearch_UsesTargetedQuery_AndConfirmedMetadata()
        {
            var svc = new FakeOrganizationService().On("workflow", Rule("Rule A", "var v1 = f.data.entity.attributes.get('new_status');"));
            ScanOutcome o = new BusinessRuleScanner().Scan(Ctx(svc, SearchMode.Table), CancellationToken.None);
            Assert.AreEqual(ScanStatus.Completed, o.Diagnostic.Status);
            Assert.AreEqual(1, o.Results.Count);
            Assert.AreEqual(ConfidenceLevel.ConfirmedMetadataReference, o.Results[0].Confidence);
            StringAssert.Contains(o.Results[0].AdditionalInfo, "new_status");
            QueryExpression q = svc.Queries[0];
            Assert.IsTrue(HasCondition(q, "category", 2), "business rules only");
            Assert.IsTrue(HasCondition(q, "primaryentity", "account"), "targeted to the table");
            Assert.IsNotNull(q.PageInfo);
        }

        [TestMethod]
        public void BusinessRule_ColumnSearch_FindsOnlyRulesThatUseTheColumn()
        {
            var svc = new FakeOrganizationService().On("workflow",
                Rule("Uses", "var v1 = f.data.entity.attributes.get('new_status');\nif (v1.getValue() == 1) {\n v1.setRequiredLevel('required');\n}"),
                Rule("Does not", "var v1 = f.data.entity.attributes.get('new_other');"));
            ScanOutcome o = new BusinessRuleScanner().Scan(Ctx(svc, SearchMode.Column), CancellationToken.None);
            Assert.AreEqual(1, o.Results.Count);
            DependencyResult r = o.Results[0];
            Assert.AreEqual("Uses", r.ComponentName);
            Assert.AreEqual(ConfidenceLevel.StrongDefinitionReference, r.Confidence);
            Assert.AreEqual("Condition and action reference", r.MatchType);
            Assert.AreEqual(2, r.RecordsLine());
            Assert.AreEqual(2, o.Diagnostic.RecordsExamined);
        }

        [TestMethod]
        public void BusinessRule_MalformedXaml_IsWarnedAndFallsBackToTextMatch()
        {
            var svc = new FakeOrganizationService().On("workflow", Rule("R", "var x = 1;", "<Activity><broken new_status"));
            ScanOutcome o = new BusinessRuleScanner().Scan(Ctx(svc, SearchMode.Column), CancellationToken.None);
            Assert.AreEqual(1, o.Diagnostic.MalformedDefinitions);
            Assert.AreEqual(ScanStatus.CompletedWithWarnings, o.Diagnostic.Status);
            Assert.AreEqual(1, o.Results.Count);
            Assert.AreEqual(ConfidenceLevel.TextMatchReviewRequired, o.Results[0].Confidence);
        }

        [TestMethod]
        public void BusinessRule_UnreadableDefinition_IsWarnedNotSilentlySkipped()
        {
            var svc = new FakeOrganizationService().On("workflow", F.Row("workflow", null, "name", "Empty", "primaryentity", "account"));
            ScanOutcome o = new BusinessRuleScanner().Scan(Ctx(svc, SearchMode.Column), CancellationToken.None);
            Assert.AreEqual(0, o.Results.Count);
            StringAssert.Contains(o.Diagnostic.Message, "no readable definition");
            Assert.AreEqual(ScanStatus.CompletedWithWarnings, o.Diagnostic.Status);
        }

        // ------------------------------------------------------------ classic workflows

        private const string Xaml = @"<Activity xmlns:mxswa='clr-namespace:Microsoft.Crm.Workflow.Activities;assembly=Microsoft.Crm.Workflow' xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'>
  <mxswa:ActivityReference DisplayName='Update Account'>
    <mxswa:ActivityReference.Properties>
      <InArgument x:Key='EntityName'>[""account""]</InArgument>
      <InArgument x:Key='Attribute'>[""new_status""]</InArgument>
    </mxswa:ActivityReference.Properties>
  </mxswa:ActivityReference>
  <mxswa:ActivityReference DisplayName='Update Contact'>
    <mxswa:ActivityReference.Properties>
      <InArgument x:Key='EntityName'>[""contact""]</InArgument>
      <InArgument x:Key='Attribute'>[""new_status""]</InArgument>
    </mxswa:ActivityReference.Properties>
  </mxswa:ActivityReference>
</Activity>";

        private static Entity Proc(string name, string xaml, string primary = "account", int category = 0)
        {
            return F.Row("workflow", null, "name", name, "primaryentity", primary, "category", new OptionSetValue(category),
                         "statecode", new OptionSetValue(1), "xaml", xaml);
        }

        [TestMethod]
        public void ClassicWorkflow_ColumnSearch_UsesStepEntityToIgnoreOtherTables()
        {
            var svc = new FakeOrganizationService().On("workflow", Proc("WF", Xaml));
            ScanOutcome o = new ClassicWorkflowScanner().Scan(Ctx(svc, SearchMode.Column), CancellationToken.None);
            Assert.AreEqual(1, o.Results.Count, "the contact step must be ignored");
            DependencyResult r = o.Results[0];
            Assert.AreEqual(ConfidenceLevel.StrongDefinitionReference, r.Confidence);
            StringAssert.Contains(r.Evidence, "Update Account");
            StringAssert.Contains(r.Location, "InArgument");
            StringAssert.Contains(string.Join(" ", o.Diagnostic.Limitations), "different table");
        }

        [TestMethod]
        public void ClassicWorkflow_DefaultScan_IsTargetedToPrimaryTable_AndLimitsAreStated()
        {
            var svc = new FakeOrganizationService().On("workflow");
            ScanOutcome o = new ClassicWorkflowScanner().Scan(Ctx(svc, SearchMode.Column), CancellationToken.None);
            Assert.IsTrue(HasCondition(svc.Queries[0], "primaryentity", "account"));
            Assert.IsFalse(HasCondition(svc.Queries[0], "category", 2), "business rules are not classic workflows");
            StringAssert.Contains(string.Join(" ", o.Diagnostic.Limitations), "deep");
        }

        [TestMethod]
        public void ClassicWorkflow_DeepScan_DoesNotFilterByPrimaryTable_AndMarksUnknownEntityAsPossible()
        {
            string noEntity = "<Activity><Step Attribute='new_status' DisplayName='Set'/></Activity>";
            var svc = new FakeOrganizationService().On("workflow", Proc("Other table WF", noEntity, primary: "contact"));
            ScanOutcome o = new ClassicWorkflowScanner().Scan(Ctx(svc, SearchMode.Column, tweak: r => r.DeepProcessScan = true), CancellationToken.None);
            Assert.IsFalse(svc.Queries[0].Criteria.Conditions.Any(c => c.AttributeName == "primaryentity"));
            Assert.AreEqual(1, o.Results.Count);
            Assert.AreEqual(ConfidenceLevel.PossibleOrDynamic, o.Results[0].Confidence);
        }

        [TestMethod]
        public void ClassicWorkflow_MalformedXaml_IsCountedAndTextMatched()
        {
            var svc = new FakeOrganizationService().On("workflow", Proc("Bad", "<Activity Attribute='new_status'"));
            ScanOutcome o = new ClassicWorkflowScanner().Scan(Ctx(svc, SearchMode.Column), CancellationToken.None);
            Assert.AreEqual(1, o.Diagnostic.MalformedDefinitions);
            Assert.AreEqual(ConfidenceLevel.TextMatchReviewRequired, o.Results.Single().Confidence);
        }

        [TestMethod]
        public void ClassicWorkflow_MissingXamlColumn_IsPartiallySupported()
        {
            var svc = new FakeOrganizationService();
            int call = 0;
            svc.Handlers["workflow"] = q =>
            {
                if (call++ == 0) throw F.Fault(-2147217149, "'workflow' entity doesn't contain attribute with Name = 'xaml'");
                return new[] { F.Row("workflow", null, "name", "P", "primaryentity", "account", "category", new OptionSetValue(0)) };
            };
            ScanOutcome o = new ClassicWorkflowScanner().Scan(Ctx(svc, SearchMode.Table), CancellationToken.None);
            Assert.AreEqual(ScanStatus.PartiallySupported, o.Diagnostic.Status);
            Assert.AreEqual(1, o.Results.Count);
        }

        // ------------------------------------------------------------ plugins

        private static Entity Step(string name, string filtering, string message = "Update", Guid? id = null)
        {
            var e = F.Row("sdkmessageprocessingstep", id, "name", name, "stage", new OptionSetValue(40), "mode", new OptionSetValue(0),
                          "statecode", new OptionSetValue(0), "rank", 1, "ismanaged", false);
            e["m.name"] = new AliasedValue("sdkmessage", "name", message);
            e["f.primaryobjecttypecode"] = new AliasedValue("sdkmessagefilter", "primaryobjecttypecode", "account");
            if (filtering != null) e["filteringattributes"] = filtering;
            return e;
        }

        private static Entity Image(Guid stepId, string name, string attributes, int type = 0)
        {
            var e = F.Row("sdkmessageprocessingstepimage", null, "name", name, "entityalias", name, "imagetype", new OptionSetValue(type),
                          "sdkmessageprocessingstepid", new EntityReference("sdkmessageprocessingstep", stepId));
            if (attributes != null) e["attributes"] = attributes;
            return e;
        }

        [TestMethod]
        public void Plugin_TableSearch_QueriesByFilterAndDoesNotRetrieveConfiguration()
        {
            var step = Step("S1", "name");
            var svc = new FakeOrganizationService().On("sdkmessageprocessingstep", step).On("sdkmessageprocessingstepimage").On("plugintype");
            ScanOutcome o = new PluginRegistrationScanner().Scan(Ctx(svc, SearchMode.Table), CancellationToken.None);
            Assert.AreEqual(1, o.Results.Count);
            Assert.AreEqual(ConfidenceLevel.ConfirmedMetadataReference, o.Results[0].Confidence);
            QueryExpression q = svc.Queries.First(x => x.EntityName == "sdkmessageprocessingstep");
            Assert.IsTrue(q.LinkEntities.Any(l => l.LinkToEntityName == "sdkmessagefilter" && l.LinkCriteria.Conditions.Any(c => c.AttributeName == "primaryobjecttypecode")));
            Assert.IsFalse(q.ColumnSet.Columns.Contains("configuration"), "configuration may hold secrets");
            Assert.IsFalse(q.ColumnSet.Columns.Contains("secureconfig"));
            StringAssert.Contains(string.Join(" ", o.Diagnostic.Limitations), "compiled plugin");
        }

        [TestMethod]
        public void Plugin_ColumnSearch_ClassifiesFilteringImageAndAllAttributes()
        {
            Guid a = Guid.NewGuid(), b = Guid.NewGuid(), c = Guid.NewGuid(), d = Guid.NewGuid();
            var svc = new FakeOrganizationService()
                .On("sdkmessageprocessingstep", Step("Filtered", "name,new_status", id: a), Step("ImageOnly", "name", id: b), Step("AllAttrs", null, id: c), Step("Unrelated", "name", id: d))
                .On("sdkmessageprocessingstepimage", Image(b, "Pre", "name,new_status"), Image(d, "Pre", "name"))
                .On("plugintype");
            ScanOutcome o = new PluginRegistrationScanner().Scan(Ctx(svc, SearchMode.Column), CancellationToken.None);

            Assert.IsTrue(o.Results.Any(r => r.ComponentName == "Filtered" && r.MatchType == "Filtering attribute" && r.Confidence == ConfidenceLevel.ConfirmedMetadataReference));
            Assert.IsTrue(o.Results.Any(r => r.ComponentName == "ImageOnly" && r.MatchType.StartsWith("Image attribute") && r.Confidence == ConfidenceLevel.ConfirmedMetadataReference));
            Assert.IsTrue(o.Results.Any(r => r.ComponentName == "AllAttrs" && r.Confidence == ConfidenceLevel.PossibleOrDynamic), "Update step without filter fires on any column");
            Assert.IsFalse(o.Results.Any(r => r.ComponentName == "Unrelated"));
        }

        [TestMethod]
        public void Plugin_ImageQueryDenied_IsPartiallySupportedWithPermissionFlag()
        {
            var svc = new FakeOrganizationService().On("sdkmessageprocessingstep", Step("S", "name"))
                .OnThrow("sdkmessageprocessingstepimage", F.PermissionFault()).On("plugintype");
            ScanOutcome o = new PluginRegistrationScanner().Scan(Ctx(svc, SearchMode.Column), CancellationToken.None);
            Assert.AreEqual(ScanStatus.PartiallySupported, o.Diagnostic.Status);
            Assert.IsTrue(o.Diagnostic.PermissionLimited);
        }

        [TestMethod]
        public void Plugin_NoSteps_IsCompleted_WithZeroMatches()
        {
            var svc = new FakeOrganizationService().On("sdkmessageprocessingstep");
            ScanOutcome o = new PluginRegistrationScanner().Scan(Ctx(svc, SearchMode.Table), CancellationToken.None);
            Assert.AreEqual(ScanStatus.Completed, o.Diagnostic.Status);
            Assert.AreEqual(0, o.Diagnostic.RecordsExamined);
        }

        // ------------------------------------------------------------ forms & javascript

        private static Entity FormRow(string name, string xml = null)
        {
            return F.Row("systemform", null, "name", name, "objecttypecode", "account", "type", new OptionSetValue(2),
                         "formxml", xml ?? ParserTests.FormXml, "ismanaged", false, "formactivationstate", new OptionSetValue(1));
        }

        private static Entity Script(string name, string js)
        {
            return F.Row("webresource", null, "name", name, "content", Convert.ToBase64String(new UTF8Encoding(false).GetBytes(js)), "ismanaged", false);
        }

        private const string AccountJs = "var NS = NS || {};\nNS.Account = {\n  onLoad: function (ctx) {\n    var s = ctx.getFormContext().getAttribute('new_status');\n  },\n  onSave: function () { return 1; }\n};";

        [TestMethod]
        public void Form_TableSearch_ReturnsFormLibrariesAndHandlers()
        {
            var svc = new FakeOrganizationService().On("systemform", FormRow("Main"));
            ScanOutcome o = new FormScanner().Scan(Ctx(svc, SearchMode.Table), CancellationToken.None);
            Assert.AreEqual(1, o.Results.Count(r => r.ComponentType == ComponentType.Form && r.Confidence == ConfidenceLevel.ConfirmedMetadataReference));
            Assert.AreEqual(2, o.Results.Count(r => r.MatchType == "JavaScript library registered on form"));
            Assert.AreEqual(3, o.Results.Count(r => r.MatchType.EndsWith("handler")));
            Assert.IsTrue(o.Results.Any(r => r.ComponentState != null && r.ComponentState.Contains("Handler disabled")));
            Assert.IsTrue(HasCondition(svc.Queries[0], "objecttypecode", "account"));
        }

        [TestMethod]
        public void Form_ColumnSearch_DistinguishesControlEventAndHandlerSourceReference()
        {
            var svc = new FakeOrganizationService().On("systemform", FormRow("Main"))
                .OnWebResources(Script("new_/js/account.js", AccountJs), Script("new_/js/other.js", "var x = 1;"));
            ScanOutcome o = new FormScanner().Scan(Ctx(svc, SearchMode.Column), CancellationToken.None);

            DependencyResult control = o.Results.Single(r => r.MatchType == "Field control on form");
            Assert.AreEqual(ConfidenceLevel.StrongDefinitionReference, control.Confidence);
            DependencyResult change = o.Results.Single(r => r.MatchType.StartsWith("Event registered against the column"));
            Assert.AreEqual(ConfidenceLevel.StrongDefinitionReference, change.Confidence);
            DependencyResult source = o.Results.Single(r => r.MatchType == "Handler function source references the column");
            Assert.AreEqual(ConfidenceLevel.TextMatchReviewRequired, source.Confidence);
            Assert.AreEqual("NS.Account.onLoad", source.ComponentUniqueName);
            Assert.AreEqual(4, source.LineNumber);
            Assert.IsFalse(o.Results.Any(r => r.ComponentUniqueName == "NS.Account.onSave"), "onSave does not use the column and is disabled");
        }

        [TestMethod]
        public void Form_MalformedFormXml_IsReportedNotTreatedAsEmpty()
        {
            var svc = new FakeOrganizationService().On("systemform", FormRow("Broken", "<form><tabs>"));
            ScanOutcome o = new FormScanner().Scan(Ctx(svc, SearchMode.Table), CancellationToken.None);
            Assert.AreEqual(0, o.Results.Count);
            Assert.AreEqual(1, o.Diagnostic.MalformedDefinitions);
            Assert.AreEqual(ScanStatus.CompletedWithWarnings, o.Diagnostic.Status);
        }

        [TestMethod]
        public void JavaScript_ColumnSearch_ReportsTextMatchesOnlyAndExplainsFormRegistration()
        {
            var svc = new FakeOrganizationService().On("systemform", FormRow("Main"))
                .OnWebResources(Script("new_/js/account.js", AccountJs), Script("new_/js/unrelated.js", "var a = f.getAttribute('new_status'); // new_status"));
            ScanOutcome o = new JavaScriptScanner().Scan(Ctx(svc, SearchMode.Column), CancellationToken.None);

            Assert.AreEqual(2, o.Results.Count);
            Assert.IsTrue(o.Results.All(r => r.Confidence == ConfidenceLevel.TextMatchReviewRequired));
            DependencyResult reg = o.Results.Single(r => r.ComponentName == "new_/js/account.js");
            StringAssert.Contains(reg.Explanation, "Registered on 1 form");
            Assert.AreEqual(4, reg.LineNumber);
            DependencyResult unreg = o.Results.Single(r => r.ComponentName == "new_/js/unrelated.js");
            StringAssert.Contains(unreg.Explanation, "Not registered on any form");
            StringAssert.Contains(string.Join(" ", o.Diagnostic.Limitations), "comments");
            Assert.AreEqual(2, o.Diagnostic.RecordsExamined);
        }

        [TestMethod]
        public void JavaScript_RegisteredScope_OnlyDownloadsRegisteredLibraries()
        {
            var svc = new FakeOrganizationService().On("systemform", FormRow("Main"))
                .OnWebResources(Script("new_/js/account.js", AccountJs));
            ScanContext ctx = Ctx(svc, SearchMode.Column, tweak: r => r.JavaScriptScope = JavaScriptScope.RegisteredOnSelectedTableForms);
            ScanOutcome o = new JavaScriptScanner().Scan(ctx, CancellationToken.None);
            QueryExpression wq = svc.Queries.First(x => x.EntityName == "webresource");  // targeted: by name, not a full listing
            Assert.IsTrue(wq.Criteria.Conditions.Any(c => c.AttributeName == "name" && c.Operator == ConditionOperator.In), "targeted by name");
            StringAssert.Contains(o.Diagnostic.Message, "other.js");   // registered but not found as a web resource
            Assert.AreEqual(ScanStatus.CompletedWithWarnings, o.Diagnostic.Status);
        }

        [TestMethod]
        public void JavaScript_DynamicAttributeAccessOnRegisteredLibrary_IsPossibleNotConfirmed()
        {
            string js = "function onLoad(c){ var f = c.getFormContext(); f.getAttribute('new_' + suffix); }";
            var svc = new FakeOrganizationService().On("systemform", FormRow("Main")).OnWebResources(Script("new_/js/account.js", js));
            ScanOutcome o = new JavaScriptScanner().Scan(Ctx(svc, SearchMode.Column), CancellationToken.None);
            DependencyResult d = o.Results.Single(r => r.MatchType == "Dynamic attribute access");
            Assert.AreEqual(ConfidenceLevel.PossibleOrDynamic, d.Confidence);
        }

        [TestMethod]
        public void JavaScript_EmptyAndUndecodableContent_IsWarned()
        {
            var svc = new FakeOrganizationService().On("systemform").OnWebResources(
                F.Row("webresource", null, "name", "new_/js/empty.js"), F.Row("webresource", null, "name", "new_/js/bad.js", "content", "!!not-base64!!"));
            ScanOutcome o = new JavaScriptScanner().Scan(Ctx(svc, SearchMode.Column), CancellationToken.None);
            StringAssert.Contains(o.Diagnostic.Message, "empty or undecodable");
        }

        [TestMethod]
        public void WebResource_DecodingHonoursByteOrderMarks()
        {
            string enc;
            Assert.AreEqual("abc", WebResourceRepository.DecodeText(new byte[] { 0xEF, 0xBB, 0xBF, 0x61, 0x62, 0x63 }, out enc));
            Assert.AreEqual("abc", WebResourceRepository.DecodeText(Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes("abc")).ToArray(), out enc));
            Assert.AreEqual("UTF-16 LE", enc);
            Assert.AreEqual("caf\u00e9", WebResourceRepository.DecodeText(new byte[] { 0x63, 0x61, 0x66, 0xE9 }, out enc));
            StringAssert.Contains(enc, "1252");
            Assert.AreEqual("", WebResourceRepository.DecodeText(new byte[0], out enc));
        }

        // ------------------------------------------------------------ cloud flows

        [TestMethod]
        public void CloudFlow_IsAlwaysPartiallySupported_AndExplainsWhy()
        {
            var svc = new FakeOrganizationService().On("workflow");
            ScanOutcome o = new CloudFlowScanner().Scan(Ctx(svc, SearchMode.Column), CancellationToken.None);
            Assert.AreEqual(ScanStatus.PartiallySupported, o.Diagnostic.Status);
            StringAssert.Contains(string.Join(" ", o.Diagnostic.Limitations), "Management API");
            StringAssert.Contains(o.Diagnostic.Message, "does NOT mean");
            Assert.IsTrue(HasCondition(svc.Queries[0], "category", 5));
        }

        [TestMethod]
        public void CloudFlow_FindsStructuredReferences_AndReportsBadJson()
        {
            string json = "{\"properties\":{\"definition\":{\"actions\":{\"Update\":{\"inputs\":{\"host\":{\"apiId\":\"shared_commondataserviceforapps\"},\"parameters\":{\"entityName\":\"accounts\",\"item/new_status\":1}}}}}}}";
            var svc = new FakeOrganizationService().On("workflow",
                F.Row("workflow", null, "name", "Flow A", "clientdata", json, "statecode", new OptionSetValue(1)),
                F.Row("workflow", null, "name", "Flow B", "clientdata", "{ nope"));
            ScanOutcome o = new CloudFlowScanner().Scan(Ctx(svc, SearchMode.Column), CancellationToken.None);
            Assert.AreEqual(1, o.Results.Count);
            Assert.AreEqual(ConfidenceLevel.StrongDefinitionReference, o.Results[0].Confidence);
            StringAssert.Contains(o.Results[0].Location, "item/new_status");
            Assert.AreEqual(1, o.Diagnostic.MalformedDefinitions);
            Assert.AreEqual(ScanStatus.PartiallySupported, o.Diagnostic.Status);
        }

        [TestMethod]
        public void CloudFlow_PermissionDenied_IsNotSilentlySkipped()
        {
            var svc = new FakeOrganizationService().OnThrow("workflow", F.PermissionFault());
            ScanOutcome o = new CloudFlowScanner().Scan(Ctx(svc, SearchMode.Table), CancellationToken.None);
            Assert.AreEqual(ScanStatus.PermissionDenied, o.Diagnostic.Status);
        }
    }

    internal static class ResultExtensions
    {
        public static int RecordsLine(this DependencyResult r) { return r.LineNumber ?? -1; }
    }
}
