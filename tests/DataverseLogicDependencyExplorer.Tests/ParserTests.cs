// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Linq;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DataverseLogicDependencyExplorer.Services;

namespace DataverseLogicDependencyExplorer.Tests
{
    [TestClass]
    public class ParserTests
    {
        public const string FormXml = @"<form>
  <tabs><tab id='t1'><columns><column><sections><section id='s1'><rows>
    <row><cell><control id='name' datafieldname='name' /></cell></row>
    <row><cell><control id='new_status' datafieldname='new_status' /></cell></row>
    <row><cell><control id='contacts' classid='{x}'><parameters><TargetEntityType>contact</TargetEntityType></parameters></control></cell></row>
  </rows></section></sections></column></columns></tab></tabs>
  <formLibraries><Library name='new_/js/account.js' libraryUniqueId='1' /><Library name='new_/js/other.js' libraryUniqueId='2' /></formLibraries>
  <events>
    <event name='onload' application='false' active='false'><Handlers><Handler functionName='NS.Account.onLoad' libraryName='new_/js/account.js' handlerUniqueId='a' enabled='true' parameters='' passExecutionContext='true' /></Handlers></event>
    <event name='onsave' application='false' active='false'><Handlers><Handler functionName='NS.Account.onSave' libraryName='new_/js/account.js' handlerUniqueId='b' enabled='false' /></Handlers></event>
    <event name='onchange' application='false' active='true' attribute='new_status'><Handlers><Handler functionName='NS.Account.statusChanged' libraryName='new_/js/other.js' handlerUniqueId='c' enabled='true' /></Handlers></event>
  </events>
</form>";

        [TestMethod]
        public void FormXml_LibrariesHandlersAndControlsAreParsed()
        {
            FormDefinition d = FormXmlReader.Parse(FormXml);
            Assert.IsTrue(d.IsValid);
            Assert.AreEqual(2, d.Libraries.Count);
            Assert.AreEqual(3, d.Handlers.Count);
            Assert.AreEqual(2, d.Controls.Count);
            FormHandler change = d.Handlers.Single(h => h.EventName == "onchange");
            Assert.AreEqual("new_status", change.Attribute);
            Assert.AreEqual("NS.Account.statusChanged", change.FunctionName);
            Assert.IsTrue(d.Handlers.Single(h => h.EventName == "onload").PassExecutionContext);
            Assert.IsFalse(d.Handlers.Single(h => h.EventName == "onsave").Enabled);
        }

        [TestMethod]
        public void FormXml_Malformed_ReportsErrorInsteadOfEmptyForm()
        {
            FormDefinition d = FormXmlReader.Parse("<form><tabs>");
            Assert.IsFalse(d.IsValid);
            StringAssert.Contains(d.Error, "malformed");
            Assert.IsFalse(FormXmlReader.Parse("  ").IsValid);
            Assert.IsFalse(FormXmlReader.Parse(null).IsValid);
        }

        [TestMethod]
        public void FormXml_OtherReferences_ExcludeControlsAndEventRegistrations()
        {
            FormDefinition d = FormXmlReader.Parse(FormXml);
            Assert.AreEqual(0, FormXmlReader.FindOtherReferences(d, "new_status").Count);
            var refs = FormXmlReader.FindOtherReferences(d, "contact");
            Assert.AreEqual(1, refs.Count);
            StringAssert.Contains(refs[0].Path, "TargetEntityType");
        }

        [TestMethod]
        public void XmlPath_NormalizeStripsExpressionDecoration()
        {
            Assert.AreEqual("name", XmlPath.Normalize("[\"name\"]"));
            Assert.AreEqual("name", XmlPath.Normalize(" 'name' "));
            Assert.AreEqual("name", XmlPath.Normalize("name"));
        }

        [TestMethod]
        public void XmlPath_FindsExactValuesAndExpressionTokens_NotSubstrings()
        {
            XDocument doc = XDocument.Parse("<r><a Attribute='name'/><b>[\"name\"]</b><c Entity='[InputEntities(\"primaryEntity\").Get(\"name\")]'/><d Attribute='fullname'/><e>The name of it</e></r>");
            var m = XmlPath.FindExactValues(doc, "name");
            Assert.AreEqual(3, m.Count);
            Assert.IsTrue(m.Any(x => x.Kind == "expression token"));
        }

        [TestMethod]
        public void BusinessRule_ClientCodeIsExtractedFromCdata()
        {
            string clientData = "<clientdata><clientcode><![CDATA[var v1 = formContext.data.entity.attributes.get('new_status');\nif (v1.getValue() == 1) {\n v2.setRequiredLevel('required');\n}]]></clientcode></clientdata>";
            string js = BusinessRuleDefinition.ExtractClientCode(clientData);
            StringAssert.Contains(js, "attributes.get('new_status')");
            Assert.AreEqual("", BusinessRuleDefinition.ExtractClientCode("not xml"));
            Assert.AreEqual("", BusinessRuleDefinition.ExtractClientCode(null));
        }

        [TestMethod]
        public void BusinessRule_UsageDistinguishesConditionAndAction()
        {
            string js = "var v1 = formContext.data.entity.attributes.get('new_status');\nvar v2 = formContext.data.entity.attributes.get('new_other');\nif (v1.getValue() == 1) {\n  v2.setValue(5);\n}\nv1.setRequiredLevel('required');";
            var usage = BusinessRuleDefinition.FindUsage(js, "new_status");
            Assert.AreEqual("Condition and action reference", BusinessRuleDefinition.DescribeRoles(usage));
            Assert.AreEqual(3, usage.First().Line);
            var other = BusinessRuleDefinition.FindUsage(js, "new_other");
            Assert.AreEqual("Action reference", BusinessRuleDefinition.DescribeRoles(other));
            Assert.AreEqual(0, BusinessRuleDefinition.FindUsage(js, "missing").Count);
        }

        [TestMethod]
        public void BusinessRule_ReferencedColumnsAreListed()
        {
            string js = "var v1 = f.data.entity.attributes.get('a');\nvar v2 = f.data.entity.attributes.get('b');\nf.getControl('c').setVisible(true);";
            CollectionEquals(new[] { "a", "b", "c" }, BusinessRuleDefinition.ReferencedColumns(js));
        }

        private static void CollectionEquals(string[] expected, System.Collections.Generic.List<string> actual)
        {
            Assert.AreEqual(string.Join(",", expected), string.Join(",", actual));
        }

        [TestMethod]
        public void CloudFlow_ParsesStructuredColumnAndTableReferences()
        {
            string json = @"{ ""properties"": {
  ""connectionReferences"": { ""shared_commondataserviceforapps"": { ""connection"": { ""connectionReferenceLogicalName"": ""new_sharedcds"" } } },
  ""definition"": {
    ""triggers"": { ""When_a_row_is_added"": { ""type"": ""OpenApiConnectionWebhook"", ""inputs"": { ""host"": { ""apiId"": ""/providers/Microsoft.PowerApps/apis/shared_commondataserviceforapps"", ""operationId"": ""SubscribeWebhookTrigger"" },
        ""parameters"": { ""subscriptionRequest/entityname"": ""account"", ""subscriptionRequest/filteringattributes"": ""name,new_status"" } } } },
    ""actions"": {
      ""Update_row"": { ""type"": ""OpenApiConnection"", ""inputs"": { ""host"": { ""apiId"": ""shared_commondataserviceforapps"", ""operationId"": ""UpdateRecord"" }, ""parameters"": { ""entityName"": ""accounts"", ""item/new_status"": 2, ""$select"": ""name"" } } },
      ""Other_table"": { ""type"": ""OpenApiConnection"", ""inputs"": { ""host"": { ""apiId"": ""shared_commondataserviceforapps"" }, ""parameters"": { ""entityName"": ""contacts"", ""item/new_status"": 1 } } },
      ""Dynamic_table"": { ""type"": ""OpenApiConnection"", ""inputs"": { ""host"": { ""apiId"": ""shared_commondataserviceforapps"" }, ""parameters"": { ""entityName"": ""@variables('t')"", ""$select"": ""new_status"" } } },
      ""Compose"": { ""type"": ""Compose"", ""inputs"": ""@{triggerOutputs()?['body/new_status']}"" }
    } } } }";
            FlowAnalysis col = CloudFlowDefinition.Analyze(json, "account", "accounts", "new_status");
            Assert.IsNull(col.Error);
            Assert.AreEqual(1, col.ConnectionReferences.Count);
            Assert.IsTrue(col.ColumnReferences.Any(r => r.IsTrigger && r.Structured));
            Assert.IsTrue(col.ColumnReferences.Any(r => r.ActionName == "Update_row" && r.Structured && !r.TableUnresolved));
            Assert.IsFalse(col.ColumnReferences.Any(r => r.ActionName == "Other_table"), "contacts action must be ignored");
            Assert.IsTrue(col.ColumnReferences.Any(r => r.ActionName == "Dynamic_table" && r.TableUnresolved));
            Assert.IsTrue(col.ColumnReferences.Any(r => r.ActionName == "Compose" && !r.Structured));

            FlowAnalysis tbl = CloudFlowDefinition.Analyze(json, "account", "accounts", null);
            Assert.AreEqual(2, tbl.TableReferences.Count);
            Assert.AreEqual(1, tbl.UnresolvedTableSteps);
        }

        [TestMethod]
        public void CloudFlow_InvalidJson_IsReportedAsError()
        {
            Assert.IsNotNull(CloudFlowDefinition.Analyze("{ not json", "account", "accounts", "name").Error);
            Assert.IsNotNull(CloudFlowDefinition.Analyze("{\"properties\":{}}", "account", "accounts", "name").Error);
        }
    }
}
