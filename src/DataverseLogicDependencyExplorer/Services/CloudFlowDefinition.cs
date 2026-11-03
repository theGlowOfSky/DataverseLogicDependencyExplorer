// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Newtonsoft.Json.Linq;

namespace DataverseLogicDependencyExplorer.Services
{
    public class FlowReference
    {
        public string ActionName { get; set; }
        public bool IsTrigger { get; set; }
        public string Path { get; set; }
        public string Expression { get; set; }
        /// <summary>Structured = parsed parameter of a Dataverse action; Expression = dynamic-content token in an expression.</summary>
        public bool Structured { get; set; }
        public string Role { get; set; }
        /// <summary>True when the owning Dataverse step's table is not statically known (expression).</summary>
        public bool TableUnresolved { get; set; }
        public string StepTable { get; set; }
    }

    public class FlowAnalysis
    {
        public List<FlowReference> TableReferences { get; } = new List<FlowReference>();
        public List<FlowReference> ColumnReferences { get; } = new List<FlowReference>();
        public int DataverseSteps { get; set; }
        public int UnresolvedTableSteps { get; set; }
        public List<string> ConnectionReferences { get; } = new List<string>();
        public string Error { get; set; }
    }

    /// <summary>Parses the cloud flow JSON stored in workflow.clientdata. Matches are tied to a step and a JSON path.</summary>
    public static class CloudFlowDefinition
    {
        public static FlowAnalysis Analyze(string clientData, string table, string entitySet, string column)
        {
            var result = new FlowAnalysis();
            JToken root;
            try { root = JToken.Parse(clientData); }
            catch (Exception ex) { result.Error = "The flow definition is not valid JSON: " + ex.Message; return result; }

            var props = root.SelectToken("properties") ?? root;
            var connRefs = props.SelectToken("connectionReferences") as JObject;
            if (connRefs != null)
                foreach (JProperty p in connRefs.Properties())
                {
                    string logical = (string)p.Value.SelectToken("connection.connectionReferenceLogicalName");
                    result.ConnectionReferences.Add(logical == null ? p.Name : p.Name + " (" + logical + ")");
                }

            var def = props.SelectToken("definition") as JObject;
            if (def == null) { result.Error = "The flow JSON has no properties.definition section."; return result; }

            var structuredPaths = new HashSet<string>();
            var triggers = def["triggers"] as JObject;
            if (triggers != null)
                foreach (JProperty t in triggers.Properties())
                    if (t.Value is JObject) AnalyzeStep((JObject)t.Value, t.Name, true, table, entitySet, column, result, structuredPaths);

            WalkActions(def["actions"] as JObject, table, entitySet, column, result, structuredPaths);

            if (!string.IsNullOrEmpty(column)) ExpressionPass(def, column, result, structuredPaths);
            return result;
        }

        private static void WalkActions(JObject actions, string table, string entitySet, string column, FlowAnalysis result, HashSet<string> structuredPaths)
        {
            if (actions == null) return;
            foreach (JProperty p in actions.Properties())
            {
                var action = p.Value as JObject;
                if (action == null) continue;
                AnalyzeStep(action, p.Name, false, table, entitySet, column, result, structuredPaths);

                WalkActions(action["actions"] as JObject, table, entitySet, column, result, structuredPaths);
                var elseBlock = action["else"] as JObject;
                if (elseBlock != null) WalkActions(elseBlock["actions"] as JObject, table, entitySet, column, result, structuredPaths);
                var cases = action["cases"] as JObject;
                if (cases != null)
                    foreach (JProperty c in cases.Properties())
                        if (c.Value is JObject) WalkActions(((JObject)c.Value)["actions"] as JObject, table, entitySet, column, result, structuredPaths);
                var dflt = action["default"] as JObject;
                if (dflt != null) WalkActions(dflt["actions"] as JObject, table, entitySet, column, result, structuredPaths);
            }
        }

        private static bool IsDataverse(JObject step)
        {
            string api = (string)step.SelectToken("inputs.host.apiId") ?? (string)step.SelectToken("inputs.host.connectionName") ?? string.Empty;
            return api.IndexOf("commondataservice", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void AnalyzeStep(JObject step, string name, bool isTrigger, string table, string entitySet, string column, FlowAnalysis result, HashSet<string> structuredPaths)
        {
            if (!IsDataverse(step)) return;
            result.DataverseSteps++;

            var parameters = step.SelectToken("inputs.parameters") as JObject;
            if (parameters == null) return;

            string entityValue = (string)parameters["subscriptionRequest/entityname"] ?? (string)parameters["entityName"];
            bool isExpression = entityValue != null && entityValue.TrimStart().StartsWith("@");
            bool tableMatches = entityValue != null && !isExpression &&
                (string.Equals(entityValue, table, StringComparison.OrdinalIgnoreCase) ||
                 (!string.IsNullOrEmpty(entitySet) && string.Equals(entityValue, entitySet, StringComparison.OrdinalIgnoreCase)));
            bool unresolved = entityValue == null || isExpression;
            if (unresolved) result.UnresolvedTableSteps++;

            if (string.IsNullOrEmpty(column))
            {
                if (tableMatches)
                {
                    JProperty prop = parameters.Property(parameters["subscriptionRequest/entityname"] != null ? "subscriptionRequest/entityname" : "entityName");
                    result.TableReferences.Add(new FlowReference
                    {
                        ActionName = name, IsTrigger = isTrigger, Structured = true, Path = prop.Path,
                        Expression = prop.Name + " = \"" + entityValue + "\"", Role = isTrigger ? "Trigger table" : "Action table", StepTable = entityValue
                    });
                }
                return;
            }

            if (!tableMatches && !unresolved) return; // the step works on a different table

            foreach (JProperty p in parameters.Properties())
            {
                string key = p.Name;
                string value = p.Value.Type == JTokenType.String ? (string)p.Value : null;
                string role = null;

                if (key.StartsWith("item/", StringComparison.OrdinalIgnoreCase))
                {
                    string field = key.Substring(5);
                    int at = field.IndexOf('@');
                    if (at >= 0) field = field.Substring(0, at);
                    if (string.Equals(field, column, StringComparison.OrdinalIgnoreCase)) role = "Column set by the step (" + key + ")";
                }
                else if (value != null)
                {
                    string k = key.ToLowerInvariant();
                    if ((k == "$select" || k == "subscriptionrequest/filteringattributes" || k == "$orderby") &&
                        TextLocator.SplitList(value.Replace(" asc", string.Empty).Replace(" desc", string.Empty)).Any(x => string.Equals(x, column, StringComparison.OrdinalIgnoreCase)))
                        role = key + " lists the column";
                    else if (k == "$filter" || k == "$apply" || k == "$expand")
                    {
                        if (TextLocator.ContainsWord(value, column) || TextLocator.ContainsWord(value, "_" + column + "_value")) role = key + " expression uses the column";
                    }
                    else if (k == "fetchxml" && FetchXmlUses(value, column)) role = "FetchXML uses the column";
                }

                if (role == null) continue;
                structuredPaths.Add(p.Path);
                result.ColumnReferences.Add(new FlowReference
                {
                    ActionName = name, IsTrigger = isTrigger, Structured = true, Path = p.Path,
                    Expression = DataverseHelper.Truncate(key + ": " + (value ?? p.Value.ToString(Newtonsoft.Json.Formatting.None)), 500),
                    Role = role, TableUnresolved = unresolved && !tableMatches, StepTable = entityValue
                });
            }
        }

        private static bool FetchXmlUses(string fetch, string column)
        {
            try
            {
                XDocument doc = XDocument.Parse(fetch);
                return doc.Descendants().Any(e => e.Attributes().Any(a =>
                    (a.Name.LocalName == "name" || a.Name.LocalName == "attribute") && string.Equals(a.Value, column, StringComparison.OrdinalIgnoreCase)));
            }
            catch (System.Xml.XmlException) { return TextLocator.ContainsWord(fetch, column); }
        }

        /// <summary>Dynamic-content tokens such as @{triggerOutputs()?['body/col']}: they prove a textual use, not the owning table.</summary>
        private static void ExpressionPass(JObject def, string column, FlowAnalysis result, HashSet<string> structuredPaths)
        {
            var rx = new Regex(@"\[\s*'(?:[\w./]*/)?" + Regex.Escape(column) + @"(?:@[\w.]+)?'\s*\]|\b_" + Regex.Escape(column) + @"_value\b", RegexOptions.IgnoreCase);
            foreach (JToken t in def.DescendantsAndSelf())
            {
                if (t.Type != JTokenType.String) continue;
                string s = (string)t;
                if (s.IndexOf('@') < 0 || !rx.IsMatch(s)) continue;
                string path = t.Path;
                if (structuredPaths.Contains(path)) continue;
                result.ColumnReferences.Add(new FlowReference
                {
                    ActionName = OwningStep(t), Structured = false, Path = path,
                    Expression = DataverseHelper.Truncate(s, 500), Role = "Dynamic-content expression"
                });
            }
        }

        private static string OwningStep(JToken t)
        {
            for (JToken p = t; p != null; p = p.Parent)
            {
                JProperty prop = p as JProperty;
                if (prop != null && prop.Parent != null && prop.Parent.Parent is JProperty)
                {
                    string container = ((JProperty)prop.Parent.Parent).Name;
                    if (container == "actions" || container == "triggers") return prop.Name;
                }
            }
            return null;
        }
    }
}
