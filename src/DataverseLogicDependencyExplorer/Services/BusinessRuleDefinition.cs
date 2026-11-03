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

namespace DataverseLogicDependencyExplorer.Services
{
    public class BusinessRuleUsage
    {
        public int Line { get; set; }
        /// <summary>Condition, Action or Reference.</summary>
        public string Role { get; set; }
        public string Text { get; set; }
    }

    /// <summary>
    /// Reads the generated client code stored in a business rule's clientdata (the same representation the Business Rule
    /// Exporter interprets) and reports how a column is used.
    /// </summary>
    public static class BusinessRuleDefinition
    {
        private static readonly Regex ClientCode = new Regex(
            @"<clientcode>\s*<!\[CDATA\[(.*?)\]\]>\s*</clientcode>",
            RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);

        private static readonly Regex AttributeGet = new Regex(
            @"\b(v\d+)\s*=(?!=)\s*[^;\n]*?\.attributes\.get\(\s*['""]([^'""]+)['""]\s*\)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex AnyAttributeRef = new Regex(
            @"(?:attributes|controls)\.get\(\s*['""]([^'""]+)['""]\s*\)|getAttribute\(\s*['""]([^'""]+)['""]\s*\)|getControl\(\s*['""]([^'""]+)['""]\s*\)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ConditionLine = new Regex(@"^\s*(?:\}\s*)?(?:else\s+)?if\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ActionLine = new Regex(
            @"\.(setValue|setDisabled|setVisible|setRequiredLevel|setSubmitMode|setNotification|clearNotification|setFieldNotification|clearFieldNotification|addNotification|setLabel)\s*\(",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>Returns the generated JavaScript of the rule, or an empty string when none can be read.</summary>
        public static string ExtractClientCode(string clientData)
        {
            if (string.IsNullOrWhiteSpace(clientData)) return string.Empty;

            string decoded = System.Net.WebUtility.HtmlDecode(clientData);
            Match match = ClientCode.Match(decoded);
            if (match.Success) return match.Groups[1].Value.Trim();

            try
            {
                XDocument doc = XDocument.Parse(decoded);
                XElement node = doc.Descendants().FirstOrDefault(e => e.Name.LocalName.Equals("clientcode", StringComparison.OrdinalIgnoreCase));
                return node == null ? string.Empty : string.Concat(node.Nodes()).Trim();
            }
            catch (System.Xml.XmlException)
            {
                return string.Empty;
            }
        }

        /// <summary>Variable (v1, v2...) to column logical name map built from the generated declarations.</summary>
        public static Dictionary<string, string> MapVariables(string js)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrEmpty(js)) return map;

            foreach (Match m in AttributeGet.Matches(js)) map[m.Groups[1].Value] = m.Groups[2].Value;

            var derived = new[]
            {
                @"\b(v\d+)\s*=\s*(v\d+)\s*\.\s*get(?:Utc)?Value\s*\(\s*\)",
                @"\b(v\d+)\s*=\s*\(\s*(v\d+)\s*\)\s*\?\s*\2\s*\.\s*get(?:Utc)?Value\s*\(\s*\)"
            };
            foreach (string pattern in derived)
            {
                foreach (Match m in Regex.Matches(js, pattern, RegexOptions.IgnoreCase))
                {
                    string created = m.Groups[1].Value, source = m.Groups[2].Value;
                    if (map.ContainsKey(source) && !map.ContainsKey(created)) map[created] = map[source];
                }
            }
            return map;
        }

        /// <summary>Every distinct column the generated code touches.</summary>
        public static List<string> ReferencedColumns(string js)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string v in MapVariables(js).Values) names.Add(v);
            foreach (Match m in AnyAttributeRef.Matches(js ?? string.Empty))
                for (int g = 1; g <= 3; g++) if (m.Groups[g].Success) names.Add(m.Groups[g].Value);
            return names.OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList();
        }

        /// <summary>
        /// Lines of the generated code that use <paramref name="column"/>: conditions, actions, or (when nothing else uses it) the
        /// declaration. Line numbers refer to the extracted client code, not the raw clientdata XML.
        /// </summary>
        public static List<BusinessRuleUsage> FindUsage(string js, string column)
        {
            var usages = new List<BusinessRuleUsage>();
            if (string.IsNullOrEmpty(js) || string.IsNullOrEmpty(column)) return usages;

            Dictionary<string, string> map = MapVariables(js);
            var vars = map.Where(kv => string.Equals(kv.Value, column, StringComparison.OrdinalIgnoreCase)).Select(kv => kv.Key).ToList();
            var declarations = new List<BusinessRuleUsage>();

            string[] lines = js.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                bool direct = AnyAttributeRef.Matches(line).Cast<Match>().Any(m =>
                    Enumerable.Range(1, 3).Any(g => m.Groups[g].Success && string.Equals(m.Groups[g].Value, column, StringComparison.OrdinalIgnoreCase)));
                bool viaVariable = vars.Any(v => TextLocator.ContainsWord(line, v));
                if (!direct && !viaVariable) continue;

                bool isDeclaration = AttributeGet.IsMatch(line) && !ConditionLine.IsMatch(line) && !ActionLine.IsMatch(line);
                string role = ConditionLine.IsMatch(line) ? "Condition" : ActionLine.IsMatch(line) ? "Action" : "Reference";
                var usage = new BusinessRuleUsage { Line = i + 1, Role = role, Text = line.Trim() };
                if (isDeclaration) declarations.Add(usage); else usages.Add(usage);
            }

            if (usages.Count == 0) usages.AddRange(declarations);
            return usages;
        }

        public static string DescribeRoles(IEnumerable<BusinessRuleUsage> usages)
        {
            bool cond = usages.Any(u => u.Role == "Condition");
            bool act = usages.Any(u => u.Role == "Action");
            if (cond && act) return "Condition and action reference";
            if (cond) return "Condition reference";
            if (act) return "Action reference";
            return "Field reference";
        }
    }
}
