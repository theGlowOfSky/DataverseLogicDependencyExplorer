// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace DataverseLogicDependencyExplorer.Services
{
    /// <summary>One value found inside an XML/XAML document.</summary>
    public class XmlValueMatch
    {
        public XElement Element { get; set; }
        /// <summary>Null when the value is the element's text.</summary>
        public XAttribute Attribute { get; set; }
        /// <summary>"attribute", "element text" or "expression token".</summary>
        public string Kind { get; set; }
        public string RawValue { get; set; }
        public string Path { get; set; }
    }

    /// <summary>Builds readable XPath-like locations and finds exact values in parsed XML.</summary>
    public static class XmlPath
    {
        public static string PathOf(XElement element)
        {
            if (element == null) return string.Empty;
            var parts = new List<string>();
            for (XElement e = element; e != null; e = e.Parent)
            {
                string name = e.Name.LocalName;
                XAttribute id = e.Attribute("id") ?? e.Attribute("Id") ?? e.Attribute("name");
                if (id != null && !string.IsNullOrEmpty(id.Value) && id.Value.Length <= 80)
                {
                    parts.Add(name + "[@" + id.Name.LocalName + "='" + id.Value + "']");
                }
                else if (e.Parent != null)
                {
                    var same = e.Parent.Elements().Where(x => x.Name == e.Name).ToList();
                    parts.Add(same.Count > 1 ? name + "[" + (same.IndexOf(e) + 1) + "]" : name);
                }
                else parts.Add(name);
            }
            parts.Reverse();
            return "/" + string.Join("/", parts);
        }

        /// <summary>Strips XAML/expression decoration: [&quot;name&quot;] -> name, 'name' -> name.</summary>
        public static string Normalize(string value)
        {
            if (value == null) return string.Empty;
            string v = value.Trim();
            if (v.Length >= 2 && v[0] == '[' && v[v.Length - 1] == ']') v = v.Substring(1, v.Length - 2).Trim();
            if (v.Length >= 2 && (v[0] == '"' || v[0] == '\'') && v[v.Length - 1] == v[0]) v = v.Substring(1, v.Length - 2).Trim();
            return v;
        }

        /// <summary>Quoted tokens inside a longer expression, e.g. Foo("a").Bar("b") -> a, b.</summary>
        public static IEnumerable<string> QuotedTokens(string expression)
        {
            if (string.IsNullOrEmpty(expression)) yield break;
            var rx = new System.Text.RegularExpressions.Regex("\"([^\"]+)\"|'([^']+)'");
            foreach (System.Text.RegularExpressions.Match m in rx.Matches(expression))
                yield return m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value;
        }

        /// <summary>
        /// Finds attributes and leaf-element texts whose (normalised) value equals <paramref name="value"/>, or that contain it as
        /// a quoted token inside an expression. Matching is on parsed values, never on arbitrary substrings.
        /// </summary>
        public static List<XmlValueMatch> FindExactValues(XDocument doc, string value, Func<XElement, XAttribute, bool> exclude = null)
        {
            var found = new List<XmlValueMatch>();
            if (doc == null || doc.Root == null || string.IsNullOrEmpty(value)) return found;

            foreach (XElement el in doc.Root.DescendantsAndSelf())
            {
                foreach (XAttribute a in el.Attributes())
                {
                    if (a.IsNamespaceDeclaration) continue;
                    if (exclude != null && exclude(el, a)) continue;
                    Evaluate(found, el, a, a.Value, "attribute", value);
                }

                if (!el.HasElements)
                {
                    string text = el.Value;
                    if (!string.IsNullOrWhiteSpace(text)) Evaluate(found, el, null, text, "element text", value);
                }
            }
            return found;
        }

        private static void Evaluate(List<XmlValueMatch> found, XElement el, XAttribute a, string raw, string kind, string value)
        {
            if (string.IsNullOrEmpty(raw)) return;
            if (string.Equals(Normalize(raw), value, StringComparison.OrdinalIgnoreCase))
            {
                found.Add(new XmlValueMatch { Element = el, Attribute = a, Kind = kind, RawValue = raw, Path = PathOf(el) + (a != null ? "/@" + a.Name.LocalName : "/text()") });
                return;
            }
            if (raw.Length > value.Length && raw.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                foreach (string token in QuotedTokens(raw))
                {
                    if (string.Equals(token.Trim(), value, StringComparison.OrdinalIgnoreCase))
                    {
                        found.Add(new XmlValueMatch { Element = el, Attribute = a, Kind = "expression token", RawValue = raw, Path = PathOf(el) + (a != null ? "/@" + a.Name.LocalName : "/text()") });
                        return;
                    }
                }
            }
        }

        /// <summary>The nearest ancestor-or-self DisplayName (workflow step caption), if any.</summary>
        public static string NearestDisplayName(XElement element)
        {
            for (XElement e = element; e != null; e = e.Parent)
            {
                XAttribute d = e.Attribute("DisplayName");
                if (d != null && !string.IsNullOrWhiteSpace(d.Value)) return d.Value;
            }
            return null;
        }

        /// <summary>One-line rendering of a node for use as evidence (opening tag + trimmed value).</summary>
        public static string Render(XmlValueMatch m, int max = 400)
        {
            string text;
            if (m.Attribute != null)
                text = "<" + m.Element.Name.LocalName + " " + m.Attribute.Name.LocalName + "=\"" + m.RawValue + "\" ...>";
            else
                text = "<" + m.Element.Name.LocalName + ">" + m.RawValue.Trim() + "</" + m.Element.Name.LocalName + ">";
            text = DataverseHelper.OneLine(text);
            return text.Length > max ? text.Substring(0, max) + "..." : text;
        }
    }
}
