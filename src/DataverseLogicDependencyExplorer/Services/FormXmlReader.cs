// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace DataverseLogicDependencyExplorer.Services
{
    public class FormHandler
    {
        /// <summary>onload, onsave, onchange, tabstatechange, onrecordselect ...</summary>
        public string EventName { get; set; }
        /// <summary>The field an OnChange handler is registered against (null for form-level events).</summary>
        public string Attribute { get; set; }
        /// <summary>The control id when the event is registered on a specific control (e.g. a subgrid).</summary>
        public string ControlId { get; set; }
        public string FunctionName { get; set; }
        public string LibraryName { get; set; }
        public bool Enabled { get; set; }
        public bool PassExecutionContext { get; set; }
        public string Parameters { get; set; }
        public string Location { get; set; }
    }

    public class FormControl
    {
        public string Id { get; set; }
        public string DataFieldName { get; set; }
        public string Location { get; set; }
    }

    public class FormDefinition
    {
        /// <summary>Set when the form XML was empty or malformed. Scanners must report it instead of treating the form as empty.</summary>
        public string Error { get; set; }
        public XDocument Document { get; set; }
        public List<string> Libraries { get; } = new List<string>();
        public List<FormHandler> Handlers { get; } = new List<FormHandler>();
        public List<FormControl> Controls { get; } = new List<FormControl>();
        public bool IsValid { get { return Error == null; } }
    }

    /// <summary>Parses systemform.formxml: registered libraries, event handlers and field controls, with exact node paths.</summary>
    public static class FormXmlReader
    {
        public static FormDefinition Parse(string formXml)
        {
            var def = new FormDefinition();
            if (string.IsNullOrWhiteSpace(formXml))
            {
                def.Error = "The form has no form XML.";
                return def;
            }

            try
            {
                def.Document = XDocument.Parse(formXml);
            }
            catch (Exception ex)
            {
                def.Error = "The form XML is malformed: " + ex.Message;
                return def;
            }

            foreach (XElement lib in def.Document.Descendants().Where(e => e.Name.LocalName.Equals("Library", StringComparison.OrdinalIgnoreCase)
                                                                          && e.Parent != null
                                                                          && e.Parent.Name.LocalName.Equals("formLibraries", StringComparison.OrdinalIgnoreCase)))
            {
                string name = (string)lib.Attribute("name");
                if (!string.IsNullOrWhiteSpace(name)) def.Libraries.Add(name.Trim());
            }

            foreach (XElement ev in def.Document.Descendants().Where(e => e.Name.LocalName.Equals("event", StringComparison.OrdinalIgnoreCase)))
            {
                string eventName = ((string)ev.Attribute("name") ?? string.Empty).Trim();
                string attribute = (string)ev.Attribute("attribute");
                XElement control = ev.Ancestors().FirstOrDefault(a => a.Name.LocalName.Equals("control", StringComparison.OrdinalIgnoreCase));
                string controlId = control == null ? null : (string)control.Attribute("id");

                foreach (XElement h in ev.Descendants().Where(e => e.Name.LocalName.Equals("Handler", StringComparison.OrdinalIgnoreCase)))
                {
                    def.Handlers.Add(new FormHandler
                    {
                        EventName = eventName,
                        Attribute = string.IsNullOrWhiteSpace(attribute) ? null : attribute.Trim(),
                        ControlId = controlId,
                        FunctionName = ((string)h.Attribute("functionName") ?? string.Empty).Trim(),
                        LibraryName = ((string)h.Attribute("libraryName") ?? string.Empty).Trim(),
                        Enabled = !string.Equals((string)h.Attribute("enabled"), "false", StringComparison.OrdinalIgnoreCase),
                        PassExecutionContext = string.Equals((string)h.Attribute("passExecutionContext"), "true", StringComparison.OrdinalIgnoreCase),
                        Parameters = (string)h.Attribute("parameters"),
                        Location = XmlPath.PathOf(h)
                    });
                }
            }

            foreach (XElement c in def.Document.Descendants().Where(e => e.Name.LocalName.Equals("control", StringComparison.OrdinalIgnoreCase)))
            {
                string field = (string)c.Attribute("datafieldname");
                if (string.IsNullOrWhiteSpace(field)) continue;
                def.Controls.Add(new FormControl { Id = (string)c.Attribute("id"), DataFieldName = field.Trim(), Location = XmlPath.PathOf(c) });
            }

            return def;
        }

        /// <summary>
        /// Exact-value references elsewhere in the form XML (subgrid target table, quick view table...) that are not
        /// already reported as a field control or an event registration.
        /// </summary>
        public static List<XmlValueMatch> FindOtherReferences(FormDefinition def, string value)
        {
            if (def == null || def.Document == null) return new List<XmlValueMatch>();
            return XmlPath.FindExactValues(def.Document, value, (el, attr) =>
            {
                string elName = el.Name.LocalName;
                string attrName = attr.Name.LocalName;
                // A control's id normally repeats its field name; both belong to the control that is already reported.
                if (elName.Equals("control", StringComparison.OrdinalIgnoreCase) &&
                    (attrName.Equals("datafieldname", StringComparison.OrdinalIgnoreCase) || attrName.Equals("id", StringComparison.OrdinalIgnoreCase))) return true;
                if (elName.Equals("event", StringComparison.OrdinalIgnoreCase) && attrName.Equals("attribute", StringComparison.OrdinalIgnoreCase)) return true;
                return false;
            });
        }

        public static string DescribeFormType(int? type)
        {
            switch (type)
            {
                case 2: return "Main";
                case 5: return "Mobile - Express";
                case 6: return "Quick View";
                case 7: return "Quick Create";
                case 11: return "Card";
                case 12: return "Main - Interactive experience";
                default: return type.HasValue ? "Type " + type.Value : "Unknown";
            }
        }
    }
}
