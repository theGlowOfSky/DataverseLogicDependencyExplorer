// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;

namespace DataverseLogicDependencyExplorer.Reporting
{
    public enum XlsxStyle { Default = 0, Header = 1, Wrap = 2, DateTime = 3, Bold = 4, Title = 5 }

    /// <summary>One worksheet. Rows are lists of values: string, int/long/double, DateTime, bool or null.</summary>
    public class XlsxSheet
    {
        internal readonly List<KeyValuePair<object[], XlsxStyle?>> Rows = new List<KeyValuePair<object[], XlsxStyle?>>();
        public string Name { get; internal set; }
        /// <summary>Number of top rows kept visible when scrolling (0 = none).</summary>
        public int FreezeRows { get; set; }
        /// <summary>1-based header row that receives an auto-filter covering the data below it (0 = no filter).</summary>
        public int AutoFilterHeaderRow { get; set; }
        public Dictionary<int, double> ColumnWidths { get; } = new Dictionary<int, double>();

        public void AddRow(IEnumerable<object> values, XlsxStyle? style = null)
        {
            Rows.Add(new KeyValuePair<object[], XlsxStyle?>(values.ToArray(), style));
        }

        public void AddRow(params object[] values) { AddRow((IEnumerable<object>)values, null); }
        public int RowCount { get { return Rows.Count; } }
    }

    /// <summary>
    /// Minimal, dependency-free .xlsx writer (OpenXML packaged with System.IO.Compression), extending the approach of the
    /// Business Rule Exporter's XlsxWriter to several worksheets. Text is always stored as a string, so values that start with
    /// '=', '+', '-' or '@' are never interpreted as formulas.
    /// </summary>
    public class XlsxWorkbookWriter
    {
        public const int MaxCellLength = 32767;

        private readonly List<XlsxSheet> sheets = new List<XlsxSheet>();
        private readonly List<string> strings = new List<string>();
        private readonly Dictionary<string, int> stringIndex = new Dictionary<string, int>(StringComparer.Ordinal);

        /// <summary>Cells whose text exceeded Excel's per-cell limit and was cut (with a visible marker).</summary>
        public int TruncatedCells { get; private set; }

        public XlsxSheet AddSheet(string name)
        {
            var sheet = new XlsxSheet { Name = SafeSheetName(name) };
            string baseName = sheet.Name;
            int n = 2;
            while (sheets.Any(s => string.Equals(s.Name, sheet.Name, StringComparison.OrdinalIgnoreCase)))
            {
                string suffix = " (" + n++ + ")";
                sheet.Name = baseName.Substring(0, Math.Min(baseName.Length, 31 - suffix.Length)) + suffix;
            }
            sheets.Add(sheet);
            return sheet;
        }

        public IReadOnlyList<XlsxSheet> Sheets { get { return sheets; } }

        public static string SafeSheetName(string name)
        {
            string n = string.IsNullOrWhiteSpace(name) ? "Sheet" : name;
            foreach (char c in new[] { '[', ']', ':', '*', '?', '/', '\\' }) n = n.Replace(c, '-');
            n = n.Trim('\'');
            if (n.Length == 0) n = "Sheet";
            return n.Length > 31 ? n.Substring(0, 31) : n;
        }

        /// <summary>Removes characters that are illegal in XML 1.0 (control characters, unpaired surrogates).</summary>
        public static string CleanText(string s)
        {
            if (string.IsNullOrEmpty(s)) return s ?? string.Empty;
            var sb = new StringBuilder(s.Length);
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '\t' || c == '\n' || c == '\r' || (c >= 0x20 && c <= 0xD7FF) || (c >= 0xE000 && c <= 0xFFFD)) sb.Append(c);
                else if (char.IsHighSurrogate(c) && i + 1 < s.Length && char.IsLowSurrogate(s[i + 1])) { sb.Append(c).Append(s[i + 1]); i++; }
            }
            return sb.ToString();
        }

        public static string ColumnName(int index) // 0-based
        {
            var sb = new StringBuilder();
            for (int n = index + 1; n > 0; n = (n - 1) / 26) sb.Insert(0, (char)('A' + (n - 1) % 26));
            return sb.ToString();
        }

        public void Save(string path)
        {
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
                Save(fs);
        }

        public void Save(Stream output)
        {
            if (sheets.Count == 0) throw new InvalidOperationException("A workbook needs at least one worksheet.");
            strings.Clear(); stringIndex.Clear(); TruncatedCells = 0;

            var sheetXml = sheets.Select(BuildSheetXml).ToList(); // builds the shared string table as a side effect

            using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true))
            {
                Write(zip, "[Content_Types].xml", ContentTypes());
                Write(zip, "_rels/.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
                Write(zip, "xl/workbook.xml", WorkbookXml());
                Write(zip, "xl/_rels/workbook.xml.rels", WorkbookRels());
                Write(zip, "xl/styles.xml", StylesXml());
                Write(zip, "xl/sharedStrings.xml", SharedStringsXml());
                for (int i = 0; i < sheetXml.Count; i++) Write(zip, "xl/worksheets/sheet" + (i + 1) + ".xml", sheetXml[i]);
            }
        }

        private static void Write(ZipArchive zip, string name, string content)
        {
            ZipArchiveEntry e = zip.CreateEntry(name, CompressionLevel.Optimal);
            using (var w = new StreamWriter(e.Open(), new UTF8Encoding(false))) w.Write(content);
        }

        private int Intern(string text)
        {
            int idx;
            if (!stringIndex.TryGetValue(text, out idx)) { idx = strings.Count; strings.Add(text); stringIndex[text] = idx; }
            return idx;
        }

        private string Prepare(string s)
        {
            string t = CleanText(s);
            if (t.Length > MaxCellLength)
            {
                TruncatedCells++;
                const string marker = " ...[TRUNCATED: exceeds Excel's 32,767 character cell limit]";
                t = t.Substring(0, MaxCellLength - marker.Length) + marker;
            }
            return t;
        }

        private string BuildSheetXml(XlsxSheet sheet)
        {
            int cols = sheet.Rows.Count == 0 ? 1 : Math.Max(1, sheet.Rows.Max(r => r.Key.Length));
            var widths = new double[cols];
            foreach (var row in sheet.Rows)
                for (int c = 0; c < row.Key.Length; c++)
                {
                    object v = row.Key[c];
                    if (v == null) continue;
                    string s = v is DateTime ? "0000-00-00 00:00:00" : Convert.ToString(v, CultureInfo.InvariantCulture);
                    int longestLine = (s ?? string.Empty).Split('\n').Max(l => l.Length);
                    widths[c] = Math.Max(widths[c], Math.Min(longestLine + 2, 70));
                }

            var sb = new StringBuilder();
            // The declaration is written by hand: a StringWriter would otherwise announce utf-16 while the part is stored as UTF-8.
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            var settings = new XmlWriterSettings { OmitXmlDeclaration = true, CheckCharacters = false };
            using (var sw = new StringWriter(sb))
            using (var x = XmlWriter.Create(sw, settings))
            {
                const string ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                x.WriteStartElement("worksheet", ns);

                x.WriteStartElement("sheetViews");
                x.WriteStartElement("sheetView");
                x.WriteAttributeString("workbookViewId", "0");
                if (sheet.FreezeRows > 0)
                {
                    x.WriteStartElement("pane");
                    x.WriteAttributeString("ySplit", sheet.FreezeRows.ToString(CultureInfo.InvariantCulture));
                    x.WriteAttributeString("topLeftCell", "A" + (sheet.FreezeRows + 1));
                    x.WriteAttributeString("activePane", "bottomLeft");
                    x.WriteAttributeString("state", "frozen");
                    x.WriteEndElement();
                }
                x.WriteEndElement();
                x.WriteEndElement();

                x.WriteStartElement("sheetFormatPr");
                x.WriteAttributeString("defaultRowHeight", "15");
                x.WriteEndElement();

                x.WriteStartElement("cols");
                for (int c = 0; c < cols; c++)
                {
                    double w;
                    if (!sheet.ColumnWidths.TryGetValue(c, out w)) w = Math.Max(10, widths[c]);
                    x.WriteStartElement("col");
                    x.WriteAttributeString("min", (c + 1).ToString(CultureInfo.InvariantCulture));
                    x.WriteAttributeString("max", (c + 1).ToString(CultureInfo.InvariantCulture));
                    x.WriteAttributeString("width", w.ToString("0.##", CultureInfo.InvariantCulture));
                    x.WriteAttributeString("customWidth", "1");
                    x.WriteEndElement();
                }
                x.WriteEndElement();

                x.WriteStartElement("sheetData");
                for (int r = 0; r < sheet.Rows.Count; r++)
                {
                    object[] values = sheet.Rows[r].Key;
                    XlsxStyle? rowStyle = sheet.Rows[r].Value;
                    x.WriteStartElement("row");
                    x.WriteAttributeString("r", (r + 1).ToString(CultureInfo.InvariantCulture));
                    for (int c = 0; c < values.Length; c++)
                    {
                        object v = values[c];
                        string cellRef = ColumnName(c) + (r + 1);
                        if (v == null)
                        {
                            if (rowStyle == XlsxStyle.Header)
                            {
                                x.WriteStartElement("c"); x.WriteAttributeString("r", cellRef); x.WriteAttributeString("s", "1"); x.WriteEndElement();
                            }
                            continue;
                        }
                        x.WriteStartElement("c");
                        x.WriteAttributeString("r", cellRef);
                        if (v is DateTime)
                        {
                            x.WriteAttributeString("s", ((int)XlsxStyle.DateTime).ToString());
                            x.WriteElementString("v", ((DateTime)v).ToOADate().ToString("R", CultureInfo.InvariantCulture));
                        }
                        else if (v is int || v is long || v is double || v is decimal || v is short)
                        {
                            if (rowStyle.HasValue) x.WriteAttributeString("s", ((int)rowStyle.Value).ToString());
                            x.WriteElementString("v", Convert.ToString(v, CultureInfo.InvariantCulture));
                        }
                        else
                        {
                            string text = v is bool ? ((bool)v ? "Yes" : "No") : Prepare(Convert.ToString(v, CultureInfo.InvariantCulture));
                            XlsxStyle style = rowStyle ?? (text.IndexOf('\n') >= 0 || text.Length > 70 ? XlsxStyle.Wrap : XlsxStyle.Default);
                            x.WriteAttributeString("s", ((int)style).ToString());
                            x.WriteAttributeString("t", "s");
                            x.WriteElementString("v", Intern(text).ToString(CultureInfo.InvariantCulture));
                        }
                        x.WriteEndElement();
                    }
                    x.WriteEndElement();
                }
                x.WriteEndElement(); // sheetData

                if (sheet.AutoFilterHeaderRow > 0 && sheet.Rows.Count >= sheet.AutoFilterHeaderRow)
                {
                    x.WriteStartElement("autoFilter");
                    x.WriteAttributeString("ref", "A" + sheet.AutoFilterHeaderRow + ":" + ColumnName(cols - 1) + Math.Max(sheet.Rows.Count, sheet.AutoFilterHeaderRow));
                    x.WriteEndElement();
                }
                x.WriteEndElement(); // worksheet
            }
            return sb.ToString();
        }

        private string SharedStringsXml()
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"" + strings.Count + "\" uniqueCount=\"" + strings.Count + "\">");
            foreach (string s in strings)
                sb.Append("<si><t xml:space=\"preserve\">").Append(System.Security.SecurityElement.Escape(s)).Append("</t></si>");
            sb.Append("</sst>");
            return sb.ToString();
        }

        private string ContentTypes()
        {
            var sb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/><Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>");
            for (int i = 1; i <= sheets.Count; i++)
                sb.Append("<Override PartName=\"/xl/worksheets/sheet" + i + ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            return sb.Append("</Types>").ToString();
        }

        private string WorkbookXml()
        {
            var sb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>");
            for (int i = 0; i < sheets.Count; i++)
                sb.Append("<sheet name=\"" + System.Security.SecurityElement.Escape(sheets[i].Name) + "\" sheetId=\"" + (i + 1) + "\" r:id=\"rId" + (i + 1) + "\"/>");
            return sb.Append("</sheets></workbook>").ToString();
        }

        private string WorkbookRels()
        {
            var sb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            for (int i = 1; i <= sheets.Count; i++)
                sb.Append("<Relationship Id=\"rId" + i + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet" + i + ".xml\"/>");
            sb.Append("<Relationship Id=\"rId" + (sheets.Count + 1) + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
            sb.Append("<Relationship Id=\"rId" + (sheets.Count + 2) + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>");
            return sb.Append("</Relationships>").ToString();
        }

        private static string StylesXml()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                + "<numFmts count=\"1\"><numFmt numFmtId=\"164\" formatCode=\"yyyy\\-mm\\-dd\\ hh:mm:ss\"/></numFmts>"
                + "<fonts count=\"3\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font><font><b/><sz val=\"11\"/><name val=\"Calibri\"/></font><font><b/><sz val=\"14\"/><name val=\"Calibri\"/></font></fonts>"
                + "<fills count=\"3\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill><fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFDCE6F1\"/><bgColor indexed=\"64\"/></patternFill></fill></fills>"
                + "<borders count=\"2\"><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top/><bottom style=\"thin\"><color auto=\"1\"/></bottom><diagonal/></border></borders>"
                + "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
                + "<cellXfs count=\"6\">"
                + "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment vertical=\"top\"/></xf>"
                + "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment vertical=\"top\" wrapText=\"1\"/></xf>"
                + "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment vertical=\"top\" wrapText=\"1\"/></xf>"
                + "<xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyAlignment=\"1\"><alignment vertical=\"top\" horizontal=\"left\"/></xf>"
                + "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyAlignment=\"1\"><alignment vertical=\"top\"/></xf>"
                + "<xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>"
                + "</cellXfs>"
                + "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>"
                + "</styleSheet>";
        }
    }
}
