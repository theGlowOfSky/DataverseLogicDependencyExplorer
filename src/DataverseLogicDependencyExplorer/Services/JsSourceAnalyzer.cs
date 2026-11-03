// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace DataverseLogicDependencyExplorer.Services
{
    public enum JsSegmentKind { Code, Comment, StringLiteral, TemplateLiteral, RegexLiteral }

    public class JsSegment
    {
        public JsSegmentKind Kind { get; set; }
        public int Start { get; set; }
        public int Length { get; set; }
        public int End { get { return Start + Length; } }
    }

    /// <summary>How a name occurrence in JavaScript was found. None of these prove a runtime dependency.</summary>
    public enum JsMatchKind
    {
        /// <summary>String literal passed directly to a recognised Xrm/WebAPI call (getAttribute("x"), retrieveRecord("x") ...).</summary>
        ApiCallArgument,
        /// <summary>Name inside an OData / FetchXML / Web API URL string ($select=..., &lt;attribute name='x'/&gt;, /api/data/...).</summary>
        QueryOrFetchToken,
        /// <summary>Name inside an ordinary string literal.</summary>
        StringLiteral,
        /// <summary>Name used as an identifier or property in code (record.new_x).</summary>
        Identifier,
        /// <summary>Name only in a comment - listed only when the user opts in.</summary>
        Comment
    }

    public class JsMatch
    {
        public int Index { get; set; }
        public int Length { get; set; }
        public int Line { get; set; }
        public JsMatchKind Kind { get; set; }
        public string ApiName { get; set; }
        public string Snippet { get; set; }
        public string MatchedText { get; set; }
    }

    public class JsAnalysis
    {
        public List<JsMatch> Matches { get; } = new List<JsMatch>();
        /// <summary>Matches that exist only in comments and were therefore not returned (unless comments were requested).</summary>
        public int CommentOnlyMatches { get; set; }
        public int LineCount { get; set; }
    }

    public class DynamicAccessInfo
    {
        public int Count { get; set; }
        public int FirstLine { get; set; }
        public string FirstSnippet { get; set; }
    }

    /// <summary>
    /// Lightweight, dependency-free static analysis of JavaScript source. It separates code from comments and string
    /// literals so that a name appearing in a comment or an unrelated identifier is not mistaken for a reference.
    /// This is NOT a full JavaScript parser: minified or heavily obfuscated code, computed property names and
    /// string-building cannot be resolved and are reported as such by the callers.
    /// </summary>
    public static class JsSourceAnalyzer
    {
        private static readonly Regex ColumnApiBefore = new Regex(
            @"(?<api>getAttribute|getControl|getTab|getSection|attributes\.get|controls\.get|tabs\.get|sections\.get|getSelectedAttribute|setFieldNotification|clearFieldNotification)\s*\(\s*$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex TableApiBefore = new Regex(
            @"(?<api>retrieveRecord|retrieveMultipleRecords|createRecord|updateRecord|deleteRecord|retrieveMultiple|openForm|lookupObjects|createRecord|getEntityPrimaryIdAttribute|entityName|entityType|entityLogicalName|logicalName|etn|typeName)\s*[\(:=]\s*$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex QueryHint = new Regex(
            @"\$select|\$filter|\$expand|\$orderby|\$apply|<\s*attribute\b|<\s*entity\b|<\s*condition\b|<\s*order\b|<\s*link-entity\b|/api/data/|fetchXml|savedquery|userquery",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex DynamicCall = new Regex(
            @"\b(?:getAttribute|getControl|attributes\.get|controls\.get)\s*\(",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly HashSet<string> RegexAllowedAfterWord = new HashSet<string>(StringComparer.Ordinal)
        {
            "return", "typeof", "instanceof", "in", "of", "new", "delete", "void", "throw", "case", "do", "else", "yield", "await"
        };

        private const string RegexAllowedAfterChars = "(,=:[!&|?{;+-*%<>~^";

        private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<string, List<JsSegment>> SegmentCache =
            new System.Runtime.CompilerServices.ConditionalWeakTable<string, List<JsSegment>>();

        /// <summary>Tokenises once per source string (the same library is analysed by several scanners).</summary>
        public static List<JsSegment> GetSegments(string source)
        {
            if (string.IsNullOrEmpty(source)) return new List<JsSegment>();
            List<JsSegment> segs;
            if (!SegmentCache.TryGetValue(source, out segs))
            {
                segs = Tokenize(source);
                try { SegmentCache.Add(source, segs); } catch (ArgumentException) { }
            }
            return segs;
        }

        // ------------------------------------------------------------------ tokenizer

        public static List<JsSegment> Tokenize(string s)
        {
            var segs = new List<JsSegment>();
            if (string.IsNullOrEmpty(s)) return segs;

            int n = s.Length, i = 0, codeStart = 0;
            char prevSig = '\0';
            string prevWord = string.Empty;

            Action<int, int, JsSegmentKind> addSeg = (start, end, kind) =>
            {
                if (start > codeStart) segs.Add(new JsSegment { Kind = JsSegmentKind.Code, Start = codeStart, Length = start - codeStart });
                segs.Add(new JsSegment { Kind = kind, Start = start, Length = end - start });
                codeStart = end;
            };

            while (i < n)
            {
                char c = s[i];

                if (c == '/' && i + 1 < n && s[i + 1] == '/')
                {
                    int e = i + 2;
                    while (e < n && s[e] != '\n' && s[e] != '\r') e++;
                    addSeg(i, e, JsSegmentKind.Comment);
                    i = e;
                    continue;
                }
                if (c == '/' && i + 1 < n && s[i + 1] == '*')
                {
                    int e = s.IndexOf("*/", i + 2, StringComparison.Ordinal);
                    e = e < 0 ? n : e + 2;
                    addSeg(i, e, JsSegmentKind.Comment);
                    i = e;
                    continue;
                }
                if (c == '"' || c == '\'')
                {
                    int j = i + 1;
                    while (j < n && s[j] != c)
                    {
                        if (s[j] == '\\') { j += 2; continue; }
                        if (s[j] == '\n' || s[j] == '\r') break; // unterminated string: stop at the line end
                        j++;
                    }
                    int e = (j < n && s[j] == c) ? j + 1 : Math.Min(j, n);
                    addSeg(i, e, JsSegmentKind.StringLiteral);
                    i = e;
                    prevSig = 'a'; prevWord = string.Empty;
                    continue;
                }
                if (c == '`')
                {
                    int j = i + 1;
                    while (j < n && s[j] != '`') { if (s[j] == '\\') j++; j++; }
                    int e = j < n ? j + 1 : n;
                    addSeg(i, e, JsSegmentKind.TemplateLiteral);
                    i = e;
                    prevSig = 'a'; prevWord = string.Empty;
                    continue;
                }
                if (c == '/' && RegexAllowed(prevSig, prevWord))
                {
                    int j = i + 1;
                    bool inClass = false, closed = false;
                    while (j < n && s[j] != '\n' && s[j] != '\r')
                    {
                        if (s[j] == '\\') { j += 2; continue; }
                        if (s[j] == '[') inClass = true;
                        else if (s[j] == ']') inClass = false;
                        else if (s[j] == '/' && !inClass) { closed = true; break; }
                        j++;
                    }
                    if (closed)
                    {
                        int e = j + 1;
                        while (e < n && char.IsLetter(s[e])) e++; // flags
                        addSeg(i, e, JsSegmentKind.RegexLiteral);
                        i = e;
                        prevSig = 'a'; prevWord = string.Empty;
                        continue;
                    }
                }

                if (TextLocator.IsIdentifierChar(c))
                {
                    int j = i;
                    while (j < n && TextLocator.IsIdentifierChar(s[j])) j++;
                    prevWord = s.Substring(i, j - i);
                    prevSig = 'a';
                    i = j;
                    continue;
                }
                if (!char.IsWhiteSpace(c)) { prevSig = c; prevWord = string.Empty; }
                i++;
            }

            if (codeStart < n) segs.Add(new JsSegment { Kind = JsSegmentKind.Code, Start = codeStart, Length = n - codeStart });
            return segs;
        }

        private static bool RegexAllowed(char prevSig, string prevWord)
        {
            if (prevSig == '\0') return true;
            if (prevSig == 'a') return !string.IsNullOrEmpty(prevWord) && RegexAllowedAfterWord.Contains(prevWord);
            return RegexAllowedAfterChars.IndexOf(prevSig) >= 0;
        }

        // ------------------------------------------------------------------ searching

        /// <summary>Finds references to a column logical name.</summary>
        public static JsAnalysis FindColumn(string source, string column, bool includeComments)
        {
            return Find(source, new[] { column }, null, true, includeComments);
        }

        /// <summary>Finds references to a table logical name (and, optionally, its Web API entity set name).</summary>
        public static JsAnalysis FindTable(string source, string table, string entitySetName, bool includeComments)
        {
            return Find(source, new[] { table }, entitySetName, false, includeComments);
        }

        private static JsAnalysis Find(string source, string[] names, string entitySetName, bool isColumn, bool includeComments)
        {
            var analysis = new JsAnalysis();
            if (string.IsNullOrEmpty(source)) return analysis;

            string name = names[0];
            // Names without an underscore are generic ("name", "account"): bare identifiers/prose matches would be noise.
            bool distinctive = name.IndexOf('_') >= 0;
            string lookupForm = isColumn ? "_" + name + "_value" : null;

            var map = new LineMap(source);
            analysis.LineCount = map.LineCount;
            var seen = new HashSet<int>();

            foreach (JsSegment seg in GetSegments(source))
            {
                if (seg.Kind == JsSegmentKind.RegexLiteral) continue;

                int pos = seg.Start;
                string[] searchTerms = lookupForm == null ? new[] { name } : new[] { name, lookupForm };
                foreach (string term in searchTerms)
                {
                    pos = seg.Start;
                    while (true)
                    {
                        int idx = TextLocator.IndexOfWord(source, term, pos, seg.End);
                        if (idx < 0) break;
                        pos = idx + term.Length;
                        if (!seen.Add(idx)) continue;

                        JsMatchKind? kind = Classify(source, seg, idx, term.Length, isColumn, distinctive, term != name, out string api);
                        if (kind == null) continue;

                        if (kind == JsMatchKind.Comment && !includeComments) { analysis.CommentOnlyMatches++; continue; }

                        int line = map.LineOf(idx);
                        analysis.Matches.Add(new JsMatch
                        {
                            Index = idx,
                            Length = term.Length,
                            Line = line,
                            Kind = kind.Value,
                            ApiName = api,
                            Snippet = map.SnippetAt(idx, term.Length),
                            MatchedText = source.Substring(idx, term.Length)
                        });
                    }
                }

                if (!isColumn && !string.IsNullOrEmpty(entitySetName) && (seg.Kind == JsSegmentKind.StringLiteral || seg.Kind == JsSegmentKind.TemplateLiteral))
                {
                    pos = seg.Start;
                    while (true)
                    {
                        int idx = TextLocator.IndexOfWord(source, entitySetName, pos, seg.End);
                        if (idx < 0) break;
                        pos = idx + entitySetName.Length;
                        // Only a path segment of a Web API style URL ("/accounts?$select=", "/accounts(") counts.
                        bool afterSlash = idx > 0 && source[idx - 1] == '/';
                        if (!afterSlash || !seen.Add(idx)) continue;
                        analysis.Matches.Add(new JsMatch
                        {
                            Index = idx, Length = entitySetName.Length, Line = map.LineOf(idx),
                            Kind = JsMatchKind.QueryOrFetchToken, Snippet = map.SnippetAt(idx, entitySetName.Length),
                            MatchedText = source.Substring(idx, entitySetName.Length)
                        });
                    }
                }
            }

            analysis.Matches.Sort((a, b) => a.Index.CompareTo(b.Index));
            return analysis;
        }

        private static JsMatchKind? Classify(string source, JsSegment seg, int idx, int len, bool isColumn, bool distinctive, bool lookupForm, out string api)
        {
            api = null;

            if (seg.Kind == JsSegmentKind.Comment) return JsMatchKind.Comment;

            if (seg.Kind == JsSegmentKind.Code)
            {
                // Bare identifiers are only meaningful for distinctive (prefixed/custom) names.
                return distinctive ? (JsMatchKind?)JsMatchKind.Identifier : null;
            }

            // String or template literal.
            string content = StringContent(source, seg);
            bool exactLiteral = string.Equals(content.Trim(), source.Substring(idx, len), StringComparison.OrdinalIgnoreCase);

            if (exactLiteral && seg.Kind == JsSegmentKind.StringLiteral)
            {
                string before = CodeBefore(source, seg.Start, 120);
                Regex rx = isColumn ? ColumnApiBefore : TableApiBefore;
                Match m = rx.Match(before);
                if (m.Success) { api = m.Groups["api"].Value; return JsMatchKind.ApiCallArgument; }
            }

            if (QueryHint.IsMatch(content)) return JsMatchKind.QueryOrFetchToken;
            if (lookupForm) return JsMatchKind.QueryOrFetchToken;

            // Plain string: prose such as "Name is required" must not count for generic names.
            if (distinctive || exactLiteral) return JsMatchKind.StringLiteral;
            return null;
        }

        private static string StringContent(string source, JsSegment seg)
        {
            int s = seg.Start + 1;
            int e = seg.End - 1;
            if (e < s) return string.Empty;
            return source.Substring(s, Math.Max(0, Math.Min(source.Length, e) - s));
        }

        private static string CodeBefore(string source, int index, int maxChars)
        {
            int from = Math.Max(0, index - maxChars);
            return source.Substring(from, index - from);
        }

        // ------------------------------------------------------------------ dynamic access

        /// <summary>
        /// Counts attribute/control lookups whose argument is not a plain string literal (variables, concatenation). Such calls
        /// cannot be tied to a column by static analysis, so they are reported as possible dependencies.
        /// </summary>
        public static DynamicAccessInfo FindDynamicAccess(string source)
        {
            var info = new DynamicAccessInfo();
            if (string.IsNullOrEmpty(source)) return info;

            List<JsSegment> segs = GetSegments(source);
            var map = new LineMap(source);

            foreach (Match m in DynamicCall.Matches(source))
            {
                JsSegment seg = SegmentAt(segs, m.Index);
                if (seg == null || seg.Kind != JsSegmentKind.Code) continue;

                int p = m.Index + m.Length;
                while (p < source.Length && char.IsWhiteSpace(source[p])) p++;
                if (p >= source.Length) continue;

                bool dynamicArg;
                char c = source[p];
                if (c == '"' || c == '\'')
                {
                    // literal: dynamic only when it is concatenated ("new_" + suffix)
                    JsSegment lit = SegmentAt(segs, p);
                    int q = lit == null ? p + 1 : lit.End;
                    while (q < source.Length && char.IsWhiteSpace(source[q])) q++;
                    dynamicArg = q < source.Length && source[q] == '+';
                }
                else if (c == '`') dynamicArg = true;
                else dynamicArg = true; // identifier/expression/no argument

                if (!dynamicArg) continue;

                info.Count++;
                if (info.FirstLine == 0)
                {
                    info.FirstLine = map.LineOf(m.Index);
                    info.FirstSnippet = map.SnippetAt(m.Index, m.Length);
                }
            }
            return info;
        }

        // ------------------------------------------------------------------ function bodies

        /// <summary>
        /// Locates the body of a named function (declaration, assigned function expression, arrow function or method shorthand).
        /// For a namespaced name such as "Contoso.Account.onLoad" the last segment is used. Returns false when the function
        /// cannot be located reliably; callers must then fall back to library-level evidence.
        /// </summary>
        public static bool TryFindFunctionBody(string source, string functionName, out int bodyStart, out int bodyEnd)
        {
            bodyStart = bodyEnd = -1;
            if (string.IsNullOrEmpty(source) || string.IsNullOrWhiteSpace(functionName)) return false;

            string name = functionName.Trim();
            int dot = name.LastIndexOf('.');
            if (dot >= 0) name = name.Substring(dot + 1);
            if (name.Length == 0) return false;

            string n = Regex.Escape(name);
            var patterns = new[]
            {
                @"\bfunction\s+" + n + @"\s*\([^)]*\)\s*\{",
                @"\b" + n + @"\s*[:=]\s*(?:async\s+)?function\s*\*?\s*(?:\w+\s*)?\([^)]*\)\s*\{",
                @"\b" + n + @"\s*[:=]\s*(?:async\s*)?\([^)]*\)\s*=>\s*\{",
                @"\b" + n + @"\s*[:=]\s*(?:async\s+)?\w+\s*=>\s*\{",
                @"(?<=[{,;\s])" + n + @"\s*\([^)]*\)\s*\{"
            };

            List<JsSegment> segs = GetSegments(source);
            foreach (string pattern in patterns)
            {
                foreach (Match m in Regex.Matches(source, pattern))
                {
                    JsSegment seg = SegmentAt(segs, m.Index);
                    if (seg == null || seg.Kind != JsSegmentKind.Code) continue;
                    int open = m.Index + m.Length - 1;
                    int close = MatchBrace(source, segs, open);
                    if (close < 0) continue;
                    bodyStart = open;
                    bodyEnd = close;
                    return true;
                }
            }
            return false;
        }

        private static int MatchBrace(string source, List<JsSegment> segs, int openIndex)
        {
            int depth = 0;
            foreach (JsSegment seg in segs)
            {
                if (seg.End <= openIndex || seg.Kind != JsSegmentKind.Code) continue;
                for (int i = Math.Max(seg.Start, openIndex); i < seg.End; i++)
                {
                    char c = source[i];
                    if (c == '{') depth++;
                    else if (c == '}')
                    {
                        depth--;
                        if (depth == 0) return i;
                    }
                }
            }
            return -1;
        }

        private static JsSegment SegmentAt(List<JsSegment> segs, int index)
        {
            int lo = 0, hi = segs.Count - 1;
            while (lo <= hi)
            {
                int mid = (lo + hi) / 2;
                JsSegment s = segs[mid];
                if (index < s.Start) hi = mid - 1;
                else if (index >= s.End) lo = mid + 1;
                else return s;
            }
            return null;
        }
    }
}
