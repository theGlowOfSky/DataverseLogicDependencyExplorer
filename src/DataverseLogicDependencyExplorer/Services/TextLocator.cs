// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;

namespace DataverseLogicDependencyExplorer.Services
{
    /// <summary>Identifier-aware searching and accurate line numbers for source text.</summary>
    public static class TextLocator
    {
        public static bool IsIdentifierChar(char c)
        {
            return char.IsLetterOrDigit(c) || c == '_' || c == '$';
        }

        /// <summary>
        /// Index of the next occurrence of <paramref name="word"/> that is a whole identifier (not part of a longer
        /// identifier such as "new_accountid" when searching for "account"). Case-insensitive. -1 when none.
        /// </summary>
        public static int IndexOfWord(string text, string word, int start, int end)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(word)) return -1;
            if (end > text.Length) end = text.Length;
            int i = Math.Max(0, start);
            while (i < end)
            {
                int idx = text.IndexOf(word, i, end - i, StringComparison.OrdinalIgnoreCase);
                if (idx < 0) return -1;
                bool leftOk = idx == 0 || !IsIdentifierChar(text[idx - 1]);
                int after = idx + word.Length;
                bool rightOk = after >= text.Length || !IsIdentifierChar(text[after]);
                if (leftOk && rightOk) return idx;
                i = idx + 1;
            }
            return -1;
        }

        public static int IndexOfWord(string text, string word, int start)
        {
            return text == null ? -1 : IndexOfWord(text, word, start, text.Length);
        }

        public static bool ContainsWord(string text, string word)
        {
            return IndexOfWord(text, word, 0) >= 0;
        }

        /// <summary>Splits a comma/semicolon/space separated list ("name,telephone1") into trimmed items.</summary>
        public static List<string> SplitList(string list)
        {
            var result = new List<string>();
            if (string.IsNullOrWhiteSpace(list)) return result;
            foreach (string part in list.Split(new[] { ',', ';', ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
                result.Add(part.Trim());
            return result;
        }

        public static bool ListContains(string list, string item)
        {
            foreach (string part in SplitList(list))
                if (string.Equals(part, item, StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }
    }

    /// <summary>Maps character offsets to 1-based line numbers (and line text) with one pass over the text.</summary>
    public class LineMap
    {
        private readonly string text;
        private readonly List<int> starts = new List<int>();

        public LineMap(string text)
        {
            this.text = text ?? string.Empty;
            starts.Add(0);
            for (int i = 0; i < this.text.Length; i++)
            {
                char c = this.text[i];
                if (c == '\n') starts.Add(i + 1);
                else if (c == '\r' && (i + 1 >= this.text.Length || this.text[i + 1] != '\n')) starts.Add(i + 1);
            }
        }

        public int LineCount { get { return starts.Count; } }

        public int LineOf(int index)
        {
            int lo = 0, hi = starts.Count - 1;
            while (lo < hi)
            {
                int mid = (lo + hi + 1) / 2;
                if (starts[mid] <= index) lo = mid; else hi = mid - 1;
            }
            return lo + 1;
        }

        public string LineText(int line)
        {
            if (line < 1 || line > starts.Count) return string.Empty;
            int s = starts[line - 1];
            int e = line < starts.Count ? starts[line] : text.Length;
            return text.Substring(s, e - s).TrimEnd('\r', '\n');
        }

        /// <summary>The source line containing <paramref name="index"/>, trimmed and limited to a readable window around the match.</summary>
        public string SnippetAt(int index, int length, int maxChars = 300)
        {
            int line = LineOf(index);
            int s = starts[line - 1];
            string lineText = LineText(line);
            int col = index - s;
            if (lineText.Length <= maxChars) return lineText.Trim();
            int from = Math.Max(0, col - maxChars / 2);
            int len = Math.Min(maxChars, lineText.Length - from);
            string window = lineText.Substring(from, len).Trim();
            return (from > 0 ? "..." : string.Empty) + window + (from + len < lineText.Length ? "..." : string.Empty);
        }
    }
}
