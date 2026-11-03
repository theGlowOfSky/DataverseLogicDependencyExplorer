// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;

namespace DataverseLogicDependencyExplorer
{
    /// <summary>Exact sizing rules of the inventory-style layout (kept free of WinForms so they can be unit tested).</summary>
    public static class UiLayout
    {
        public const int SplitMinWidthForRatio = 700;
        public const int Panel1MinSize = 270;
        public const int Panel2MinSize = 400;
        public const int InitialSplitterDistance = 305;
        public const int LeftHeaderHeight = 176;
        public const int RightHeaderHeight = 78;
        public const int CardHeight = 62;

        /// <summary>
        /// Left pane width: 27 % of the split width, clamped to 280-350 px, and never so wide that the right pane drops below its
        /// minimum. Returns null when the split is narrower than 700 px (the width is then left alone).
        /// </summary>
        public static int? LeftPaneWidth(int splitWidth, int splitterWidth)
        {
            if (splitWidth < SplitMinWidthForRatio) return null;
            int w = (int)(splitWidth * 0.27);
            w = Math.Max(280, Math.Min(350, w));
            w = Math.Min(w, splitWidth - Panel2MinSize - splitterWidth);
            return w;
        }

        public static int CardWidth(int listClientWidth) { return Math.Max(215, listClientWidth - 25); }
        public static int HeaderInnerWidth(int panelWidth) { return Math.Max(230, panelWidth - 24); }
        public static int HalfButtonWidth(int innerWidth) { return Math.Max(85, (innerWidth - 7) / 2); }
        public static int StatusLabelWidth(int panelWidth) { return Math.Max(200, panelWidth - 20); }

        /// <summary>Single-line version of multi-line text for grid cells; the full text stays in the tooltip and the Excel export.</summary>
        public static string Collapse(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            var parts = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < parts.Length; i++) parts[i] = parts[i].Trim();
            return string.Join("  \u203A  ", Array.FindAll(parts, p => p.Length > 0));
        }
    }
}
