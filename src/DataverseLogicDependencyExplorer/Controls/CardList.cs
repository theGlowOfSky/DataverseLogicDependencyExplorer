// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace DataverseLogicDependencyExplorer.Controls
{
    public class CardItem
    {
        public string Key { get; set; }
        public string Title { get; set; }
        public string Subtitle { get; set; }
    }

    /// <summary>
    /// A vertical list of 62 px cards (title + italic subtitle) with single selection. The selected card has a light-blue
    /// background. Very long lists are capped (<see cref="MaxCards"/>) with a hint to refine the search, because thousands of
    /// controls would make filtering sluggish.
    /// </summary>
    public class CardList : FlowLayoutPanel
    {
        private static readonly Color SelectedBack = Color.FromArgb(236, 244, 252);
        private static readonly Color TitleColor = Color.FromArgb(45, 45, 45);

        private readonly List<Panel> cards = new List<Panel>();
        private List<CardItem> items = new List<CardItem>();
        private string selectedKey;
        private bool raising = true;

        public int MaxCards { get; set; } = 250;
        public int TotalItems { get { return items.Count; } }
        public int RenderedCards { get { return cards.Count; } }
        public string SelectedKey { get { return selectedKey; } }
        public CardItem SelectedItem { get { return items.Find(i => i.Key == selectedKey); } }
        public event EventHandler SelectionChanged;

        public CardList()
        {
            FlowDirection = FlowDirection.TopDown;
            WrapContents = false;
            AutoScroll = true;
            BackColor = Color.White;
            Padding = new Padding(5, 3, 5, 5);
            Margin = new Padding(0);
            Dock = DockStyle.Fill;
            // Cards are always narrower than the list, so a horizontal scrollbar is never needed (it only appears transiently while the width settles).
            HorizontalScroll.Enabled = false;
            HorizontalScroll.Visible = false;
        }

        /// <summary>Replaces the cards. The selection is kept when <paramref name="keepKey"/> is still among the items.</summary>
        public void SetItems(IList<CardItem> newItems, string keepKey)
        {
            SuspendLayout();
            try
            {
                for (int i = Controls.Count - 1; i >= 0; i--) { Control old = Controls[i]; Controls.RemoveAt(i); old.Dispose(); }
                cards.Clear();
                items = new List<CardItem>(newItems);
                selectedKey = keepKey != null && items.Exists(i => i.Key == keepKey) ? keepKey : null;

                int n = Math.Min(MaxCards, items.Count);
                for (int k = 0; k < n; k++)
                {
                    Panel card = BuildCard(items[k]);
                    cards.Add(card);
                    Controls.Add(card);
                }
                if (items.Count > n)
                {
                    Controls.Add(new Label
                    {
                        AutoSize = false, Height = 34, Width = UiLayout.CardWidth(ListClientWidth), ForeColor = Color.DimGray,
                        Font = new Font("Segoe UI", 8F, FontStyle.Italic),
                        Text = "Showing the first " + n + " of " + items.Count + ". Type in the search box to narrow the list."
                    });
                }
                ApplySelectionColors();
            }
            finally { ResumeLayout(true); }
        }

        /// <summary>Selects (or clears, with null) without raising <see cref="SelectionChanged"/>.</summary>
        public void SetSelected(string key)
        {
            selectedKey = key != null && items.Exists(i => i.Key == key) ? key : null;
            ApplySelectionColors();
        }

        /// <summary>Selects the card at an index as if the user had clicked it.</summary>
        public void ChooseIndex(int index)
        {
            if (index >= 0 && index < cards.Count) Choose((string)cards[index].Tag);
        }

        public void ClearSelection()
        {
            if (selectedKey == null) return;
            selectedKey = null;
            ApplySelectionColors();
            if (raising && SelectionChanged != null) SelectionChanged(this, EventArgs.Empty);
        }

        private Panel BuildCard(CardItem item)
        {
            int width = UiLayout.CardWidth(ListClientWidth);
            var card = new Panel
            {
                Height = UiLayout.CardHeight, Width = width, BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 0, 0, 4), BackColor = Color.White, Tag = item.Key
            };
            var title = new Label
            {
                Left = 10, Top = 7, Height = 20, Width = width - 18, AutoEllipsis = true, AutoSize = false,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = TitleColor, Text = item.Title, BackColor = Color.Transparent
            };
            var sub = new Label
            {
                Left = 10, Top = 31, Height = 19, Width = width - 18, AutoEllipsis = true, AutoSize = false,
                Font = new Font("Segoe UI", 8F, FontStyle.Italic), ForeColor = Color.DimGray, Text = item.Subtitle, BackColor = Color.Transparent
            };
            card.Controls.Add(title);
            card.Controls.Add(sub);
            EventHandler click = (s, e) => { Choose(item.Key); Focus(); };
            card.Click += click; title.Click += click; sub.Click += click;
            return card;
        }

        private void Choose(string key)
        {
            if (selectedKey == key) return;
            selectedKey = key;
            ApplySelectionColors();
            if (SelectionChanged != null) SelectionChanged(this, EventArgs.Empty);
        }

        private void ApplySelectionColors()
        {
            foreach (Panel c in cards) c.BackColor = (string)c.Tag == selectedKey ? SelectedBack : Color.White;
        }

        /// <summary>
        /// The width the spec formula starts from: the list's client width. On Windows ClientSize already excludes the vertical
        /// scrollbar; where it does not (Mono), the usable display width is used so a card never overflows the list.
        /// </summary>
        private int ListClientWidth
        {
            get { return Math.Min(ClientSize.Width, DisplayRectangle.Width + Padding.Horizontal); }
        }

        private void ResizeCards()
        {
            int width = UiLayout.CardWidth(ListClientWidth);
            foreach (Control c in Controls)
            {
                if (c.Width == width) continue;
                c.Width = width;
                foreach (Control child in c.Controls) child.Width = width - 18;
            }
        }

        protected override void OnClientSizeChanged(EventArgs e)
        {
            base.OnClientSizeChanged(e);
            HorizontalScroll.Visible = false;
            ResizeCards();
        }

        protected override void OnLayout(LayoutEventArgs levent)
        {
            base.OnLayout(levent);
            ResizeCards(); // a scrollbar appearing or disappearing changes the usable width without changing ClientSize
        }
    }
}
