// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.ComponentModel;
using DataverseLogicDependencyExplorer.Controls;
using DataverseLogicDependencyExplorer.Models;
using DataverseLogicDependencyExplorer.Services;
using WinFormsLabel = System.Windows.Forms.Label;

namespace DataverseLogicDependencyExplorer
{
    /// <summary>Tab 1: Dependency Finder.</summary>
    public partial class MainControl
    {
        private class StatusRow
        {
            public string Category { get; set; }
            public string Status { get; set; }
            public string Examined { get; set; }
            public string Matches { get; set; }
            public string Notes { get; set; }
            public ScanStatus StatusCode { get; set; }
        }

        private SplitContainer split;
        private bool binding;

        private WinFormsLabel lblEnv, lblSearchTables, lblTableSummary, lblColumnHeader, lblSelection, lblValidation, lblProgress, lblSummary;
        private Button btnLoad, btnClearSelection, btnFind, btnCancel, btnClear, btnViewReport, btnScopeAll, btnScopeNone;
        private ComboBox cmbTableKind, cmbJsScope;
        private TextBox txtTableSearch, txtColumnSearch;
        private CardList cardsTables, cardsColumns;
        private DataGridView gridStatus;
        private readonly BindingSource statusSource = new BindingSource();
        private RadioButton rbTable, rbColumn;
        private CheckBox[] scopeChecks;
        private CheckBox chkComments, chkDeep, chkSkipManaged;
        private ProgressBar bar;
        private LinkLabel lnkAbout;
        private Panel leftHeader, rightHeader, columnHeader;
        private System.Windows.Forms.Timer tableSearchTimer, columnSearchTimer;
        private AnalysisResult liveAnalysis;      // provisional findings while a scan is running
        private DateTime lastLiveRefresh = DateTime.MinValue;

        // ------------------------------------------------------------------ layout

        private void BuildFinderTab(TabPage page)
        {
            split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical, // vertical splitter bar: left pane | right pane
                BorderStyle = BorderStyle.None,
                BackColor = Color.White,
                Size = new Size(1096, 600)          // design size, so the minimum sizes below can be applied
            };
            split.Panel1.BackColor = Color.White;
            split.Panel2.BackColor = Color.White;
            try
            {
                split.Panel1MinSize = UiLayout.Panel1MinSize;
                split.Panel2MinSize = UiLayout.Panel2MinSize;
                split.SplitterDistance = UiLayout.InitialSplitterDistance;
            }
            catch (InvalidOperationException) { /* applied again on the first real resize */ }
            split.SizeChanged += (s, e) => ApplyLeftWidth();

            BuildLeftPanel(split.Panel1);
            BuildRightPanel(split.Panel2);
            page.Controls.Add(split);
        }

        /// <summary>Left pane = 27 % of the split width, clamped to 280-350 px, never starving the right pane (only when the split is at least 700 px wide).</summary>
        private void ApplyLeftWidth()
        {
            int? width = UiLayout.LeftPaneWidth(split.Width, split.SplitterWidth);
            if (!width.HasValue) return;
            try { split.SplitterDistance = Math.Max(split.Panel1MinSize, width.Value); }
            catch (InvalidOperationException) { }
        }

        private void BuildLeftPanel(Control host)
        {
            // rows: header block (176) | table cards | column header (56) | column cards
            var layout = NewStack();

            leftHeader = new Panel { Dock = DockStyle.Fill, Padding = new Padding(8), BackColor = Color.White, Margin = new Padding(0) };
            const AnchorStyles stretch = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;

            lblEnv = new WinFormsLabel { Left = 8, Top = 4, Height = 20, AutoSize = false, AutoEllipsis = true, ForeColor = MutedColor, Anchor = stretch, TextAlign = ContentAlignment.MiddleLeft };
            btnLoad = Btn("Load Tables", 230, (s, e) => LoadTables(), 32);
            btnLoad.Left = 8; btnLoad.Top = 27; btnLoad.Anchor = stretch;
            lblSearchTables = new WinFormsLabel { Left = 8, Top = 65, AutoSize = true, ForeColor = TextColor, Text = "Search tables:" };
            txtTableSearch = new TextBox { Left = 8, Top = 84, Height = 24, Width = 230, Anchor = stretch };
            cmbTableKind = new ComboBox { Left = 8, Top = 117, Width = 110, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbTableKind.Items.AddRange(new object[] { "All tables", "Custom tables", "Standard tables" });
            cmbTableKind.SelectedIndex = 0;
            btnClearSelection = Btn("Clear Selection", 110, (s, e) => ClearTableSelection(), 29);
            btnClearSelection.Left = 122; btnClearSelection.Top = 114;
            lblTableSummary = new WinFormsLabel { Left = 8, Top = 148, Height = 19, AutoSize = false, ForeColor = TextColor, Anchor = stretch, TextAlign = ContentAlignment.MiddleLeft };
            leftHeader.Controls.AddRange(new Control[] { lblEnv, btnLoad, lblSearchTables, txtTableSearch, cmbTableKind, btnClearSelection, lblTableSummary });
            leftHeader.SizeChanged += (s, e) => LayoutLeftHeader();

            tableSearchTimer = new System.Windows.Forms.Timer { Interval = 200 };
            tableSearchTimer.Tick += (s, e) => { tableSearchTimer.Stop(); BindTables(); };
            txtTableSearch.TextChanged += (s, e) => { tableSearchTimer.Stop(); tableSearchTimer.Start(); };
            cmbTableKind.SelectedIndexChanged += (s, e) => BindTables();

            cardsTables = new CardList();
            cardsTables.SelectionChanged += (s, e) => OnTableSelected();

            columnHeader = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Margin = new Padding(0) };
            lblColumnHeader = new WinFormsLabel { Left = 8, Top = 4, Height = 18, AutoSize = false, AutoEllipsis = true, ForeColor = TextColor, Anchor = stretch };
            txtColumnSearch = new TextBox { Left = 8, Top = 26, Height = 24, Width = 230, Anchor = stretch };
            columnHeader.Controls.AddRange(new Control[] { lblColumnHeader, txtColumnSearch });
            columnHeader.SizeChanged += (s, e) =>
            {
                int inner = UiLayout.HeaderInnerWidth(columnHeader.ClientSize.Width);
                lblColumnHeader.Width = inner; txtColumnSearch.Width = inner;
            };
            columnSearchTimer = new System.Windows.Forms.Timer { Interval = 200 };
            columnSearchTimer.Tick += (s, e) => { columnSearchTimer.Stop(); BindColumns(); };
            txtColumnSearch.TextChanged += (s, e) => { columnSearchTimer.Stop(); columnSearchTimer.Start(); };

            cardsColumns = new CardList();
            cardsColumns.SelectionChanged += (s, e) => OnColumnSelected();

            AddFixedRow(layout, leftHeader, UiLayout.LeftHeaderHeight);
            AddPercentRow(layout, cardsTables, 55);
            AddFixedRow(layout, columnHeader, 56);
            AddPercentRow(layout, cardsColumns, 45);
            host.Controls.Add(layout);
            LayoutLeftHeader();
        }

        /// <summary>Controls of the left header keep the inner width = max(230, panel width - 24); the two buttons share a row.</summary>
        private void LayoutLeftHeader()
        {
            if (leftHeader == null) return;
            int inner = UiLayout.HeaderInnerWidth(leftHeader.ClientSize.Width);
            int half = UiLayout.HalfButtonWidth(inner);
            lblEnv.Width = inner;
            btnLoad.Width = inner;
            txtTableSearch.Width = inner;
            lblTableSummary.Width = inner;
            cmbTableKind.Width = half;
            btnClearSelection.Left = 8 + half + 4;
            btnClearSelection.Width = Math.Max(60, inner - half - 4);
        }

        private void BuildRightPanel(Control host)
        {
            // rows: header (78) | progress bar (8) | configuration (auto) | status grid (fill) | maintainer link (auto)
            var layout = NewStack();

            rightHeader = new Panel { Dock = DockStyle.Fill, Padding = new Padding(8, 4, 8, 4), BackColor = Color.White, Margin = new Padding(0) };
            var commandBar = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 36, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, BackColor = Color.White };
            btnFind = Btn("Find Dependencies", 150, (s, e) => FindDependencies());
            btnClear = Btn("Clear Results", 108, (s, e) => ClearResults());
            btnViewReport = Btn("View Impact Report", 140, (s, e) => tabs.SelectedTab = tabReport);
            btnCancel = Btn("Cancel Scan", 90, (s, e) => { if (cts != null) cts.Cancel(); });
            commandBar.Controls.AddRange(new Control[] { btnFind, btnClear, btnViewReport, btnCancel });

            lblProgress = new WinFormsLabel
            {
                Left = 8, Top = 40, Height = 30, Width = 400, AutoSize = false, ForeColor = MutedColor, TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(0, 2, 0, 0), Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right
            };
            rightHeader.Controls.Add(lblProgress);
            rightHeader.Controls.Add(commandBar);
            rightHeader.SizeChanged += (s, e) => lblProgress.Width = UiLayout.StatusLabelWidth(rightHeader.ClientSize.Width);

            bar = new ProgressBar { Dock = DockStyle.Fill, Minimum = 0, Maximum = 100, Margin = new Padding(8, 0, 8, 0) };

            // configuration
            var cfg = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 1, Padding = new Padding(8, 6, 8, 0), BackColor = Color.White, Margin = new Padding(0) };
            cfg.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            lblSelection = new WinFormsLabel { ForeColor = TextColor };
            WrapLabel(lblSelection, cfg);

            var modeBox = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Margin = new Padding(0, 4, 0, 0) };
            rbTable = new RadioButton { Text = "Analyze the selected table and its associated components", AutoSize = true, Checked = true };
            rbColumn = new RadioButton { Text = "Analyze the selected column and components that reference it", AutoSize = true };
            rbTable.CheckedChanged += (s, e) => OnModeChanged();
            modeBox.Controls.AddRange(new Control[] { rbTable, rbColumn });

            lblValidation = new WinFormsLabel { Margin = new Padding(0, 4, 0, 4), ForeColor = MutedColor };
            WrapLabel(lblValidation, cfg);

            var scopeLabel = new WinFormsLabel { Text = "Component scopes:", AutoSize = true, Margin = new Padding(0, 6, 0, 2) };
            var scopeFlow = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true, Margin = new Padding(0) };
            scopeChecks = ScanRequest.AllScopes.Select(k => new CheckBox { Text = k.Display(), AutoSize = true, Checked = true, Tag = k, Margin = new Padding(0, 0, 14, 2) }).ToArray();
            foreach (CheckBox c in scopeChecks) c.CheckedChanged += (s, e) => UpdateState();
            scopeFlow.Controls.AddRange(scopeChecks);

            var scopeButtons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true, Margin = new Padding(0, 2, 0, 0) };
            btnScopeAll = Btn("Select All", 85, (s, e) => SetAllScopes(true), 29);
            btnScopeNone = Btn("Clear All", 85, (s, e) => SetAllScopes(false), 29);
            scopeButtons.Controls.AddRange(new Control[] { btnScopeAll, btnScopeNone });

            var optionsFlow = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true, Margin = new Padding(0, 6, 0, 0) };
            var jsLabel = new WinFormsLabel { Text = "JavaScript scope:", AutoSize = true, Margin = new Padding(0, 6, 4, 0) };
            cmbJsScope = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 330 };
            cmbJsScope.Items.AddRange(new object[] { "All script web resources (most complete)", "Only libraries registered on the selected table's forms (faster)" });
            cmbJsScope.SelectedIndex = 0;
            chkComments = new CheckBox { Text = "Include comment-only JavaScript matches", AutoSize = true, Margin = new Padding(14, 5, 0, 0) };
            chkSkipManaged = new CheckBox { Text = "Skip managed web resources (faster)", AutoSize = true, Margin = new Padding(0, 5, 0, 0) };
            cmbJsScope.SelectedIndexChanged += (s2, e2) => UpdateState();
            chkSkipManaged.CheckedChanged += (s2, e2) => UpdateState();
            chkDeep = new CheckBox { Text = "Cross-table process scan (downloads every process definition)", AutoSize = true, Margin = new Padding(0, 5, 0, 0) };
            var tip = new ToolTip();
            tip.SetToolTip(cmbJsScope, "All script web resources downloads every JavaScript library in the environment. The targeted option only reads libraries registered on the selected table's forms.");
            tip.SetToolTip(chkComments, "Matches that appear only inside JavaScript comments are not runtime references; they are counted but not listed unless this is ticked.");
            tip.SetToolTip(chkDeep, "By default only processes whose primary table is the selected table are read. Tick this to also find processes of other tables that update or read the selected table or column.");
            tip.SetToolTip(chkSkipManaged, "Managed web resources come from installed solutions. Skipping them shortens the download a lot, but references inside them are not searched.");
            optionsFlow.Controls.AddRange(new Control[] { jsLabel, cmbJsScope, chkComments });

            lblSummary = new WinFormsLabel { Margin = new Padding(0, 8, 0, 4), ForeColor = TextColor };
            WrapLabel(lblSummary, cfg);

            AddRow(cfg, lblSelection);
            AddRow(cfg, modeBox);
            AddRow(cfg, lblValidation);
            AddRow(cfg, scopeLabel);
            AddRow(cfg, scopeFlow);
            AddRow(cfg, scopeButtons);
            AddRow(cfg, optionsFlow);
            AddRow(cfg, chkSkipManaged);
            AddRow(cfg, chkDeep);
            AddRow(cfg, lblSummary);

            // per-category status grid
            gridStatus = NewGrid(true);
            gridStatus.DataSource = statusSource;
            gridStatus.DataBindingComplete += (s2, e2) => ClearGridSelection(gridStatus); // no row is highlighted until the user picks one
            statusSource.DataSource = new BindingList<StatusRow>();
            AddStatusColumn("Component category", "Category", 1.6, 190);
            AddStatusColumn("Scan status", "Status", 1.1, 130);
            AddStatusColumn("Examined", "Examined", 0.6, 80);
            AddStatusColumn("Matches", "Matches", 0.5, 70);
            AddStatusColumn("Notes", "Notes", 3.0, 200);
            gridStatus.CellToolTipTextNeeded += (s, e) =>
            {
                var r = RowItem<StatusRow>(gridStatus, e.RowIndex);
                if (r == null) return;
                ScanDiagnostic d = analysis == null ? null : analysis.Diagnostics.FirstOrDefault(x => x.Category.Display() == r.Category);
                if (d == null) return;
                var lines = new List<string>(d.Errors); lines.AddRange(d.Warnings); lines.AddRange(d.Limitations);
                e.ToolTipText = DataverseHelper.Truncate(string.Join("\n", lines), 1500);
            };
            var gridHost = new Panel { Dock = DockStyle.Fill, Padding = new Padding(8, 4, 8, 8), BackColor = Color.White, Margin = new Padding(0) };
            gridHost.Controls.Add(gridStatus);

            lnkAbout = new LinkLabel
            {
                AutoSize = true, Margin = new Padding(8, 0, 8, 6), LinkColor = Color.FromArgb(48, 105, 190),
                Text = ProductInfo.Name + " v" + ProductInfo.Version + "  \u00B7  Maintainer: " + ProductInfo.Maintainer + " (LinkedIn)"
            };
            int linkStart = lnkAbout.Text.IndexOf("Maintainer:");
            lnkAbout.Links.Add(linkStart, lnkAbout.Text.Length - linkStart, ProductInfo.LinkedInUrl);
            lnkAbout.LinkClicked += (s2, e2) =>
            {
                try { System.Diagnostics.Process.Start((string)e2.Link.LinkData); }
                catch (Exception ex) { LogFailure("Could not open the link: " + ex.Message); }
            };
            tip.SetToolTip(lnkAbout, "Host: " + ProductInfo.Host + "  |  Data platform: " + ProductInfo.DataPlatform + "  |  Export: " + ProductInfo.ExportFormat + "  |  " + ProductInfo.RepositoryUrl);

            AddFixedRow(layout, rightHeader, UiLayout.RightHeaderHeight);
            AddFixedRow(layout, bar, 8);
            AddRow(layout, cfg);
            AddRow(layout, gridHost, true);
            AddRow(layout, lnkAbout);
            host.Controls.Add(layout);
        }

        private void AddStatusColumn(string header, string property, double weight, int minWidth)
        {
            gridStatus.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = header, DataPropertyName = property, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = (float)(weight * 100), MinimumWidth = minWidth, SortMode = DataGridViewColumnSortMode.Automatic, ReadOnly = true
            });
        }

        private static T RowItem<T>(DataGridView g, int rowIndex) where T : class
        {
            return rowIndex < 0 || rowIndex >= g.Rows.Count ? null : g.Rows[rowIndex].DataBoundItem as T;
        }

        // ------------------------------------------------------------------ messages

        /// <summary>Messages about loading metadata share the single status line of the right header.</summary>
        private void SetLeft(string message, bool error = false)
        {
            SetProgressText(message, error);
        }

        private void SetProgressText(string message, bool error = false)
        {
            lblProgress.Text = message;
            lblProgress.ForeColor = error ? ErrorColor : MutedColor;
        }

        /// <summary>Idle instruction that depends on the current state (nothing loaded / tables loaded).</summary>
        private void RefreshIdleMessages()
        {
            if (analysis != null || liveAnalysis != null) return;
            if (loadedEnvironment == null)
                SetProgressText("Load tables from the left pane, select a table, then click Find Dependencies.");
            else
                SetProgressText("Choose a table" + (rbColumn.Checked ? " and a column" : string.Empty) + ", pick the component scopes, then click 'Find Dependencies'.");
        }

        // ------------------------------------------------------------------ state

        private int CheckedScopeCount { get { return scopeChecks == null ? 0 : scopeChecks.Count(c => c.Checked); } }

        private string CurrentProblem()
        {
            return ScanRequest.Validate(rbColumn.Checked ? SearchMode.Column : SearchMode.Table,
                selectedTable == null ? null : selectedTable.LogicalName,
                selectedColumn == null ? null : selectedColumn.LogicalName,
                CheckedScopeCount);
        }

        private void UpdateState()
        {
            if (btnFind == null || btnCancel == null) return;

            bool tablesLoaded = allTables.Count > 0;
            bool columnMode = rbColumn.Checked;

            // selection summary
            string table = selectedTable == null ? "(none)" : selectedTable.Title;
            string column = !columnMode ? "(not used in table mode)" : selectedColumn == null ? "(none)" : selectedColumn.Title;
            lblSelection.Text = "Table: " + table + "\nColumn: " + column;

            string problem = CurrentProblem();
            if (!tablesLoaded) { lblValidation.Text = string.Empty; }
            else if (problem != null) { lblValidation.Text = problem; lblValidation.ForeColor = WarnColor; }
            else
            {
                lblValidation.Text = columnMode
                    ? "Ready: analyze column '" + selectedColumn.LogicalName + "' of table '" + selectedTable.LogicalName + "'."
                    : "Ready: analyze table '" + selectedTable.LogicalName + "' (a selected column is ignored in table mode).";
                lblValidation.ForeColor = MutedColor;
            }

            btnLoad.Text = allTables.Count > 0 ? "Refresh Metadata" : "Load Tables";
            btnLoad.Enabled = !busy;

            txtTableSearch.Enabled = cmbTableKind.Enabled = cardsTables.Enabled = btnClearSelection.Enabled = !busy && tablesLoaded;
            bool columnsEnabled = !busy && tablesLoaded && selectedTable != null && columnMode;
            txtColumnSearch.Enabled = cardsColumns.Enabled = columnsEnabled;
            lblColumnHeader.Text = selectedTable == null
                ? "Columns (select a table first)"
                : columnMode ? "Columns of " + selectedTable.LogicalName : "Columns of " + selectedTable.LogicalName + " (used in column mode only)";
            lblTableSummary.Text = (selectedTable == null ? 0 : 1) + " of " + allTables.Count + " tables selected.";

            rbTable.Enabled = rbColumn.Enabled = !busy;
            foreach (CheckBox c in scopeChecks) c.Enabled = !busy;
            btnScopeAll.Enabled = btnScopeNone.Enabled = !busy;
            cmbJsScope.Enabled = chkComments.Enabled = chkDeep.Enabled = !busy;
            bool jsSelected = scopeChecks.Any(c => (ScannerKind)c.Tag == ScannerKind.JavaScript && c.Checked);
            bool allJs = cmbJsScope.SelectedIndex == 0;
            chkSkipManaged.Enabled = !busy && jsSelected && allJs;

            btnFind.Enabled = !busy && tablesLoaded && Service != null && problem == null;
            btnClear.Enabled = !busy && analysis != null;
            btnViewReport.Enabled = analysis != null;
            if (!busy) btnCancel.Enabled = false;

            UpdateReportControls();
        }

        private void SetBusy(bool isBusy, bool allowCancel = true)
        {
            busy = isBusy;
            btnCancel.Enabled = isBusy && allowCancel;
            if (!isBusy)
            {
                bar.Style = ProgressBarStyle.Continuous;
                bar.Value = 0;
            }
            UpdateState();
            UseWaitCursor = false;
        }

        private void SetAllScopes(bool on)
        {
            foreach (CheckBox c in scopeChecks) c.Checked = on;
        }

        private void OnModeChanged()
        {
            if (rbTable == null) return;
            // Switching mode changes what a scan means; the loaded results keep the request they were produced with.
            RefreshIdleMessages();
            UpdateState();
        }

        // ------------------------------------------------------------------ metadata: tables

        private async void LoadTables()
        {
            if (busy || cts != null) return;

            if (Service == null)
            {
                SetLeft("Not connected. Connect to an environment with XrmToolBox first.", true);
                MessageBox.Show(this, "Connect to a Dataverse environment with XrmToolBox first, then load the tables.", "Not connected", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            bool refresh = allTables.Count > 0;
            cts = new CancellationTokenSource();
            CancellationToken token = cts.Token;
            Microsoft.Xrm.Sdk.IOrganizationService svc = Service;
            string env = EnvironmentName();
            var timer = System.Diagnostics.Stopwatch.StartNew();

            SetBusy(true);
            bar.Style = ProgressBarStyle.Marquee;
            SetLeft(refresh ? "Refreshing metadata..." : "Loading tables...");

            try
            {
                List<TableInfo> tables = await Task.Run(() =>
                {
                    if (refresh) metadata.Invalidate();
                    return metadata.LoadTables(svc, refresh, token);
                }, token);

                if (!string.Equals(loadedEnvironment, env, StringComparison.OrdinalIgnoreCase) || refresh)
                {
                    // A new table list starts a new session: earlier selections and results no longer apply.
                    ClearResults();
                    selectedTable = null; selectedColumn = null; columns = new List<ColumnInfo>();
                }
                allTables = tables;
                loadedEnvironment = env;

                BindTables();
                BindColumns();

                SetLeft(allTables.Count == 0
                    ? "No eligible tables were found in this environment."
                    : "Loaded " + allTables.Count + " tables (" + allTables.Count(t => t.IsCustom) + " custom) in " +
                      (timer.ElapsedMilliseconds / 1000.0).ToString("0.0") + "s. Select a table.");
                RefreshIdleMessages();
            }
            catch (OperationCanceledException)
            {
                SetLeft("Cancelled. The previous table list was not changed.");
            }
            catch (Exception ex)
            {
                SetLeft("Loading tables failed: " + ex.Message, true);
                LogFailure("Loading tables failed: " + ex);
                MessageBox.Show(this,
                    "Loading the table list failed:\n\n" + ex.Message + "\n\nCheck your connection and that your user can read table metadata.",
                    "Load failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                SetBusy(false);
                cts.Dispose();
                cts = null;
            }
        }

        private void BindTables()
        {
            if (cardsTables == null) return;
            TableKindFilter kind = cmbTableKind.SelectedIndex == 1 ? TableKindFilter.Custom : cmbTableKind.SelectedIndex == 2 ? TableKindFilter.Standard : TableKindFilter.All;
            List<TableInfo> shown = EntityMetadataService.FilterTables(allTables, txtTableSearch.Text, kind);

            binding = true;
            try
            {
                cardsTables.SetItems(shown.Select(t => new CardItem
                {
                    Key = t.LogicalName,
                    Title = t.Title,
                    Subtitle = (t.IsCustom ? "Custom" : "Standard") + " \u00B7 " + (string.IsNullOrEmpty(t.EntitySetName) ? "table" : t.EntitySetName)
                }).ToList(), selectedTable == null ? null : selectedTable.LogicalName);
            }
            finally { binding = false; }
            UpdateState();
        }

        private void ClearTableSelection()
        {
            // The table selection (and with it the column selection) is cleared; the list itself stays loaded.
            cardsTables.ClearSelection();
            if (selectedTable != null) { selectedTable = null; OnTableSelected(); }
        }

        private void OnTableSelected()
        {
            if (binding) return;
            string key = cardsTables.SelectedKey;

            if (key == null)
            {
                columnTicket++;
                if (columnCts != null) columnCts.Cancel();
                selectedTable = null;
                selectedColumn = null;
                columns = new List<ColumnInfo>();
                txtColumnSearch.Text = string.Empty;
                BindColumns();
                RefreshIdleMessages();
                UpdateState();
                return;
            }

            TableInfo t = allTables.FirstOrDefault(x => x.LogicalName == key);
            if (t == null || (selectedTable != null && selectedTable.LogicalName == t.LogicalName)) return;

            selectedTable = t;
            selectedColumn = null;
            columns = new List<ColumnInfo>();
            txtColumnSearch.Text = string.Empty;
            BindColumns();
            RefreshIdleMessages();
            UpdateState();
            LoadColumnsFor(t);
        }

        // ------------------------------------------------------------------ metadata: columns

        private async void LoadColumnsFor(TableInfo table)
        {
            int ticket = ++columnTicket;
            if (columnCts != null) columnCts.Cancel();
            columnCts = new CancellationTokenSource();
            CancellationToken token = columnCts.Token;
            Microsoft.Xrm.Sdk.IOrganizationService svc = Service;
            if (svc == null) return;

            SetLeft("Loading columns of " + table.LogicalName + "...");
            try
            {
                List<ColumnInfo> list = await Task.Run(() => metadata.LoadColumns(svc, table.LogicalName, false, token), token);
                if (ticket != columnTicket) return; // a newer selection superseded this one
                columns = list;
                BindColumns();
                SetLeft(list.Count + " columns of " + table.LogicalName + " loaded.");
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                if (ticket != columnTicket) return;
                columns = new List<ColumnInfo>();
                BindColumns();
                SetLeft("Columns of " + table.LogicalName + " could not be loaded: " + ex.Message, true);
                LogFailure("Loading columns failed: " + ex);
            }
            UpdateState();
        }

        private void BindColumns()
        {
            if (cardsColumns == null) return;
            List<ColumnInfo> shown = EntityMetadataService.FilterColumns(columns, txtColumnSearch.Text);
            binding = true;
            try
            {
                cardsColumns.SetItems(shown.Select(c => new CardItem
                {
                    Key = c.LogicalName,
                    Title = c.Title,
                    Subtitle = (string.IsNullOrEmpty(c.AttributeType) ? "Column" : c.AttributeType) + " \u00B7 " + (c.IsCustom ? "custom" : "standard")
                }).ToList(), selectedColumn == null ? null : selectedColumn.LogicalName);
            }
            finally { binding = false; }
        }

        private void OnColumnSelected()
        {
            if (binding) return;
            string key = cardsColumns.SelectedKey;
            selectedColumn = key == null ? null : columns.FirstOrDefault(c => c.LogicalName == key);
            RefreshIdleMessages();
            UpdateState();
        }

        // ------------------------------------------------------------------ scan

        private ScanRequest BuildRequest()
        {
            bool column = rbColumn.Checked;
            return new ScanRequest
            {
                Mode = column ? SearchMode.Column : SearchMode.Table,
                TableLogicalName = selectedTable == null ? null : selectedTable.LogicalName,
                TableDisplayName = selectedTable == null ? null : selectedTable.DisplayName,
                TableEntitySetName = selectedTable == null ? null : selectedTable.EntitySetName,
                ColumnLogicalName = column && selectedColumn != null ? selectedColumn.LogicalName : null,
                ColumnDisplayName = column && selectedColumn != null ? selectedColumn.DisplayName : null,
                Scopes = new HashSet<ScannerKind>(scopeChecks.Where(c => c.Checked).Select(c => (ScannerKind)c.Tag)),
                JavaScriptScope = cmbJsScope.SelectedIndex == 1 ? JavaScriptScope.RegisteredOnSelectedTableForms : JavaScriptScope.AllScriptWebResources,
                SkipManagedWebResources = chkSkipManaged.Checked,
                IncludeCommentMatches = chkComments.Checked,
                DeepProcessScan = chkDeep.Checked,
                EnvironmentName = EnvironmentName(),
                EnvironmentUrl = EnvironmentUrl(),
                EnvironmentId = EnvironmentId()
            };
        }

        private async void FindDependencies()
        {
            if (busy || cts != null) return;

            ScanRequest request = BuildRequest();
            string problem = request.Validate();
            if (problem != null)
            {
                SetProgressText(problem, true);
                return;
            }
            if (Service == null)
            {
                SetProgressText("Not connected. Connect to an environment with XrmToolBox first.", true);
                return;
            }

            // Searching every web resource can take a long time: ask before starting, and say how to make it faster.
            if (request.Scopes.Contains(ScannerKind.JavaScript) && request.JavaScriptScope == JavaScriptScope.AllScriptWebResources)
            {
                DialogResult answer = MessageBox.Show(this,
                    ScanWarnings.AllWebResources(request.SkipManagedWebResources),
                    "Search all JavaScript web resources?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
                if (answer != DialogResult.Yes)
                {
                    SetProgressText("Scan not started. Change the JavaScript scope or the options, then click 'Find Dependencies' again.");
                    return;
                }
            }

            cts = new CancellationTokenSource();
            CancellationToken token = cts.Token;
            Microsoft.Xrm.Sdk.IOrganizationService svc = Service;

            // A new scan replaces the previous result; it is never merged with it.
            ClearResults();
            liveAnalysis = new AnalysisResult { Request = request };
            lastLiveRefresh = DateTime.MinValue;
            bool finished = false;
            SetBusy(true);
            SetProgressText("Starting scan of " + (request.Mode == SearchMode.Column ? request.TableLogicalName + "." + request.ColumnLogicalName : request.TableLogicalName) + "...");

            var progress = new Progress<ProgressInfo>(p =>
            {
                if (finished) return; // a late callback must not overwrite the final message
                bar.Value = Math.Max(0, Math.Min(100, p.Percent));
                if (!string.IsNullOrEmpty(p.Message)) SetProgressText(p.Message);

                if (p.NewResults != null && p.NewResults.Count > 0 && liveAnalysis != null)
                {
                    liveAnalysis.Results.AddRange(p.NewResults);
                    lblSummary.ForeColor = TextColor;
                    lblSummary.Text = "Live: " + liveAnalysis.MatchCount + " match(es) found so far (provisional)";
                    if ((DateTime.UtcNow - lastLiveRefresh).TotalMilliseconds > 700)
                    {
                        lastLiveRefresh = DateTime.UtcNow;
                        RefreshReport(true);
                    }
                }
            });

            try
            {
                AnalysisResult result = await Task.Run(() => analysisService.Analyze(svc, request, progress, token));
                finished = true;
                analysis = result;
                liveAnalysis = null;

                BindStatus();
                lblSummary.Text = AnalysisSummary.Line(analysis);
                lblSummary.ForeColor = analysis.FailedScans > 0 ? ErrorColor : TextColor;
                SetProgressText(FinishMessage(analysis), analysis.FailedScans > 0);
                RefreshReport(true);
            }
            catch (Exception ex)
            {
                finished = true;
                liveAnalysis = null;
                RefreshReport(true);
                SetProgressText("The scan failed: " + ex.Message, true);
                LogFailure("Scan failed: " + ex);
                MessageBox.Show(this, "The scan failed:\n\n" + ex.Message, "Scan failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                SetBusy(false);
                cts.Dispose();
                cts = null;
            }
        }

        private static string FinishMessage(AnalysisResult a)
        {
            if (a.WasCancelled) return "Scan cancelled. Partial results are shown; categories that did not finish are marked in the table below.";
            if (a.MatchCount > 0) return "Scan finished: " + a.MatchCount + " finding(s). Open the 'Impact Report' tab to filter and export them." +
                                         (a.UnsupportedOrPartialScans + a.FailedScans > 0 ? " Some categories were not fully inspected - see the status table." : string.Empty);
            return a.IsCleanEmptyResult
                ? "Scan finished: no references were found, and every selected category was inspected."
                : "Scan finished with NO matches, but not every selected category was fully inspected. This is not a clean result - see the status table.";
        }

        private void BindStatus()
        {
            if (analysis == null) { statusSource.DataSource = new BindingList<StatusRow>(); return; }
            var rows = analysis.Diagnostics.Select(d => new StatusRow
            {
                Category = d.Category.Display(),
                Status = d.Status.Display(),
                StatusCode = d.Status,
                Examined = d.WasInspected ? d.RecordsExamined.ToString() : "-",
                Matches = d.WasInspected ? d.MatchCount.ToString() : "-",
                Notes = DataverseHelper.OneLine(d.Message.Length > 0 ? d.Message : (d.Limitations.Count > 0 ? d.Limitations[0] : string.Empty))
            }).ToList();
            statusSource.DataSource = new BindingList<StatusRow>(rows);
            ClearGridSelection(gridStatus);
        }

        private void ClearResults()
        {
            analysis = null;
            liveAnalysis = null;
            if (gridStatus != null) statusSource.DataSource = new BindingList<StatusRow>();
            if (lblSummary != null) lblSummary.Text = string.Empty;
            if (lblProgress != null) RefreshIdleMessages();
            RefreshReport(true);
            UpdateState();
        }

        /// <summary>Drops everything that belongs to the previous connection.</summary>
        private void ResetForNewEnvironment()
        {
            ClearResults();
            metadata.Invalidate();
            allTables = new List<TableInfo>();
            columns = new List<ColumnInfo>();
            selectedTable = null;
            selectedColumn = null;
            loadedEnvironment = null;
            binding = true;
            try { cardsTables.SetItems(new List<CardItem>(), null); cardsColumns.SetItems(new List<CardItem>(), null); }
            finally { binding = false; }
            RefreshIdleMessages();
        }
    }
}
