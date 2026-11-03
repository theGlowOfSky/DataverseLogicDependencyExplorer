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
using System.Windows.Forms;
using McTools.Xrm.Connection;
using Microsoft.Xrm.Sdk;
using XrmToolBox.Extensibility;
using WinFormsLabel = System.Windows.Forms.Label;
using DataverseLogicDependencyExplorer.Models;
using DataverseLogicDependencyExplorer.Services;

namespace DataverseLogicDependencyExplorer
{
    /// <summary>
    /// Main plugin UI with two tabs: "Dependency Finder" (select what to analyse, run the scan) and "Impact Report" (filter, inspect
    /// and export the findings). The look and the interaction patterns follow the JavaScript Event Registration Auditor: built in
    /// code, Segoe UI 9, white background, 25:75 split, two-line list rows, system-style buttons, a thin progress bar and wrapped
    /// status labels. Long operations run with async/await + Task.Run and report through Progress&lt;T&gt;.
    /// This file holds the shared state, helpers and the XrmToolBox connection hooks.
    /// </summary>
    public partial class MainControl : PluginControlBase
    {
        // Palette (exact RGB values of the inventory-style spec). The palette has no alert colours: meaning is carried by wording.
        private static readonly Color TextColor = Color.Black;
        private static readonly Color MutedColor = Color.DimGray;
        private static readonly Color ErrorColor = Color.Black;
        private static readonly Color WarnColor = Color.Black;
        private static readonly Color ButtonBack = Color.FromArgb(242, 242, 242);
        private static readonly Color ButtonBorder = Color.FromArgb(205, 205, 205);
        private static readonly Color HeaderBeige = Color.FromArgb(241, 238, 224);
        private static readonly Color GridLine = Color.FromArgb(220, 220, 220);
        private static readonly Color SelectionBlue = Color.FromArgb(48, 105, 190);

        private readonly EntityMetadataService metadata = new EntityMetadataService();
        private readonly DependencyAnalysisService analysisService = new DependencyAnalysisService();

        private List<TableInfo> allTables = new List<TableInfo>();
        private List<ColumnInfo> columns = new List<ColumnInfo>();
        private TableInfo selectedTable;
        private ColumnInfo selectedColumn;
        private AnalysisResult analysis;

        private CancellationTokenSource cts;        // the running scan / table load / export
        private CancellationTokenSource columnCts;  // the running column load
        private int columnTicket;
        private bool busy;
        private string loadedEnvironment;

        private TabControl tabs;
        private TabPage tabFinder, tabReport;

        public MainControl()
        {
            Size = new Size(1100, 650);
            BuildUi();
            UpdateEnvLabel();
            RefreshIdleMessages();
            UpdateState();
        }

        // ------------------------------------------------------------------ UI shell

        private void BuildUi()
        {
            Dock = DockStyle.Fill;
            Font = new Font("Segoe UI", 8.5F);
            BackColor = Color.White;
            ForeColor = TextColor;

            tabs = new TabControl { Dock = DockStyle.Fill, Padding = new Point(14, 4) };
            tabFinder = new TabPage("Dependency Finder") { BackColor = Color.White, UseVisualStyleBackColor = false };
            tabReport = new TabPage("Impact Report") { BackColor = Color.White, UseVisualStyleBackColor = false };

            BuildFinderTab(tabFinder);
            BuildReportTab(tabReport);

            tabs.TabPages.Add(tabFinder);
            tabs.TabPages.Add(tabReport);
            Controls.Add(tabs);
        }

        private static TableLayoutPanel NewStack()
        {
            var t = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                Padding = new Padding(0),
                Margin = new Padding(0),
                BackColor = Color.White
            };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            return t;
        }

        private static void AddRow(TableLayoutPanel t, Control c, bool fill = false)
        {
            t.RowStyles.Add(fill ? new RowStyle(SizeType.Percent, 100) : new RowStyle(SizeType.AutoSize));
            t.Controls.Add(c, 0, t.RowStyles.Count - 1);
        }

        /// <summary>A row of an exact pixel height.</summary>
        private static void AddFixedRow(TableLayoutPanel t, Control c, int height)
        {
            t.RowStyles.Add(new RowStyle(SizeType.Absolute, height));
            t.Controls.Add(c, 0, t.RowStyles.Count - 1);
        }

        /// <summary>A row that takes a share of the remaining height.</summary>
        private static void AddPercentRow(TableLayoutPanel t, Control c, float percent)
        {
            t.RowStyles.Add(new RowStyle(SizeType.Percent, percent));
            t.Controls.Add(c, 0, t.RowStyles.Count - 1);
        }

        /// <summary>Makes a label wrap to the width of its panel so long messages never overlap other controls.</summary>
        private static void WrapLabel(WinFormsLabel label, Control host)
        {
            label.AutoSize = true;
            label.Dock = DockStyle.Fill;
            EventHandler fit = (s, e) =>
                label.MaximumSize = new Size(Math.Max(60, host.ClientSize.Width - host.Padding.Horizontal - 4), 0);
            host.SizeChanged += fit;
            fit(null, null);
        }

        /// <summary>Standard button style: flat, (242,242,242), 1 px (205,205,205) border, Segoe UI 8.5, height 30, right margin 7.</summary>
        private static void StyleButton(Button b, int width, int height = 30)
        {
            b.FlatStyle = FlatStyle.Flat;
            b.BackColor = ButtonBack;
            b.ForeColor = Color.Black;
            b.UseVisualStyleBackColor = false;
            b.FlatAppearance.BorderSize = 1;
            b.FlatAppearance.BorderColor = ButtonBorder;
            b.Font = new Font("Segoe UI", 8.5F);
            b.AutoSize = false;
            b.Width = width;
            b.Height = height;
            b.Margin = new Padding(0, 0, 7, 0);
            b.TextAlign = ContentAlignment.MiddleCenter;
        }

        /// <summary>Removes the automatic first-row selection after a (re)bind, so a row is highlighted only when the user picks one.</summary>
        private static void ClearGridSelection(DataGridView g)
        {
            g.ClearSelection();
            try { g.CurrentCell = null; } catch (InvalidOperationException) { }
        }

        private Button Btn(string text, int width, EventHandler click, int height = 30)
        {
            var b = new Button { Text = text };
            StyleButton(b, width, height);
            b.Click += click;
            return b;
        }

        /// <summary>
        /// Standard grid: white, thin grey lines, 28 px rows, (48,105,190) selection with white text and a beige bold header that
        /// never turns blue (header selection colours equal the normal colours, and the header cell is painted by hand).
        /// </summary>
        private DataGridView NewGrid(bool headers = true)
        {
            var g = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AutoGenerateColumns = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                RowHeadersVisible = false,
                ColumnHeadersVisible = headers,
                MultiSelect = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                GridColor = GridLine,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = 32
            };
            g.RowTemplate.Height = 28;
            g.DefaultCellStyle.Font = new Font("Segoe UI", 8.5F);
            g.DefaultCellStyle.BackColor = Color.White;
            g.DefaultCellStyle.ForeColor = Color.Black;
            g.DefaultCellStyle.SelectionBackColor = SelectionBlue;
            g.DefaultCellStyle.SelectionForeColor = Color.White;

            DataGridViewCellStyle h = g.ColumnHeadersDefaultCellStyle;
            h.Font = new Font("Segoe UI", 8.5F, FontStyle.Bold);
            h.Alignment = DataGridViewContentAlignment.MiddleLeft;
            h.BackColor = HeaderBeige;
            h.ForeColor = Color.Black;
            h.SelectionBackColor = h.BackColor;
            h.SelectionForeColor = h.ForeColor;

            g.CellPainting += (s, e) =>
            {
                if (e.RowIndex != -1) return;
                e.CellStyle.SelectionBackColor = HeaderBeige;
                e.CellStyle.SelectionForeColor = Color.Black;
                e.PaintBackground(e.CellBounds, false);
                e.PaintContent(e.CellBounds);
                e.Paint(e.CellBounds, DataGridViewPaintParts.Border);
                e.Handled = true;
            };
            g.DataError += (s, e) => { e.ThrowException = false; };
            return g;
        }

        private static void LogFailure(string m)
        {
            // Full exception details go to the trace output; the user sees the message in the panel and a dialog.
            System.Diagnostics.Trace.TraceError(DateTime.Now.ToString("s") + "  " + m);
        }

        // ------------------------------------------------------------------ connection

        public override void UpdateConnection(IOrganizationService newService, ConnectionDetail detail, string actionName, object parameter)
        {
            base.UpdateConnection(newService, detail, actionName, parameter);

            if (loadedEnvironment != null &&
                !string.Equals(loadedEnvironment, EnvironmentName(), StringComparison.OrdinalIgnoreCase))
            {
                // Everything belongs to the previous connection; never present it as coming from the new one.
                CancelRunning();
                ResetForNewEnvironment();
                SetLeft("Connection changed. Click Load Tables to reload.");
            }

            UpdateEnvLabel();
            UpdateState();
        }

        public override void ClosingPlugin(PluginCloseInfo info)
        {
            CancelRunning();
            base.ClosingPlugin(info);
        }

        private void CancelRunning()
        {
            if (cts != null) cts.Cancel();
            if (columnCts != null) columnCts.Cancel();
        }

        private string EnvironmentName()
        {
            var d = ConnectionDetail;
            if (d == null) return null;
            if (!string.IsNullOrWhiteSpace(d.OrganizationFriendlyName)) return d.OrganizationFriendlyName;
            if (!string.IsNullOrWhiteSpace(d.Organization)) return d.Organization;
            return d.ConnectionName;
        }

        private string EnvironmentUrl()
        {
            return ConnectionDetail == null ? null : ConnectionDetail.WebApplicationUrl;
        }

        private string EnvironmentId()
        {
            return ConnectionDetail == null ? null : ConnectionDetail.EnvironmentId;
        }

        private void UpdateEnvLabel()
        {
            string n = EnvironmentName();
            lblEnv.Text = Service == null || n == null ? "Not connected." : "Connected \u2014 " + n;
            if (Service != null && !string.IsNullOrEmpty(EnvironmentUrl())) new ToolTip().SetToolTip(lblEnv, EnvironmentUrl());
        }
    }
}
