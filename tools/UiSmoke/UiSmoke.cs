// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
// NOTE: Mono substitutes a fallback family when Segoe UI is missing, so fonts are checked via Font.OriginalFontName.
// LOCAL DEVELOPMENT SMOKE TEST (Mono WinForms under Xvfb) - NOT PART OF THE DELIVERABLE.
// Constructs the real MainControl, injects data through the same private members the UI uses, and checks that binding,
// filtering, sorting, empty states and the detail panel behave. It does not prove behaviour inside the XrmToolBox host on Windows.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Drawing;
using DataverseLogicDependencyExplorer;
using DataverseLogicDependencyExplorer.Controls;
using DataverseLogicDependencyExplorer.Models;
using DataverseLogicDependencyExplorer.Services;
using DataverseLogicDependencyExplorer.Services.Scanners;

static class Smoke
{
    static int failures;
    static void Check(bool ok, string what) { Console.WriteLine((ok ? "  ok    " : "  FAIL  ") + what); if (!ok) failures++; }
    static object Get(object o, string n) { return o.GetType().GetField(n, BindingFlags.NonPublic | BindingFlags.Instance).GetValue(o); }
    static void Set(object o, string n, object v) { o.GetType().GetField(n, BindingFlags.NonPublic | BindingFlags.Instance).SetValue(o, v); }
    static object Call(object o, string n, params object[] a) { return o.GetType().GetMethod(n, BindingFlags.NonPublic | BindingFlags.Instance).Invoke(o, a); }

    [STAThread]
    static int Main()
    {
        var plugin = new Plugin();
        var control = (MainControl)plugin.GetControl();
        var form = new Form { Width = 1250, Height = 800 };
        form.Controls.Add(control);
        form.Show();
        Application.DoEvents();

        var tabs = (TabControl)Get(control, "tabs");
        Check(tabs.TabPages.Count == 2 && tabs.TabPages[0].Text == "Dependency Finder" && tabs.TabPages[1].Text == "Impact Report", "two tabs: Dependency Finder / Impact Report");
        Check(((Control)Get(control, "btnFind")).Enabled == false, "Find Dependencies disabled before tables are loaded");
        Check(((Label)Get(control, "lblEnv")).Text == "Not connected.", "shows 'Not connected.'");

        // ---- exact geometry / style of the inventory-style spec
        var split = (SplitContainer)Get(control, "split");
        Check(split.Orientation == Orientation.Vertical && split.BorderStyle == BorderStyle.None && split.BackColor == Color.White, "vertical borderless white split container");
        Check(split.Panel1MinSize == 270 && split.Panel2MinSize == 400, "panel minimum sizes 270 / 400");
        int expectedLeft = UiLayout.LeftPaneWidth(split.Width, split.SplitterWidth) ?? split.SplitterDistance;
        Check(split.SplitterDistance == expectedLeft, "left pane width follows the 27% / 280-350 rule (" + split.SplitterDistance + " px of " + split.Width + ")");
        var leftHeader = (Panel)Get(control, "leftHeader"); var rightHeader = (Panel)Get(control, "rightHeader");
        Check(leftHeader.Parent is TableLayoutPanel && ((TableLayoutPanel)leftHeader.Parent).RowStyles[0].Height == 176 && ((TableLayoutPanel)leftHeader.Parent).RowStyles[0].SizeType == SizeType.Absolute, "left header row is an absolute 176 px");
        Check(((TableLayoutPanel)rightHeader.Parent).RowStyles[0].Height == 78 && ((TableLayoutPanel)rightHeader.Parent).RowStyles[0].SizeType == SizeType.Absolute, "right header row is an absolute 78 px");
        var btnLoadC = (Button)Get(control, "btnLoad"); var txtSearch = (TextBox)Get(control, "txtTableSearch"); var lblSum = (Label)Get(control, "lblTableSummary"); var lblEnvC = (Label)Get(control, "lblEnv");
        Check(lblEnvC.Top == 4 && lblEnvC.Height == 20 && lblEnvC.Left == 8, "connection label at (8,4) h20");
        Check(btnLoadC.Top == 27 && btnLoadC.Height == 32 && btnLoadC.Left == 8, "load button at y27 h32");
        Check(txtSearch.Top == 84 && txtSearch.Height >= 20 && lblSum.Top == 148 && lblSum.Height == 19, "search box y84, summary label y148 h19");
        Check(btnLoadC.BackColor == Color.FromArgb(242, 242, 242) && btnLoadC.FlatStyle == FlatStyle.Flat && btnLoadC.FlatAppearance.BorderColor == Color.FromArgb(205, 205, 205) && btnLoadC.FlatAppearance.BorderSize == 1 && btnLoadC.Font.OriginalFontName == "Segoe UI" && btnLoadC.Font.Size == 8.5f, "standard button style");
        var btnFindC = (Button)Get(control, "btnFind");
        Check(btnFindC.Height == 30 && btnFindC.Width == 150 && !btnFindC.UseVisualStyleBackColor, "command buttons are 30 px high");
        Check(control.Font.OriginalFontName == "Segoe UI" && control.Font.Size == 8.5f, "Segoe UI 8.5 everywhere");
        var lblProg = (Label)Get(control, "lblProgress");
        Check(lblProg.Top == 40 && lblProg.Height == 30 && lblProg.ForeColor == Color.DimGray && lblProg.Padding.Top == 2, "status label at y40 h30, DimGray");
        Check(lblProg.Text.StartsWith("Load tables from the left pane, select a table, then click Find Dependencies."), "idle status sentence");

        // tables as cards
        var tables = new List<TableInfo> {
            EntityMetadataService.MapTable("account", "Account", false, false, false, false, "accounts"),
            EntityMetadataService.MapTable("new_project", "Project", true, false, false, false, "new_projects") };
        Set(control, "allTables", tables);
        Call(control, "BindTables"); Call(control, "UpdateState");
        var cardsT = (CardList)Get(control, "cardsTables");
        Check(cardsT.RenderedCards == 2, "table list shows 2 cards");
        Panel firstCard = (Panel)cardsT.Controls[0];
        Check(firstCard.Height == 62 && firstCard.BorderStyle == BorderStyle.FixedSingle && firstCard.Width == UiLayout.CardWidth(cardsT.ClientSize.Width), "card is 62 px high with a thin border and the spec width");
        Label cardTitle = (Label)firstCard.Controls[0]; Label cardSub = (Label)firstCard.Controls[1];
        Check(cardTitle.Font.Size == 9f && cardTitle.Font.Bold && cardTitle.ForeColor == Color.FromArgb(45, 45, 45) && cardTitle.AutoEllipsis && cardTitle.Height == 20 && cardTitle.Top == 7, "card title Segoe UI 9 bold (45,45,45)");
        Check(cardSub.Font.Size == 8f && cardSub.Font.Italic && cardSub.ForeColor == Color.DimGray && cardSub.Height == 19 && cardSub.Top == 31, "card subtitle Segoe UI 8 italic DimGray");
        Check(cardSub.Text.Contains(" \u00B7 ") && cardSub.Text.StartsWith("Standard"), "subtitle uses '{type} \u00B7 {detail}'");
        Check(lblSum.Text == "0 of 2 tables selected.", "live count: '0 of 2 tables selected.'");
        ((ComboBox)Get(control, "cmbTableKind")).SelectedIndex = 1;
        Check(cardsT.RenderedCards == 1, "custom-only filter leaves 1 card");
        ((ComboBox)Get(control, "cmbTableKind")).SelectedIndex = 0;
        txtSearch.Text = "acc"; Call(control, "BindTables");
        Check(cardsT.RenderedCards == 1, "table search filters by name");
        txtSearch.Text = ""; Call(control, "BindTables");

        // validation messages + card selection
        var btnFind = (Button)Get(control, "btnFind"); var lblVal = (Label)Get(control, "lblValidation");
        Check(!btnFind.Enabled && lblVal.Text.Contains("Select a table"), "table mode without a table: disabled + message");
        cardsT.ChooseIndex(0); Application.DoEvents();
        Check(((TableInfo)Get(control, "selectedTable")) != null, "clicking a table card sets the selected table");
        Check(((Panel)cardsT.Controls[0]).BackColor == Color.FromArgb(236, 244, 252) && ((Panel)cardsT.Controls[1]).BackColor == Color.White, "selected card is light blue, others white");
        Check(lblSum.Text == "1 of 2 tables selected.", "count updates to 1 of 2");
        Set(control, "columns", new List<ColumnInfo> { EntityMetadataService.MapColumn("name", "Account Name", "String", false, null), EntityMetadataService.MapColumn("new_status", "Status", "Picklist", true, null) });
        Call(control, "BindColumns"); Call(control, "UpdateState");
        Check(lblVal.Text.StartsWith("Ready: analyze table 'account'"), "valid table-mode selection reported as ready");
        ((RadioButton)Get(control, "rbColumn")).Checked = true; Application.DoEvents();
        Check(lblVal.Text.Contains("needs a column"), "column mode without a column: actionable message");
        var cardsC = (CardList)Get(control, "cardsColumns");
        Check(cardsC.Enabled && cardsC.RenderedCards == 2, "column cards enabled in column mode with 2 columns");
        ((TextBox)Get(control, "txtColumnSearch")).Text = "stat"; Call(control, "BindColumns");
        Check(cardsC.RenderedCards == 1, "column search filters");
        ((TextBox)Get(control, "txtColumnSearch")).Text = ""; Call(control, "BindColumns");
        cardsC.ChooseIndex(1); Application.DoEvents();
        Check(lblVal.Text.StartsWith("Ready: analyze column 'new_status'"), "column selection makes the request valid");
        var checks = (CheckBox[])Get(control, "scopeChecks");
        Check(checks.Length == 6 && checks.All(c => c.Checked), "six scope checkboxes, all ticked");
        ((Button)Get(control, "btnScopeNone")).PerformClick();
        Check(checks.All(c => !c.Checked) && lblVal.Text.Contains("component scope"), "Clear All -> message about scopes");
        ((Button)Get(control, "btnScopeAll")).PerformClick();
        Check(checks.All(c => c.Checked), "Select All ticks everything");
        ((RadioButton)Get(control, "rbTable")).Checked = true; Application.DoEvents();
        Check(!cardsC.Enabled, "column cards disabled in table mode");
        ((RadioButton)Get(control, "rbColumn")).Checked = true; Application.DoEvents();
        ((Button)Get(control, "btnClearSelection")).PerformClick(); Application.DoEvents();
        Check(((TableInfo)Get(control, "selectedTable")) == null && ((ColumnInfo)Get(control, "selectedColumn")) == null && cardsC.RenderedCards == 0 && lblSum.Text == "0 of 2 tables selected.", "Clear Selection clears the table, the column and the column list");
        cardsT.ChooseIndex(0); Application.DoEvents();
        Set(control, "columns", new List<ColumnInfo> { EntityMetadataService.MapColumn("new_status", "Status", "Picklist", true, null) });
        Call(control, "BindColumns"); cardsC.ChooseIndex(0); Application.DoEvents();
        ((RadioButton)Get(control, "rbTable")).Checked = true; Application.DoEvents();

        // card cap for very large environments
        var many = Enumerable.Range(0, 600).Select(i => new CardItem { Key = "t" + i, Title = "Table " + i, Subtitle = "Standard" }).ToList();
        var probe = new CardList { Width = 300, Height = 400 }; probe.SetItems(many, "t500");
        Check(probe.RenderedCards == 250 && probe.TotalItems == 600 && probe.SelectedKey == "t500", "very long lists are capped at 250 cards and keep the selection");
        probe.Dispose();

        // standard grid style (status grid)
        var gs = (DataGridView)Get(control, "gridStatus");
        Check(gs.ColumnHeadersDefaultCellStyle.BackColor == Color.FromArgb(241, 238, 224) && gs.ColumnHeadersDefaultCellStyle.SelectionBackColor == Color.FromArgb(241, 238, 224) && !gs.EnableHeadersVisualStyles, "grid header is beige and cannot turn blue");
        Check(gs.DefaultCellStyle.SelectionBackColor == Color.FromArgb(48, 105, 190) && gs.DefaultCellStyle.SelectionForeColor == Color.White && gs.GridColor == Color.FromArgb(220, 220, 220), "selection (48,105,190)/white, grid lines (220,220,220)");
        Check(gs.RowTemplate.Height == 28 && gs.ColumnHeadersHeight == 32 && gs.ColumnHeadersDefaultCellStyle.Font.Bold && !gs.RowHeadersVisible && gs.SelectionMode == DataGridViewSelectionMode.FullRowSelect && !gs.MultiSelect, "28 px rows, 32 px bold header, full-row single selection");
        Check(gs.Columns.Cast<DataGridViewColumn>().All(c => c.AutoSizeMode == DataGridViewAutoSizeColumnMode.Fill && c.MinimumWidth > 0), "every column is Fill with a minimum width");

        // skip-managed option; there is no persistent warning label any more (the only warning is the pop-up before the scan)
        var chkSkip = (CheckBox)Get(control, "chkSkipManaged"); var cmbJs = (ComboBox)Get(control, "cmbJsScope");
        Check(control.GetType().GetField("lblJsWarning", BindingFlags.NonPublic | BindingFlags.Instance) == null, "no persistent JavaScript warning label exists");
        Check(chkSkip.Enabled, "skip-managed option available for the full search");
        cmbJs.SelectedIndex = 1; Application.DoEvents();
        Check(!chkSkip.Enabled, "targeted scope disables the skip option");
        cmbJs.SelectedIndex = 0; Application.DoEvents();
        Check(chkSkip.Enabled, "full search enables the skip option again");
        var about = (LinkLabel)Get(control, "lnkAbout");
        Check(about.Text.Contains("Gagan Deep Singh") && about.Text.Contains("1.2026.11.11") && about.Links.Count == 1 && (string)about.Links[0].LinkData == "https://www.linkedin.com/in/glowofsky/", "maintainer link and version are shown");

        // report: empty states (the user is on the Impact Report tab for these checks)
        tabs.SelectedTab = tabs.TabPages[1]; Application.DoEvents();
        var lblEmpty = (Label)Get(control, "lblEmpty");
        Check(lblEmpty.Visible && lblEmpty.Text.Contains("No results yet"), "empty state before any scan");

        // Run a real analysis with stub scanners: one completed, one failed
        var req = new ScanRequest { Mode = SearchMode.Column, TableLogicalName = "account", ColumnLogicalName = "new_status", Scopes = new HashSet<ScannerKind>(new[] { ScannerKind.BusinessRules, ScannerKind.Plugins }), EnvironmentName = "Dev" };
        var svc = new DependencyAnalysisService(k => new Stub(k));
        var a = svc.Analyze(new NullService(), req, null, System.Threading.CancellationToken.None);
        Set(control, "analysis", a);
        Call(control, "BindStatus"); Call(control, "RefreshReport", true); Call(control, "UpdateState");
        var gridStatus = (DataGridView)Get(control, "gridStatus"); var gridResults = (DataGridView)Get(control, "gridResults");
        Check(gridStatus.Rows.Count == 6, "status grid lists all six categories");
        Check(gridStatus.SelectedRows.Count == 0 && gridStatus.CurrentCell == null, "no status row is selected until the user picks one");
        Check(((DataGridView)Get(control, "gridResults")).Columns.Cast<DataGridViewColumn>().All(c => c.AutoSizeMode == DataGridViewAutoSizeColumnMode.Fill) && ((DataGridView)Get(control, "gridResults")).ColumnHeadersDefaultCellStyle.BackColor == Color.FromArgb(241, 238, 224), "results grid uses the same standard style");
        Check(((TextBox)Get(control, "txtDetail")).Font.OriginalFontName == "Segoe UI", "detail box uses Segoe UI only");
        Check(gridResults.Rows.Count == 2, "results grid shows the 2 findings");
        Check(!lblEmpty.Visible, "empty state hidden when results exist");

        // filter + sort + detail
        var cmbConf = (ComboBox)Get(control, "cmbConfidence");
        cmbConf.SelectedIndex = 1; Application.DoEvents();
        Check(gridResults.Rows.Count == 1, "confidence filter narrows to 1");
        ((Button)Get(control, "btnResetFilters")).PerformClick();
        Check(gridResults.Rows.Count == 2, "Reset Filters restores both");
        ((TextBox)Get(control, "txtName")).Text = "no-such-name";
        Check(gridResults.Rows.Count == 0 && lblEmpty.Visible && lblEmpty.Text.Contains("match the current filters"), "filtered-out state explained");
        ((Button)Get(control, "btnResetFilters")).PerformClick();
        Call(control, "SortBy", "ComponentName");
        string first = (string)gridResults.Rows[0].Cells[1].Value;
        Call(control, "SortBy", "ComponentName");
        Check((string)gridResults.Rows[0].Cells[1].Value != first, "clicking a header twice reverses the sort");
        gridResults.Rows[0].Selected = true; gridResults.CurrentCell = gridResults.Rows[0].Cells[0]; Application.DoEvents();
        var detail = (TextBox)Get(control, "txtDetail");
        Check(detail.Text.Contains("Confidence:") && detail.Text.Contains("Why this result was included"), "detail panel shows evidence fields");
        Check(((Button)Get(control, "btnExport")).Enabled, "export enabled with results");

        // zero-match + failed category -> not a clean result
        var req2 = new ScanRequest { Mode = SearchMode.Table, TableLogicalName = "account", Scopes = new HashSet<ScannerKind>(new[] { ScannerKind.Plugins }) };
        var a2 = new DependencyAnalysisService(k => new Stub(k, true)).Analyze(new NullService(), req2, null, System.Threading.CancellationToken.None);
        Set(control, "analysis", a2); Call(control, "BindStatus"); Call(control, "RefreshReport", true);
        Check(lblEmpty.Visible && lblEmpty.Text.Contains("NOT a clean result"), "failed category + no matches is not shown as clean");
        var a3 = new DependencyAnalysisService(k => new Stub(k, false, true)).Analyze(new NullService(), req2, null, System.Threading.CancellationToken.None);
        Set(control, "analysis", a3); Call(control, "BindStatus"); Call(control, "RefreshReport", true);
        Check(lblEmpty.Text.Contains("No dependencies were found"), "all inspected + no matches shown as a clean empty result");

        // live findings while a scan is running
        Set(control, "analysis", null);
        Set(control, "liveAnalysis", new AnalysisResult { Request = req });
        Call(control, "RefreshReport", true);
        Check(lblEmpty.Visible && lblEmpty.Text.Contains("Scan in progress"), "live view with no matches yet says the scan is in progress");
        var liveResult = new DependencyResult { Source = ScannerKind.JavaScript, ComponentType = ComponentType.JavaScriptWebResource, ComponentName = "new_/js/live.js", ComponentId = "x", MatchType = "API call argument", MatchedValue = "new_status", Confidence = ConfidenceLevel.TextMatchReviewRequired, Evidence = "f.getAttribute('new_status')", LineNumber = 4 };
        ((AnalysisResult)Get(control, "liveAnalysis")).Results.Add(liveResult);
        Call(control, "RefreshReport", true);
        Check(gridResults.Rows.Count == 1 && !lblEmpty.Visible, "a streamed finding appears in the grid immediately");
        Check((string)gridResults.Rows[0].Cells[6].Value == "Running", "live rows are marked as running");
        Check(((Label)Get(control, "lblReportInfo")).Text.Contains("Scan in progress"), "report header says the results are live");
        gridResults.Rows[0].Selected = true; gridResults.CurrentCell = gridResults.Rows[0].Cells[0]; Application.DoEvents();
        ((AnalysisResult)Get(control, "liveAnalysis")).Results.Add(new DependencyResult { Source = ScannerKind.JavaScript, ComponentType = ComponentType.JavaScriptWebResource, ComponentName = "new_/js/other.js", ComponentId = "y", MatchType = "m", MatchedValue = "new_status", Confidence = ConfidenceLevel.TextMatchReviewRequired });
        Call(control, "RefreshReport", false);
        Check(gridResults.Rows.Count == 2 && gridResults.SelectedRows.Count == 1 && ((DataGridViewRow)gridResults.SelectedRows[0]).Cells[1].Value.ToString() == "new_/js/live.js", "selection is kept when more live findings arrive");
        Check(!((Button)Get(control, "btnExport")).Enabled, "export stays disabled until the scan has finished");
        Set(control, "liveAnalysis", null); Set(control, "analysis", null);

        // reset on a new environment
        Call(control, "ResetForNewEnvironment");
        Check(cardsT.TotalItems == 0 && lblEmpty.Text.Contains("No results yet"), "new environment clears tables and results");

        form.Close();
        Console.WriteLine(failures == 0 ? "UI SMOKE: all checks passed" : "UI SMOKE: " + failures + " check(s) FAILED");
        return failures == 0 ? 0 : 1;
    }

    class Stub : IDependencyScanner
    {
        readonly bool fail, clean;
        public Stub(ScannerKind k, bool fail = false, bool clean = false) { Kind = k; this.fail = fail; this.clean = clean; }
        public ScannerKind Kind { get; private set; }
        public ScanOutcome Scan(ScanContext c, System.Threading.CancellationToken ct)
        {
            var o = new ScanOutcome { Diagnostic = new ScanDiagnostic { Category = Kind, Status = fail ? ScanStatus.Failed : ScanStatus.Completed, RecordsExamined = 3 } };
            if (fail) o.Diagnostic.Error("simulated failure");
            if (!fail && !clean)
                o.Results.Add(new DependencyResult { Source = Kind, ComponentType = Kind == ScannerKind.Plugins ? ComponentType.PluginStep : ComponentType.BusinessRule, ComponentName = (Kind == ScannerKind.Plugins ? "Z-step" : "A-rule"), ComponentId = Guid.NewGuid().ToString(),
                    MatchType = "m", MatchedValue = "new_status", Location = "loc", Evidence = "evidence text", Explanation = "because",
                    Confidence = Kind == ScannerKind.Plugins ? ConfidenceLevel.ConfirmedMetadataReference : ConfidenceLevel.StrongDefinitionReference, ScanTimestampUtc = DateTime.UtcNow });
            return o;
        }
    }

    class NullService : Microsoft.Xrm.Sdk.IOrganizationService
    {
        public Guid Create(Microsoft.Xrm.Sdk.Entity e) { throw new NotSupportedException(); }
        public Microsoft.Xrm.Sdk.Entity Retrieve(string n, Guid id, Microsoft.Xrm.Sdk.Query.ColumnSet c) { throw new NotSupportedException(); }
        public void Update(Microsoft.Xrm.Sdk.Entity e) { throw new NotSupportedException(); }
        public void Delete(string n, Guid id) { throw new NotSupportedException(); }
        public Microsoft.Xrm.Sdk.OrganizationResponse Execute(Microsoft.Xrm.Sdk.OrganizationRequest r) { throw new NotSupportedException(); }
        public void Associate(string n, Guid id, Microsoft.Xrm.Sdk.Relationship r, Microsoft.Xrm.Sdk.EntityReferenceCollection c) { throw new NotSupportedException(); }
        public void Disassociate(string n, Guid id, Microsoft.Xrm.Sdk.Relationship r, Microsoft.Xrm.Sdk.EntityReferenceCollection c) { throw new NotSupportedException(); }
        public Microsoft.Xrm.Sdk.EntityCollection RetrieveMultiple(Microsoft.Xrm.Sdk.Query.QueryBase q) { throw new NotSupportedException(); }
    }
}
