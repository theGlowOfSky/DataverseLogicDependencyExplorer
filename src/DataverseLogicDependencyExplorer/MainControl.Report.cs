// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DataverseLogicDependencyExplorer.Models;
using DataverseLogicDependencyExplorer.Reporting;
using DataverseLogicDependencyExplorer.Services;
using DataverseLogicDependencyExplorer.ViewModels;
using WinFormsLabel = System.Windows.Forms.Label;

namespace DataverseLogicDependencyExplorer
{
    /// <summary>Tab 2: Impact Report - filterable results, evidence detail and Excel export.</summary>
    public partial class MainControl
    {
        private WinFormsLabel lblReportInfo, lblReportCount, lblEmpty;
        private ComboBox cmbType, cmbConfidence, cmbStatus, cmbSolution;
        private TextBox txtName, txtMatched, txtSearchResults, txtDetail;
        private DataGridView gridResults;
        private readonly BindingSource resultsSource = new BindingSource();
        private Button btnExport, btnResetFilters;
        private SplitContainer reportSplit;
        private bool filtering;
        private List<ResultRow> shownRows = new List<ResultRow>();
        private string sortColumn;
        private bool sortAscending = true;

        private void BuildReportTab(TabPage page)
        {
            var layout = NewStack();
            layout.Padding = new Padding(8);

            lblReportInfo = new WinFormsLabel { ForeColor = MutedColor, Margin = new Padding(0, 0, 0, 4) };
            WrapLabel(lblReportInfo, layout);

            // Filters: two rows of label + control pairs.
            var filters = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 4, RowCount = 4, Margin = new Padding(0) };
            for (int i = 0; i < 4; i++) filters.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));

            cmbType = Combo("All component types", Enum.GetValues(typeof(ComponentType)).Cast<ComponentType>().Select(t => t.Display()));
            cmbConfidence = Combo("All confidence levels", Enum.GetValues(typeof(ConfidenceLevel)).Cast<ConfidenceLevel>().Select(c => c.Display()));
            cmbStatus = Combo("All scan statuses", Enum.GetValues(typeof(ScanStatus)).Cast<ScanStatus>().Select(s => s.Display()));
            cmbSolution = Combo("All solutions", new string[0]);
            txtName = new TextBox { Dock = DockStyle.Fill };
            txtMatched = new TextBox { Dock = DockStyle.Fill };
            txtSearchResults = new TextBox { Dock = DockStyle.Fill };
            foreach (ComboBox c in new[] { cmbType, cmbConfidence, cmbStatus, cmbSolution }) c.SelectedIndexChanged += (s, e) => RefreshReport(false);
            foreach (TextBox t in new[] { txtName, txtMatched, txtSearchResults }) t.TextChanged += (s, e) => RefreshReport(false);

            AddFilter(filters, 0, 0, "Component type", cmbType);
            AddFilter(filters, 1, 0, "Component name", txtName);
            AddFilter(filters, 2, 0, "Matched table or column", txtMatched);
            AddFilter(filters, 3, 0, "Confidence level", cmbConfidence);
            AddFilter(filters, 0, 2, "Scan status", cmbStatus);
            AddFilter(filters, 1, 2, "Solution", cmbSolution);
            AddFilter(filters, 2, 2, "Search all result fields", txtSearchResults);

            var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true, Margin = new Padding(0, 14, 0, 0) };
            btnResetFilters = Btn("Reset Filters", 90, (s, e) => ResetFilters());
            btnExport = Btn("Export Impact Report to Excel", 190, (s, e) => ExportExcel());
            var tip = new ToolTip();
            tip.SetToolTip(btnExport, "Exports the complete analysis (all findings, diagnostics and configuration), regardless of the filters set here.");
            actions.Controls.AddRange(new Control[] { btnExport, btnResetFilters });
            filters.Controls.Add(actions, 3, 2);
            filters.SetRowSpan(actions, 2);

            lblReportCount = new WinFormsLabel { Margin = new Padding(0, 6, 0, 4), ForeColor = TextColor };
            WrapLabel(lblReportCount, layout);

            // grid + empty state overlay, and the detail panel below it
            reportSplit = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal, SplitterWidth = 5, BorderStyle = BorderStyle.None, BackColor = Color.White };
            reportSplit.Panel1.BackColor = Color.White;
            reportSplit.Panel2.BackColor = Color.White;

            gridResults = NewGrid(true);
            gridResults.DataSource = resultsSource;
            resultsSource.DataSource = new BindingList<ResultRow>();
            AddResultColumn("Component type", "ComponentType", 0.9, 120);
            AddResultColumn("Component name", "ComponentName", 1.5, 180);
            AddResultColumn("Parent table", "ParentTable", 0.7, 100);
            AddResultColumn("Matched value", "MatchedValue", 0.9, 110);
            AddResultColumn("Match type", "MatchType", 1.3, 160);
            AddResultColumn("Confidence", "Confidence", 1.1, 150);
            AddResultColumn("Scan status", "ScanStatus", 0.8, 100);
            AddResultColumn("Solution", "Solution", 0.9, 110);
            AddResultColumn("Location", "Location", 1.0, 130);
            gridResults.ColumnHeaderMouseClick += (s, e) => SortBy(gridResults.Columns[e.ColumnIndex].DataPropertyName);
            gridResults.SelectionChanged += (s, e) => ShowDetail();
            gridResults.CellToolTipTextNeeded += (s, e) =>
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
                var row = RowItem<ResultRow>(gridResults, e.RowIndex);
                if (row == null) return;
                string full;
                if (row.Full.TryGetValue(gridResults.Columns[e.ColumnIndex].DataPropertyName, out full)) e.ToolTipText = full; // multi-line text: the cell shows one line
                else
                {
                    string v = gridResults[e.ColumnIndex, e.RowIndex].FormattedValue as string;
                    if (v != null && v.Length > 30) e.ToolTipText = v;
                }
            };

            lblEmpty = new WinFormsLabel
            {
                Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, BackColor = Color.White,
                ForeColor = MutedColor, Padding = new Padding(30), Visible = true, Font = new Font("Segoe UI", 8.5F)
            };
            reportSplit.Panel1.Controls.Add(gridResults);
            reportSplit.Panel1.Controls.Add(lblEmpty);
            lblEmpty.BringToFront();

            txtDetail = new TextBox
            {
                Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, WordWrap = true,
                BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Font = new Font("Segoe UI", 8.5F)
            };
            var detailHost = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 4, 0, 0), BackColor = Color.White };
            detailHost.Controls.Add(txtDetail);
            reportSplit.Panel2.Controls.Add(detailHost);
            reportSplit.SizeChanged += (s, e) =>
            {
                if (reportSplit.Height > 260)
                {
                    try { reportSplit.Panel1MinSize = 120; reportSplit.Panel2MinSize = 100; reportSplit.SplitterDistance = (int)(reportSplit.Height * 0.58); }
                    catch (InvalidOperationException) { }
                }
            };

            AddRow(layout, lblReportInfo);
            AddRow(layout, filters);
            AddRow(layout, lblReportCount);
            AddRow(layout, reportSplit, true);
            page.Controls.Add(layout);

            RefreshReport(true);
        }

        private ComboBox Combo(string allText, IEnumerable<string> items)
        {
            var c = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            c.Items.Add(allText);
            foreach (string i in items) c.Items.Add(i);
            c.SelectedIndex = 0;
            return c;
        }

        private static void AddFilter(TableLayoutPanel host, int col, int row, string label, Control input)
        {
            host.Controls.Add(new WinFormsLabel { Text = label, AutoSize = true, Margin = new Padding(0, 4, 8, 2) }, col, row);
            input.Margin = new Padding(0, 0, 8, 4);
            host.Controls.Add(input, col, row + 1);
        }

        private void AddResultColumn(string header, string property, double weight, int minWidth)
        {
            gridResults.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = header, DataPropertyName = property, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = (float)(weight * 100), MinimumWidth = minWidth,
                SortMode = DataGridViewColumnSortMode.Programmatic, // sorted by SortBy(): the confidence column sorts by strength, not alphabetically
                ReadOnly = true
            });
        }

        // ------------------------------------------------------------------ filtering, sorting, binding

        /// <summary>The finished analysis, or the provisional live findings while a scan is still running.</summary>
        private AnalysisResult ViewAnalysis { get { return analysis ?? liveAnalysis; } }

        private void UpdateReportControls()
        {
            if (btnExport == null) return;
            btnExport.Enabled = !busy && analysis != null;   // only a finished scan can be exported
            btnResetFilters.Enabled = ViewAnalysis != null;
        }

        private void ResetFilters()
        {
            filtering = true;
            try
            {
                cmbType.SelectedIndex = cmbConfidence.SelectedIndex = cmbStatus.SelectedIndex = cmbSolution.SelectedIndex = 0;
                txtName.Text = txtMatched.Text = txtSearchResults.Text = string.Empty;
            }
            finally { filtering = false; }
            RefreshReport(false);
        }

        private FilterCriteria CurrentFilter()
        {
            return new FilterCriteria
            {
                ComponentType = cmbType.SelectedIndex > 0 ? (ComponentType?)Enum.GetValues(typeof(ComponentType)).Cast<ComponentType>().ElementAt(cmbType.SelectedIndex - 1) : null,
                Confidence = cmbConfidence.SelectedIndex > 0 ? (ConfidenceLevel?)Enum.GetValues(typeof(ConfidenceLevel)).Cast<ConfidenceLevel>().ElementAt(cmbConfidence.SelectedIndex - 1) : null,
                ScanStatus = cmbStatus.SelectedIndex > 0 ? (ScanStatus?)Enum.GetValues(typeof(ScanStatus)).Cast<ScanStatus>().ElementAt(cmbStatus.SelectedIndex - 1) : null,
                Solution = cmbSolution.SelectedIndex > 0 ? (string)cmbSolution.SelectedItem : null,
                ComponentName = txtName.Text,
                MatchedTableOrColumn = txtMatched.Text,
                SearchText = txtSearchResults.Text
            };
        }

        /// <summary>Re-applies the filters. When <paramref name="newData"/> is true the solution list and the header are rebuilt.</summary>
        private void RefreshReport(bool newData)
        {
            if (gridResults == null || filtering) return;
            AnalysisResult view = ViewAnalysis;
            bool live = analysis == null && liveAnalysis != null;

            if (newData)
            {
                filtering = true;
                try
                {
                    string keep = cmbSolution.SelectedIndex > 0 ? (string)cmbSolution.SelectedItem : null;
                    cmbSolution.Items.Clear();
                    cmbSolution.Items.Add("All solutions");
                    foreach (string s in ResultFilter.DistinctSolutions(view)) cmbSolution.Items.Add(s);
                    cmbSolution.SelectedIndex = keep != null && cmbSolution.Items.Contains(keep) ? cmbSolution.Items.IndexOf(keep) : 0;
                }
                finally { filtering = false; }
                lblReportInfo.Text = view == null
                    ? "No scan has been run yet."
                    : live
                        ? "Scan in progress - showing live findings (provisional: duplicates are removed and solutions are resolved when the scan finishes). Found so far: " + view.MatchCount
                        : DescribeScan(analysis);
            }

            // Keep the user's selection and scroll position while live findings keep arriving.
            var previous = gridResults.SelectedRows.Count == 0 ? null : gridResults.SelectedRows[0].DataBoundItem as ResultRow;
            int firstVisible = gridResults.Rows.Count > 0 ? gridResults.FirstDisplayedScrollingRowIndex : -1;

            List<DependencyResult> matches = ResultFilter.Apply(view, CurrentFilter());
            shownRows = matches.Select(r =>
            {
                ResultRow row = ResultRow.From(r, view.DiagnosticFor(r.Source));
                if (live) row.ScanStatus = "Running";
                return row;
            }).ToList();
            ApplySort();

            resultsSource.DataSource = new BindingList<ResultRow>(shownRows);
            ClearGridSelection(gridResults);
            if (previous != null)
            {
                int idx = shownRows.FindIndex(r => ReferenceEquals(r.Result, previous.Result));
                if (idx >= 0)
                {
                    gridResults.Rows[idx].Selected = true;
                    try { gridResults.CurrentCell = gridResults.Rows[idx].Cells[0]; } catch (InvalidOperationException) { }
                }
            }
            if (live && firstVisible >= 0 && firstVisible < gridResults.Rows.Count)
            {
                try { gridResults.FirstDisplayedScrollingRowIndex = firstVisible; } catch (InvalidOperationException) { }
            }

            int total = view == null ? 0 : view.MatchCount;
            lblReportCount.Text = view == null ? string.Empty : "Showing " + shownRows.Count + " of " + total + " finding(s)" + (live ? " - scan in progress" : string.Empty);
            UpdateEmptyState();
            ShowDetail();
            UpdateReportControls();
        }

        private static string DescribeScan(AnalysisResult a)
        {
            ScanRequest q = a.Request;
            return "Scan of " + (q.Mode == SearchMode.Column ? "column " + q.TableLogicalName + "." + q.ColumnLogicalName : "table " + q.TableLogicalName)
                 + " at " + a.StartedUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") + " (local time). " + AnalysisSummary.Line(a);
        }

        /// <summary>Distinguishes "nothing run", "filtered out", "clean empty result" and "empty because a category could not be inspected".</summary>
        private void UpdateEmptyState()
        {
            string text = null;
            Color color = MutedColor;

            AnalysisResult view = ViewAnalysis;
            if (view == null)
                text = "No results yet.\n\nRun a scan on the 'Dependency Finder' tab.";
            else if (analysis == null && view.MatchCount == 0)
                text = "Scan in progress.\n\nNo matches yet. Findings appear here as soon as they are discovered.";
            else if (analysis != null && analysis.MatchCount == 0)
            {
                if (analysis.IsCleanEmptyResult)
                    text = "No dependencies were found.\n\nEvery selected component category was inspected and none references the selected " +
                           (analysis.Request.Mode == SearchMode.Column ? "column" : "table") + ".\nCategories that were not selected were not inspected.";
                else
                {
                    var notes = analysis.Diagnostics.Where(d => analysis.Request.Scopes.Contains(d.Category) &&
                                                                d.Status != ScanStatus.Completed && d.Status != ScanStatus.CompletedWithWarnings)
                                                    .Select(d => "  - " + d.Category.Display() + ": " + d.Status.Display()).ToList();
                    text = "No matches were found - but this is NOT a clean result.\n\nSome selected categories were not fully inspected:\n" +
                           string.Join("\n", notes) + "\n\nSee the status table on the 'Dependency Finder' tab, or the Scan Diagnostics sheet of the Excel report.";
                    color = WarnColor;
                }
            }
            else if (shownRows.Count == 0)
                text = "No findings match the current filters.\n\nClick 'Reset Filters' to see all " + view.MatchCount + " finding(s).";

            lblEmpty.Text = text ?? string.Empty;
            lblEmpty.ForeColor = color;
            lblEmpty.Visible = text != null;
            if (text != null) lblEmpty.BringToFront();
        }

        private void SortBy(string property)
        {
            if (string.IsNullOrEmpty(property)) return;
            sortAscending = sortColumn == property ? !sortAscending : true;
            sortColumn = property;
            RefreshReport(false);
            foreach (DataGridViewColumn c in gridResults.Columns)
                c.HeaderCell.SortGlyphDirection = c.DataPropertyName == sortColumn ? (sortAscending ? SortOrder.Ascending : SortOrder.Descending) : SortOrder.None;
        }

        private void ApplySort()
        {
            if (sortColumn == null) return;
            Func<ResultRow, object> key;
            switch (sortColumn)
            {
                case "Confidence": key = r => r.ConfidenceOrder; break;
                case "ComponentType": key = r => r.ComponentType; break;
                case "ComponentName": key = r => r.ComponentName; break;
                case "ParentTable": key = r => r.ParentTable; break;
                case "MatchedValue": key = r => r.MatchedValue; break;
                case "MatchType": key = r => r.MatchType; break;
                case "ScanStatus": key = r => r.ScanStatus; break;
                case "Solution": key = r => r.Solution; break;
                default: key = r => r.Location; break;
            }
            shownRows = (sortAscending ? shownRows.OrderBy(key) : shownRows.OrderByDescending(key)).ToList();
        }

        private void ShowDetail()
        {
            if (txtDetail == null) return;
            var row = gridResults.SelectedRows.Count == 0 ? null : gridResults.SelectedRows[0].DataBoundItem as ResultRow;
            if (row == null)
            {
                txtDetail.Text = shownRows.Count == 0 ? string.Empty : "Select a finding to see its evidence, confidence and limitations.";
                return;
            }
            AnalysisResult view = ViewAnalysis;
            txtDetail.Text = ResultDetailFormatter.Format(row.Result, view == null ? null : view.DiagnosticFor(row.Result.Source));
            txtDetail.SelectionStart = 0;
            txtDetail.ScrollToCaret();
        }

        // ------------------------------------------------------------------ export

        private async void ExportExcel()
        {
            if (busy || analysis == null) return;

            string path;
            using (var dlg = new SaveFileDialog
            {
                Filter = "Excel workbook (*.xlsx)|*.xlsx",
                FileName = ExcelReportExporter.DefaultFileName(DateTime.Now),
                Title = "Export Impact Report",
                OverwritePrompt = true
            })
            {
                if (dlg.ShowDialog(this) != DialogResult.OK)
                {
                    lblReportCount.Text = "Export cancelled.";
                    return;
                }
                path = dlg.FileName;
            }

            AnalysisResult snapshot = analysis;
            SetBusy(true, false);
            lblReportCount.Text = "Exporting " + snapshot.MatchCount + " finding(s)...";

            try
            {
                ExportOutcome outcome = await Task.Run(() => ExcelReportExporter.Export(snapshot, path, ProductInfo.Version));
                string msg = "Exported " + snapshot.MatchCount + " finding(s), the scan diagnostics and the configuration to:\n" + outcome.Path;
                if (outcome.TruncatedCells > 0)
                    msg += "\n\n" + outcome.TruncatedCells + " cell(s) exceeded Excel's 32,767 character limit and are marked [TRUNCATED].";
                if (!snapshot.IsCleanEmptyResult && snapshot.MatchCount == 0)
                    msg += "\n\nNote: the scan found nothing, but not every selected category was fully inspected. The workbook says so.";
                lblReportCount.Text = "Exported " + snapshot.MatchCount + " finding(s) to " + Path.GetFileName(outcome.Path) + ".";
                MessageBox.Show(this, msg, "Export complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                LogFailure("Export failed: " + ex);
                lblReportCount.Text = "Export failed: " + ex.Message;
                lblReportCount.ForeColor = ErrorColor;
                string hint = ex is IOException || ex is UnauthorizedAccessException
                    ? "\n\n(If the file is open in Excel, close it, or choose a different location, and retry.)"
                    : string.Empty;
                MessageBox.Show(this, "Export failed:\n\n" + ex.Message + hint, "Export failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                SetBusy(false);
                lblReportCount.ForeColor = TextColor;
            }
        }
    }
}
