// LOCAL DEVELOPMENT HELPER: renders the control with sample data to PNG files (Mono, Xvfb; fonts fall back to DejaVu Sans).
using System; using System.Collections.Generic; using System.Drawing; using System.Drawing.Imaging; using System.Linq; using System.Reflection; using System.Windows.Forms;
using DataverseLogicDependencyExplorer; using DataverseLogicDependencyExplorer.Models; using DataverseLogicDependencyExplorer.Services; using DataverseLogicDependencyExplorer.Services.Scanners;
static class Shot {
  static object Get(object o, string n) { return o.GetType().GetField(n, BindingFlags.NonPublic | BindingFlags.Instance).GetValue(o); }
  static void Set(object o, string n, object v) { o.GetType().GetField(n, BindingFlags.NonPublic | BindingFlags.Instance).SetValue(o, v); }
  static object Call(object o, string n, params object[] a) { return o.GetType().GetMethod(n, BindingFlags.NonPublic | BindingFlags.Instance).Invoke(o, a); }
  [STAThread] static void Main(string[] args) {
    var control = (MainControl)new Plugin().GetControl();
    var form = new Form { Width = 1250, Height = 760, StartPosition = FormStartPosition.Manual, Left = 0, Top = 0 };
    form.Controls.Add(control); form.Show(); Application.DoEvents();
    var names = new[] { "account:Account", "contact:Contact", "new_project:Project", "opportunity:Opportunity", "lead:Lead", "systemuser:User", "new_task:Task Item", "incident:Case", "new_very_long_custom_table_name_for_testing:A very long display name to test ellipsis" };
    var tables = names.Select(n => { var p = n.Split(':'); return EntityMetadataService.MapTable(p[0], p[1], p[0].StartsWith("new_"), false, false, false, p[0] + "s"); }).ToList();
    Set(control, "allTables", tables); Call(control, "BindTables");
    var cardsT = Get(control, "cardsTables"); cardsT.GetType().GetMethod("ChooseIndex").Invoke(cardsT, new object[] { 0 });
    Set(control, "columns", new List<ColumnInfo> { EntityMetadataService.MapColumn("name", "Account Name", "String", false, null), EntityMetadataService.MapColumn("new_status", "Status", "Picklist", true, null), EntityMetadataService.MapColumn("telephone1", "Main Phone", "String", false, null) });
    Call(control, "BindColumns");
    ((RadioButton)Get(control, "rbColumn")).Checked = true;
    var cardsC = Get(control, "cardsColumns"); cardsC.GetType().GetMethod("ChooseIndex").Invoke(cardsC, new object[] { 1 });
    // a finished analysis with a few findings
    var req = new ScanRequest { Mode = SearchMode.Column, TableLogicalName = "account", ColumnLogicalName = "new_status", Scopes = new HashSet<ScannerKind>(ScanRequest.AllScopes), EnvironmentName = "Contoso Dev" };
    var svc = new DependencyAnalysisService(k => new Stub(k));
    var a = svc.Analyze(null == null ? new Null() : null, req, null, System.Threading.CancellationToken.None);
    Set(control, "analysis", a); Call(control, "BindStatus"); Call(control, "RefreshReport", true); Call(control, "UpdateState");
    ((Label)Get(control, "lblSummary")).Text = AnalysisSummary.Line(a);
    ((Label)Get(control, "lblProgress")).Text = "Scan finished: " + a.MatchCount + " finding(s). Open the 'Impact Report' tab to filter and export them.";
    Wait(40);
    var tabs = (TabControl)Get(control, "tabs"); tabs.SelectedTab = tabs.TabPages[1]; Application.DoEvents();
    var grid = (DataGridView)Get(control, "gridResults"); if (grid.Rows.Count > 0) { grid.Rows[0].Selected = true; grid.CurrentCell = grid.Rows[0].Cells[0]; }
    Wait(40);
    form.Close();
  }
  static void Wait(int n) { for (int i = 0; i < n; i++) { Application.DoEvents(); System.Threading.Thread.Sleep(100); } }
  static void Save(Form f, string path) { var bmp = new Bitmap(f.ClientSize.Width, f.ClientSize.Height); f.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height)); bmp.Save(path, ImageFormat.Png); Console.WriteLine("saved " + path); }
  class Stub : IDependencyScanner { public Stub(ScannerKind k) { Kind = k; } public ScannerKind Kind { get; private set; }
    public ScanOutcome Scan(DataverseLogicDependencyExplorer.Services.Scanners.ScanContext c, System.Threading.CancellationToken ct) {
      var st = Kind == ScannerKind.CloudFlows ? ScanStatus.PartiallySupported : Kind == ScannerKind.Plugins ? ScanStatus.PermissionDenied : ScanStatus.Completed;
      var o = new ScanOutcome { Diagnostic = new ScanDiagnostic { Category = Kind, Status = st, RecordsExamined = st == ScanStatus.PermissionDenied ? 0 : 12 } };
      if (st == ScanStatus.PermissionDenied) o.Diagnostic.Error("Permission denied: the connected user cannot read the data this scanner needs.");
      if (st == ScanStatus.Completed) o.Results.Add(new DependencyResult { Source = Kind, ComponentType = Kind == ScannerKind.Forms ? ComponentType.FormEventRegistration : Kind == ScannerKind.JavaScript ? ComponentType.JavaScriptWebResource : ComponentType.BusinessRule, ComponentName = Kind + " sample component", ComponentId = Guid.NewGuid().ToString(), ParentTable = "account", MatchType = "Event registered against the column (onchange)", MatchedValue = "new_status", TargetColumn = "new_status", Location = "/form/events/event[@name='onchange']", Evidence = "onchange of 'new_status' -> new_/js/account.js :: NS.Account.statusChanged (enabled)", Explanation = "A handler is registered for the onchange event of column 'new_status'.", Limitation = "Static text match.", SolutionName = "Core", Confidence = Kind == ScannerKind.JavaScript ? ConfidenceLevel.TextMatchReviewRequired : ConfidenceLevel.StrongDefinitionReference, LineNumber = Kind == ScannerKind.JavaScript ? (int?)42 : null, ScanTimestampUtc = DateTime.UtcNow });
      return o; } }
  class Null : Microsoft.Xrm.Sdk.IOrganizationService { public Guid Create(Microsoft.Xrm.Sdk.Entity e) { throw new NotSupportedException(); } public Microsoft.Xrm.Sdk.Entity Retrieve(string n, Guid i, Microsoft.Xrm.Sdk.Query.ColumnSet c) { throw new NotSupportedException(); } public void Update(Microsoft.Xrm.Sdk.Entity e) { throw new NotSupportedException(); } public void Delete(string n, Guid i) { throw new NotSupportedException(); } public Microsoft.Xrm.Sdk.OrganizationResponse Execute(Microsoft.Xrm.Sdk.OrganizationRequest r) { throw new NotSupportedException(); } public void Associate(string n, Guid i, Microsoft.Xrm.Sdk.Relationship r, Microsoft.Xrm.Sdk.EntityReferenceCollection c) { throw new NotSupportedException(); } public void Disassociate(string n, Guid i, Microsoft.Xrm.Sdk.Relationship r, Microsoft.Xrm.Sdk.EntityReferenceCollection c) { throw new NotSupportedException(); } public Microsoft.Xrm.Sdk.EntityCollection RetrieveMultiple(Microsoft.Xrm.Sdk.Query.QueryBase q) { throw new NotSupportedException(); } }
}
