# Architecture

**Maintainer:** Gagan Deep Singh  
**LinkedIn:** [Gagan Deep Singh](https://www.linkedin.com/in/glowofsky/)  
**Host:** XrmToolBox  
**Data platform:** Microsoft Dataverse  
**Export format:** Excel workbook (`.xlsx`)  
**Version:** 1.2026.11.11

```
MainControl (UI, 3 partial files)  -->  DependencyAnalysisService  -->  IDependencyScanner x6  -->  DependencyResult / ScanDiagnostic
        |                                      |                               |
        |                                      +-- ResultDeduplicator          +-- DataverseHelper (paging, retry, permission detection)
        +-- EntityMetadataService              +-- SolutionLookup (on demand)  +-- WebResourceRepository / FormRepository (shared caches)
        +-- ResultFilter / ResultDetailFormatter / AnalysisSummary
        +-- ExcelReportExporter --> XlsxWorkbookWriter (OpenXML + System.IO.Compression)
```

* **No UI code in the engine.** `DependencyAnalysisService.Analyze` takes an `IOrganizationService`, a `ScanRequest`, an `IProgress<ProgressInfo>`
  and a `CancellationToken`. The UI runs it with `Task.Run` (the pattern of the reference plugins) and updates controls from `Progress<T>` on the UI thread.
* **One scanner per category** (`ScannerBase` template). `ScannerBase.Scan` converts any exception into a status: `PermissionDenied`, `Failed` or
  `Cancelled`; scanners themselves set `PartiallySupported`. One failing scanner never stops the run. Unselected categories get a `NotInspected` diagnostic.
* **Common model.** `DependencyResult` (finding), `ScanDiagnostic` (per category), `AnalysisResult` (run). `AnalysisResult.IsCleanEmptyResult` is true only when
  there are no matches *and* every selected category completed - the UI and workbook use it so a failed/partial scan is never shown as "nothing found".
* **Late-bound only.** `Entity`, `QueryExpression`, `RetrieveMetadataChanges`/`RetrieveEntity`; no generated proxy classes.
* **Targeted queries and paging.** Business rules and (by default) classic processes filter by `primaryentity`; plugin steps join `sdkmessagefilter.primaryobjecttypecode`;
  forms filter by `objecttypecode`; web resources by type (or by name in the targeted JavaScript scope). Everything pages through `DataverseHelper.RetrieveAll`.
  Throttling errors are retried honouring `Retry-After`.
* **Shared caches per scan.** Web resource content is downloaded once and reused by the JavaScript and Form scanners; forms of the table are read once.
* **Parsers.** `JsSourceAnalyzer` (tokenizer separating code/comments/strings/regex/templates), `FormXmlReader`, `XmlPath` (parsed-value matching with node paths),
  `BusinessRuleDefinition` (client code), `CloudFlowDefinition` (flow JSON with JSON paths).
* **Excel.** `XlsxWorkbookWriter` is a small multi-sheet extension of the approach of the Business Rule Exporter's `XlsxWriter`: no third-party libraries, so the
  plugin ships as one DLL and cannot cause assembly version conflicts with other plugins. Text is always stored as strings (never formulas).
* **Registration.** `Plugin` is `[Export(typeof(IXrmToolBoxPlugin))]` deriving from `PluginBase` and returns `MainControl : PluginControlBase`, the same pattern as the reference plugins.
  `XrmToolBoxPackage`, `MscrmTools.Xrm.Connection` and `Newtonsoft.Json` use `PrivateAssets="all" ExcludeAssets="runtime"`: compile-time only.

## JavaScript pipeline (all script web resources)
`JavaScriptScanner` avoids "download everything, then analyze":
1. **List** script web resources with a cheap query (no `content`, 5000 per page; optional `ismanaged = false`).
2. **Download** in chunks of 10 ids (`webresourceid In (...)`) by several workers. Each worker uses its own cloned connection (`ServiceCloner` calls the public
   parameterless `Clone()` of the Dataverse client by reflection); if cloning is not available one worker uses the original connection.
3. **Analyze** concurrently: workers push decoded scripts into a bounded `BlockingCollection`; up to four consumer threads run a substring pre-filter and then the
   `JsSourceAnalyzer`. Memory stays bounded; only libraries registered on the table's forms are retained for the Form scanner.
4. **Stream**: each consumer calls `ScanContext.Emit` with its findings and `ScanContext.ReportProgress` (throttled to ~5/s). `DependencyAnalysisService` forwards both through
   `IProgress<ProgressInfo>`; the UI shows provisional findings in the Impact Report and replaces them with the de-duplicated, solution-enriched result at the end.
5. **Fail/cancel safely**: a fault in any worker cancels the linked token, is rethrown after all tasks end and is mapped to `Failed` / `PermissionDenied`; user cancellation
   becomes `Cancelled` and keeps the partial results.
