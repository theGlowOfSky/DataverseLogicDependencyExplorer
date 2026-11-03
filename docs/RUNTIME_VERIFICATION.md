# Runtime verification checklist (not yet executed)

**Maintainer:** Gagan Deep Singh  
**LinkedIn:** [Gagan Deep Singh](https://www.linkedin.com/in/glowofsky/)  
**Host:** XrmToolBox  
**Data platform:** Microsoft Dataverse  
**Export format:** Excel workbook (`.xlsx`)  
**Version:** 1.2026.11.11

Compiling and unit tests do **not** prove that XrmToolBox loads the plugin or that every Dataverse query behaves as assumed. Complete this list on Windows
with XrmToolBox and a test environment before release, and record the result.

## Host load
- [ ] `dotnet build -c Release` succeeds with the real `XrmToolBoxPackage` 1.2025.10.74 / `MscrmTools.Xrm.Connection` 1.2025.9.64.
- [ ] Output folder contains only `DataverseLogicDependencyExplorer.dll` (+ pdb).
- [ ] Copy the DLL into `%APPDATA%\MscrmTools\XrmToolBox\Plugins`; XrmToolBox starts without a `CompositionContractMismatchException` and lists the tool with its icon.
- [ ] Open the tool, connect, reconnect to another environment: tables/results are cleared and the environment label updates.
- [ ] Close the tool during a scan: the scan is cancelled and XrmToolBox does not hang.

## Dataverse behaviour (assumptions to confirm)
- [ ] Table list loads (`RetrieveMetadataChanges`); custom/standard flag and entity set names are correct; private/intersect tables are hidden.
- [ ] Column list for a table (`RetrieveEntity` with attributes) - shadow columns (`*name`, `*yominame`) are hidden.
- [ ] Business rules: `workflow` category 2, `primaryentity` filter and the `clientdata` `<clientcode>` format match your environment (as used by Business Rule Exporter); `xaml` is readable.
- [ ] Plugin steps: `sdkmessagefilter.primaryobjecttypecode` equals the table logical name; `filteringattributes`, image `attributes`, `eventhandler`->`plugintype`->`pluginassembly` resolve.
- [ ] Forms: `systemform` with `objecttypecode`, `formxml`, `formactivationstate`, `componentstate = 0` filter accepted.
- [ ] JavaScript pipeline on a large environment: listing returns quickly; downloads run on several connections (check the status text '... download connection(s)'); if the host connection cannot be cloned the scan still completes with one connection; no 429 throttling storms (otherwise reduce `DownloadWorkers`).
- [ ] Live view: findings appear in the Impact Report while scanning, the UI stays responsive, Cancel Scan returns promptly and keeps partial results; the all-web-resources confirmation dialog appears and 'No' starts nothing.
- [ ] Web resources: `webresourcetype = 3`, base64 `content`, `componentstate = 0` accepted; large environments finish in acceptable time (try the targeted JavaScript scope).
- [ ] Classic processes: `category in (0,1,3,4)`, `type = 1`; XAML `EntityName`/`Attribute` shapes produce the expected step matches.
- [ ] Cloud flows: `workflow` category 5 returns solution-aware flows with `clientdata` JSON; Dataverse connector parameter names (`entityName`, `item/<col>`, `$select`, `subscriptionRequest/*`) match.
- [ ] Solution names resolve through `solutioncomponent` (types 29, 60, 61, 92).
- [ ] With a low-privilege user: each affected category shows **Permission denied**, other categories continue, and the Excel report records it.
- [ ] Throttling (large tenant): retries succeed or the category fails visibly.

## Excel
- [ ] The exported workbook opens in desktop Excel without a repair prompt; sheets *Summary, Dependencies, Scan Diagnostics, Scan Configuration*; frozen header; filters; dates readable.
- [ ] Export to a read-only folder / open file produces the friendly error dialog; Save dialog cancel does nothing.
