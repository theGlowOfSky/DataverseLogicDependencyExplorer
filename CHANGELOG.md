# Changelog

**Maintainer:** Gagan Deep Singh  
**LinkedIn:** [Gagan Deep Singh](https://www.linkedin.com/in/glowofsky/)  
**Host:** XrmToolBox  
**Data platform:** Microsoft Dataverse  
**Export format:** Excel workbook (`.xlsx`)  
**Project:** [GitHub repository](https://github.com/theGlowOfSky/DataverseLogicDependencyExplorer)

## 1.2026.11.11
Initial release.
- Dependency Finder tab: searchable table/column selectors, table or column search mode, component scopes, cancellable scan with live progress, per-category status.
- Impact Report tab: filterable/sortable results, evidence detail panel, honest empty states, Excel export. Findings stream in live while a scan runs.
- Scanners: Business Rules, JavaScript Web Resources, Forms and Form Event Registrations, Plugin Registrations (metadata only), Classic Workflows and Processes, Cloud Flows (partial, Dataverse-stored flows only).
- JavaScript: parallel chunked download on cloned connections with overlapping analysis, substring pre-filter, option to skip managed web resources, warning and confirmation before a full search.
- Confidence classification: Confirmed metadata / Strong definition / Text match - review / Possible or dynamic.
- Excel workbook (dependency-free writer): Summary, Dependencies, Scan Diagnostics, Scan Configuration.
- Single-DLL deployment; host assemblies are compile-only (`ExcludeAssets=runtime`).
