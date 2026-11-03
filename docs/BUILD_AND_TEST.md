# Build and test

**Maintainer:** Gagan Deep Singh  
**LinkedIn:** [Gagan Deep Singh](https://www.linkedin.com/in/glowofsky/)  
**Host:** XrmToolBox  
**Data platform:** Microsoft Dataverse  
**Export format:** Excel workbook (`.xlsx`)  
**Version:** 1.2026.11.11

## Windows (supported path)
```powershell
dotnet restore DataverseLogicDependencyExplorer.sln
dotnet build   DataverseLogicDependencyExplorer.sln -c Release
dotnet test    DataverseLogicDependencyExplorer.sln -c Release
powershell -ExecutionPolicy Bypass -File build\Package-Plugin.ps1
```
Expected build output (`src\DataverseLogicDependencyExplorer\bin\Release`): `DataverseLogicDependencyExplorer.dll` (+ `.pdb`) and **no other DLLs**.
`Package-Plugin.ps1` enforces that and writes `artifacts\...-deployment.zip` (and the `.nupkg` when `nuget.exe` is on PATH).

## Unit tests
MSTest, `net48`. They use `FakeOrganizationService` (canned rows, records queries) - **no live Dataverse is contacted**. They cover:
JavaScript analysis and the parallel download pipeline (chunking, several workers, live streaming, skip-managed, failure, permission, cancellation), connection cloning, live progress mapping, JavaScript analysis (comments, strings, identifiers, line numbers, dynamic access, function bodies), form XML, business-rule client code, flow JSON,
every scanner (targeted queries, confidence, malformed definitions, permission denied, failures, cancellation), the analysis service (NotInspected, partial
failure, cancellation, dedupe, clean-vs-unclean empty results), metadata mapping/filtering, result filtering, request validation and the Excel workbook
(sheets, headers, frozen header, filter, widths, special text, truncation, empty and partial exports, file errors).

## Development helpers (Linux/Mono, no NuGet)
`tools/run-tests.sh` compiles the engine and tests with Mono's `mcs` against the Dataverse SDK DLLs and runs them through `tools/MiniMsTest` (a tiny MSTest-compatible runner).
`tools/UiSmoke` builds the real `MainControl` under Mono WinForms/Xvfb against a **stub** of the host API and drives its binding logic. These are local aids; the Windows
commands above are authoritative.
