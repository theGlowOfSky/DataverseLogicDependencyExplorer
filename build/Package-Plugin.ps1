# ----------------------------------------------------------------------------
# Dataverse Logic Dependency Explorer  v1.2026.11.11
# Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
# Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
# ----------------------------------------------------------------------------
<#
  Builds, tests, verifies and packages the plugin. Run on Windows from any folder:
      powershell -ExecutionPolicy Bypass -File build\Package-Plugin.ps1
  Requires: .NET SDK (VS 2022 / .NET 8 SDK), .NET Framework 4.8 targeting pack. nuget.exe on PATH is optional (needed for the .nupkg).
  Produces (under artifacts\):
      package\Plugins\DataverseLogicDependencyExplorer.dll        <- copy this one file into %APPDATA%\MscrmTools\XrmToolBox\Plugins
      DataverseLogicDependencyExplorer-<version>-deployment.zip
      DataverseLogicDependencyExplorer.<version>.nupkg            (when nuget.exe is available)
  Nothing is published.
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$name = 'DataverseLogicDependencyExplorer'
Set-Location $root

dotnet restore "$name.sln"
dotnet build   "$name.sln" -c Release --no-restore
dotnet test    "tests\$name.Tests\$name.Tests.csproj" -c Release --no-build

$out = Join-Path $root "src\$name\bin\Release"
$dll = Join-Path $out "$name.dll"
if (-not (Test-Path $dll)) { throw "Missing $dll" }

# --- guard against the CompositionContractMismatchException class of failure: nothing host-provided may be deployed ---
$forbidden = 'XrmToolBox*','McTools*','Microsoft.Xrm*','Microsoft.Crm*','Newtonsoft.Json*','Microsoft.IdentityModel*','Microsoft.Identity.Client*'
$extra = Get-ChildItem $out -File -Filter *.dll | Where-Object { $_.Name -ne "$name.dll" }
$bad   = $extra | Where-Object { $n = $_.Name; $forbidden | Where-Object { $n -like $_ } }
if ($bad) { throw ("Host assemblies found in the build output (must not be deployed): " + ($bad.Name -join ', ')) }
if ($extra) { Write-Warning ("Unexpected extra DLLs in the build output: " + ($extra.Name -join ', ')) }

# --- the assembly version must match the package version ---
$asmVersion = [System.Reflection.AssemblyName]::GetAssemblyName($dll).Version
$specVersion = [version](([xml](Get-Content "build\$name.nuspec" -Raw)).package.metadata.version)
if ($asmVersion.ToString() -ne $specVersion.ToString()) { throw "Assembly version $asmVersion does not match nuspec version $specVersion" }

# --- stage ---
$stage = Join-Path $root 'artifacts\package'
Remove-Item $stage -Recurse -Force -ErrorAction SilentlyContinue
$plugins = Join-Path $stage 'Plugins'
New-Item -ItemType Directory -Force $plugins | Out-Null
Copy-Item $dll $plugins
Copy-Item (Join-Path $root 'assets\icon-256.png') (Join-Path $stage 'icon.png')
Copy-Item (Join-Path $root 'README.md') $stage

$zip = Join-Path $root "artifacts\$name-$($specVersion.ToString())-deployment.zip"
Remove-Item $zip -ErrorAction SilentlyContinue
Compress-Archive -Path (Join-Path $stage 'Plugins') -DestinationPath $zip
Write-Host "Deployment package: $zip"

if (Get-Command nuget -ErrorAction SilentlyContinue) {
    nuget pack (Join-Path $root "build\$name.nuspec") -BasePath $stage -OutputDirectory (Join-Path $root 'artifacts')
} else {
    Write-Warning 'nuget.exe not found on PATH: the .nupkg was not created.'
}
Write-Host "Staged at $stage"
