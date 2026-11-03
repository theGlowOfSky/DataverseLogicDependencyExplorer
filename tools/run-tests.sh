#!/bin/sh
# ----------------------------------------------------------------------------
# Dataverse Logic Dependency Explorer  v1.2026.11.11
# Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
# Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
# ----------------------------------------------------------------------------
# Local (Linux/Mono) test run for environments without NuGet access. On Windows use: dotnet test
set -e
cd "$(dirname "$0")/.."
OUT=${OUT:-/tmp/build}; mkdir -p "$OUT"
REFDIR=${REFDIR:-/home/claude/refs}
SRC=src/DataverseLogicDependencyExplorer
REFS="-r:$REFDIR/Microsoft.Xrm.Sdk.dll -r:$REFDIR/Newtonsoft.Json.dll -r:System.ServiceModel -r:System.Runtime.Serialization -r:System.Xml.Linq -r:System.Core -r:System.Xml -r:System.IO.Compression"
cp "$REFDIR/Microsoft.Xrm.Sdk.dll" "$REFDIR/Newtonsoft.Json.dll" "$OUT/"
mcs -target:library -out:"$OUT/DataverseLogicDependencyExplorer.Core.dll" $REFS $SRC/Models/*.cs $SRC/Services/*.cs $SRC/Services/Scanners/*.cs $SRC/Reporting/*.cs $SRC/Resources/*.cs $SRC/ViewModels/*.cs $SRC/UiLayout.cs tools/MonoAssemblyInfo.cs
mcs -target:exe -out:"$OUT/Tests.exe" -main:Microsoft.VisualStudio.TestTools.UnitTesting.MiniRunner $REFS -r:"$OUT/DataverseLogicDependencyExplorer.Core.dll" tools/MiniMsTest/MiniMsTest.cs tests/DataverseLogicDependencyExplorer.Tests/*.cs
mono "$OUT/Tests.exe"
