param([string]$Compiler, [switch]$InteractiveOffice, [switch]$MetadataOnly, [switch]$CompileOnly,
    [string]$OfficeInteropPath, [string]$PowerPointInteropPath)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
Push-Location $repo
try {
    New-Item -ItemType Directory -Force artifacts | Out-Null
    if (!$Compiler) {
        $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
        if (Test-Path $vswhere) {
            $vs = & $vswhere -latest -products '*' -property installationPath
            if ($vs) { $Compiler = Join-Path $vs 'MSBuild\Current\Bin\Roslyn\csc.exe' }
        }
    }
    if (!$Compiler -or !(Test-Path $Compiler)) { throw 'Supply -Compiler with a Roslyn csc.exe path (C# 7.3+). No machine changes are made.' }
    $Compiler = (Resolve-Path $Compiler).Path
    & $Compiler /nologo /target:exe /out:artifacts\MetadataTests.exe apps\powerpoint-vsto\Contracts.cs apps\powerpoint-vsto\MetadataCodec.cs tests\s0\MetadataTests.cs
    if ($LASTEXITCODE) { throw 'Metadata test compilation failed.' }
    if (!$CompileOnly) {
        & .\artifacts\MetadataTests.exe
        if ($LASTEXITCODE) { throw 'Metadata tests failed.' }
    } else { Write-Output 'COMPILE ONLY: metadata tests compiled, not executed.' }
    if ($MetadataOnly) { return }
    $piaRoot = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Shared\Visual Studio Tools for Office\PIA\Office15'
    if (!$OfficeInteropPath) {
        $candidate = Join-Path $piaRoot 'office.dll'
        if (Test-Path $candidate) { $OfficeInteropPath = $candidate }
        else { $OfficeInteropPath = Get-ChildItem (Join-Path $env:WINDIR 'assembly\GAC_MSIL\office') -Recurse -Filter office.dll -ErrorAction SilentlyContinue | Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName }
    }
    if (!$PowerPointInteropPath) {
        $candidate = Join-Path $piaRoot 'Microsoft.Office.Interop.PowerPoint.dll'
        if (Test-Path $candidate) { $PowerPointInteropPath = $candidate }
        else { $PowerPointInteropPath = Get-ChildItem (Join-Path $env:WINDIR 'assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint') -Recurse -Filter Microsoft.Office.Interop.PowerPoint.dll -ErrorAction SilentlyContinue | Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName }
    }
    if (!$OfficeInteropPath -or !$PowerPointInteropPath -or !(Test-Path $OfficeInteropPath) -or !(Test-Path $PowerPointInteropPath)) {
        throw 'Office/PowerPoint PIAs not found. Supply -OfficeInteropPath and -PowerPointInteropPath from your Office development tools.'
    }
    # Match the add-in project: embed the referenced API types rather than require a
    # particular PIA assembly to be installed alongside every Office runtime edition.
    & $Compiler /nologo /target:exe /out:artifacts\HostIntegrationTests.exe "/link:$OfficeInteropPath" "/link:$PowerPointInteropPath" /r:Microsoft.CSharp.dll apps\powerpoint-vsto\Contracts.cs apps\powerpoint-vsto\MetadataCodec.cs apps\powerpoint-vsto\PowerPointHost.cs tests\s0\HostIntegrationTests.cs
    if ($LASTEXITCODE) { throw 'Host integration compilation failed.' }
    if ($InteractiveOffice -and !$CompileOnly) {
        $fixture = Join-Path $repo ('artifacts\office-' + [Guid]::NewGuid().ToString('N'))
        & .\artifacts\HostIntegrationTests.exe --interactive-office $fixture
        if ($LASTEXITCODE) { throw 'Interactive Office integration failed.' }
    } else { Write-Output 'SKIP interactive Office integration (requires -InteractiveOffice and desktop user session).' }
} finally { Pop-Location }
