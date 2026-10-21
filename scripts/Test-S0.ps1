param([string]$Compiler, [switch]$InteractiveOffice, [switch]$MetadataOnly, [switch]$CompileOnly)
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
    $office = 'C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\OFFICE.DLL'
    $ppt = 'C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.PowerPoint.dll'
    if (!(Test-Path $office) -or !(Test-Path $ppt)) { throw 'Office 15 PIAs required for adapter checks.' }
    & $Compiler /nologo /target:exe /out:artifacts\HostIntegrationTests.exe "/r:$office" "/r:$ppt" /r:Microsoft.CSharp.dll apps\powerpoint-vsto\Contracts.cs apps\powerpoint-vsto\MetadataCodec.cs apps\powerpoint-vsto\PowerPointHost.cs tests\s0\HostIntegrationTests.cs
    if ($LASTEXITCODE) { throw 'Host integration compilation failed.' }
    if ($InteractiveOffice -and !$CompileOnly) {
        $fixture = Join-Path $repo ('artifacts\office-' + [Guid]::NewGuid().ToString('N'))
        & .\artifacts\HostIntegrationTests.exe --interactive-office $fixture
        if ($LASTEXITCODE) { throw 'Interactive Office integration failed.' }
    } else { Write-Output 'SKIP interactive Office integration (requires -InteractiveOffice and desktop user session).' }
} finally { Pop-Location }
