param([string]$Compiler, [switch]$CompileOnly)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
Push-Location $repo
try {
    if (!$Compiler) {
        $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
        if (Test-Path $vswhere) {
            $vs = & $vswhere -latest -products '*' -property installationPath
            if ($vs) { $Compiler = Join-Path $vs 'MSBuild\Current\Bin\Roslyn\csc.exe' }
        }
    }
    if (!$Compiler -or !(Test-Path $Compiler)) { throw 'Supply a Roslyn csc.exe path (C# 7.3+).' }
    $Compiler = (Resolve-Path $Compiler).Path
    $dependency = Join-Path $repo 'artifacts/refs/newtonsoft/lib/net45/Newtonsoft.Json.dll'
    if (!(Test-Path $dependency)) {
        New-Item -ItemType Directory -Force artifacts/refs | Out-Null
        Invoke-WebRequest 'https://www.nuget.org/api/v2/package/Newtonsoft.Json/13.0.3' -OutFile artifacts/refs/newtonsoft.zip
        Expand-Archive artifacts/refs/newtonsoft.zip artifacts/refs/newtonsoft -Force
    }
    Copy-Item $dependency artifacts/Newtonsoft.Json.dll -Force
    & $Compiler /nologo /langversion:7.3 /target:library /out:artifacts\Chart.Domain.dll "/r:$dependency" src\Chart.Domain\ChartDocument.cs src\Chart.Domain\ChartCodec.cs src\Chart.Domain\S0Migration.cs apps\powerpoint-vsto\Contracts.cs apps\powerpoint-vsto\MetadataCodec.cs
    if ($LASTEXITCODE) { throw 'Domain library compilation failed.' }
    & $Compiler /nologo /langversion:7.3 /target:exe /out:artifacts\DomainTests.exe "/r:$dependency" /r:artifacts\Chart.Domain.dll tests\s1\DomainTests.cs
    if ($LASTEXITCODE) { throw 'Domain tests compilation failed.' }
    if ($CompileOnly) { Write-Output 'COMPILE ONLY: domain library and tests compiled; no execution.'; return }
    & .\artifacts\DomainTests.exe tests/s1/fixtures/clustered-column.v2.json artifacts/domain-roundtrip.v2.json
    if ($LASTEXITCODE) { throw 'Domain tests failed.' }
} finally { Pop-Location }
