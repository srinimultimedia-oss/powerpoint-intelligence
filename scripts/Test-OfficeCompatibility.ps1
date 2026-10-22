param([string]$Compiler, [ValidateSet('AnyCPU', 'x86', 'x64')][string]$Platform = 'AnyCPU', [switch]$CompileOnly)
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
    if (!$Compiler -or !(Test-Path $Compiler)) { throw 'Supply a Roslyn csc.exe path.' }
    $Compiler = (Resolve-Path $Compiler).Path
    $output = "artifacts/compatibility/$Platform"
    New-Item -ItemType Directory -Force $output | Out-Null
    & $Compiler /nologo /langversion:7.3 "/platform:$Platform" /target:exe "/out:$output/OfficeCompatibilityTests.exe" /r:System.Configuration.dll apps\powerpoint-vsto\OfficeCompatibility.cs tests\compatibility\OfficeCompatibilityTests.cs
    if ($LASTEXITCODE) { throw 'Office compatibility test compilation failed.' }
    if ($CompileOnly) { Write-Output "COMPILE ONLY: $Platform compatibility tests compiled; no execution."; return }
    & "$output/OfficeCompatibilityTests.exe" apps/powerpoint-vsto/App.config $Platform
    if ($LASTEXITCODE) { throw 'Office compatibility tests failed.' }
} finally { Pop-Location }
