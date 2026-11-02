param([ValidateSet('x86','x64')][string]$Platform='x64')
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'Package-Content.ps1')
$project=Join-Path $repo 'apps/powerpoint-vsto/PowerPointIntelligence.csproj'
$assembly=Join-Path $repo "apps/powerpoint-vsto/obj/$Platform/Release/PowerPointIntelligence.dll"
$vswhere=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
$vsPath=& $vswhere -latest -products '*' -property installationPath
$msbuild=Join-Path $vsPath 'MSBuild/Current/Bin/MSBuild.exe'
$log=Join-Path $repo "artifacts/rollback-compile-$Platform.log"
if(!(Test-Path -LiteralPath $assembly)){throw 'Compile the platform once before this rollback regression.'}
$source=[IO.File]::ReadAllText((Join-Path $repo 'apps/powerpoint-vsto/DemoRibbon.cs'))
$marker=[regex]::Match($source,'BuildMarker = "([^"]+)"').Groups[1].Value
if(!$marker){throw 'Source build identity missing.'}
$originalTime=(Get-Item -LiteralPath $assembly).LastWriteTimeUtc
$future=[DateTime]::UtcNow.AddDays(1)
$passed=$false
try {
    # Model a restored source snapshot with a newer cached compiler output.
    (Get-Item -LiteralPath $assembly).LastWriteTimeUtc=$future
    & $msbuild $project /t:Compile /p:Configuration=Release "/p:Platform=$Platform" /p:UseSharedCompilation=false /v:minimal /nologo *> $log
    if($LASTEXITCODE){throw 'Baseline compile failed; inspect the regression log.'}
    if((Get-Item -LiteralPath $assembly).LastWriteTimeUtc -lt $future.AddSeconds(-1)){throw 'Regression fixture did not reproduce incremental reuse.'}
    Write-Output 'PASS reproduced cached output newer than source; ordinary Compile reused it'
    & $msbuild $project /t:Compile /p:Configuration=Release "/p:Platform=$Platform" /p:PpiForcePayloadCompile=true /p:UseSharedCompilation=false /v:minimal /nologo *>> $log
    if($LASTEXITCODE){throw 'Forced compile failed; inspect the regression log.'}
    if((Get-Item -LiteralPath $assembly).LastWriteTimeUtc -ge $future.AddSeconds(-1)){throw 'Forced compilation retained the cached output.'}
    if(!(Test-PpiPayloadMarker -Bytes ([IO.File]::ReadAllBytes($assembly)) -Marker $marker)){throw 'Compiled payload identity does not match source.'}
    Write-Output "PASS forced compile replaced the cached output and contains $marker"
    $passed=$true
} finally {
    if(!$passed -and (Test-Path -LiteralPath $assembly)){(Get-Item -LiteralPath $assembly).LastWriteTimeUtc=$originalTime}
}
Write-Output 'No package signing, Office activation, VSTO Clean or registration change was requested.'
