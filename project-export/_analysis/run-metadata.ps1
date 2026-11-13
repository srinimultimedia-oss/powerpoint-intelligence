$ErrorActionPreference='Stop'
$repo=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
. (Join-Path $repo 'scripts/Build-Integrity.ps1')
$vswhere=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
Assert-PpiMicrosoftBuildTool $vswhere 'vswhere.exe'
$vs=& $vswhere -latest -products '*' -property installationPath
$csc=Join-Path $vs 'MSBuild/Current/Bin/Roslyn/csc.exe'
Assert-PpiMicrosoftBuildTool $csc 'csc.exe'
$out=Join-Path $repo 'project-export/_analysis/runtime-temp/MetadataDocumentationTests.exe'
$priorTmp=$env:TMP;$priorTemp=$env:TEMP
try{
 $env:TMP=$env:TEMP=Split-Path $out -Parent
 & $csc /nologo /target:exe "/out:$out" (Join-Path $repo 'apps/powerpoint-vsto/Contracts.cs') (Join-Path $repo 'apps/powerpoint-vsto/MetadataCodec.cs') (Join-Path $repo 'tests/s0/MetadataTests.cs') (Join-Path $repo 'src/Library.Core/PathPolicy.cs')
 if($LASTEXITCODE){throw 'Export-local metadata test compilation failed.'}
 & $out
 if($LASTEXITCODE){throw 'Export-local metadata tests failed.'}
}finally{$env:TMP=$priorTmp;$env:TEMP=$priorTemp}
