$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'Diagnostic-OutputPath.ps1')
$count=0
function Check([bool]$Condition,[string]$Name) {
    if(!$Condition){throw "FAIL diagnostic output: $Name"}
    $script:count++
}
function Reject([string]$Value) {
    $rejected=$false
    try{$null=Get-PpiDiagnosticOutputDirectory $Value $repo}catch [IO.InvalidDataException]{$rejected=$true}
    Check $rejected 'unsafe output rejected before directory creation'
}
$owned=Join-Path $repo 'artifacts/diagnostic-output-tests'
Check ((Get-PpiDiagnosticOutputDirectory 'artifacts/diagnostic-output-tests' $repo) -eq $owned) 'relative artifact directory'
Check ((Get-PpiDiagnosticOutputDirectory $owned $repo) -eq $owned) 'absolute artifact directory'
Check ((Get-PpiDiagnosticOutputDirectory 'artifacts/diagnostic-output-tests/nested directory' $repo) -eq (Join-Path $owned 'nested directory')) 'spaces and nested paths'
$temporary=Join-Path ([IO.Path]::GetTempPath()) ('ppi-diagnostic-'+[guid]::NewGuid().ToString('N'))
Check ((Get-PpiDiagnosticOutputDirectory $temporary $repo) -eq $temporary) 'independent temporary root'
foreach($bad in @('','..\outside','artifacts/../outside','artifacts/.','C:relative','\root-relative','\\?\C:\extended','\\.\NUL','artifacts/file:stream','artifacts/C:','artifacts/CON','artifacts/con.txt','artifacts/LPT9','artifacts/trailing.','artifacts/trailing ','artifacts/bad*','artifacts/bad?',"artifacts/bad`nname",($owned+'-sibling\..\escape'),([IO.Path]::Combine($repo+'-sibling','outside')),[IO.Path]::GetTempPath(),$repo,(Join-Path $env:WINDIR 'ppi-diagnostic-output'))) { Reject $bad }
Reject ('artifacts/LPT'+[char]0x00b9)
Check (!(Test-Path -LiteralPath $temporary)) 'validation does not create temporary directories'
Check ((Get-PpiDiagnosticReportPath $owned) -eq (Join-Path $owned 'scene-threat-evidence.json')) 'fixed report name'
# The actual collector must fail before reading scanner history or creating output.
$rejected=$false
try{& (Join-Path $PSScriptRoot 'Collect-SceneThreatEvidence.ps1') -OutputDirectory '..\outside' | Out-Null}catch [IO.InvalidDataException]{$rejected=$true}
Check $rejected 'actual collector rejects traversal at entry'
$junction=Join-Path $owned ('junction-'+[guid]::NewGuid().ToString('N'))
$target=Join-Path $owned 'junction-target'
$created=$false
try {
    $null=New-Item -ItemType Directory -Force -Path $target
    $null=New-Item -ItemType Junction -Path $junction -Target $target -ErrorAction Stop
    $created=$true
} catch {
    Write-Output 'SKIP junction creation: this session does not permit the fixture; no security policy changed.'
}
if($created) {
    try { Reject (Join-Path $junction 'child') }
    finally {
        # Remove only the verified fixture junction, never its target tree.
        $full=[IO.Path]::GetFullPath($junction)
        if(!$full.StartsWith([IO.Path]::GetFullPath($owned)+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase) -or
            ([IO.File]::GetAttributes($full) -band [IO.FileAttributes]::ReparsePoint) -eq 0) {throw 'Unsafe junction fixture cleanup refused.'}
        [IO.Directory]::Delete($full)
    }
}
Write-Output "PASS $count diagnostic path assertions; no scanner settings changed and no report collected."
