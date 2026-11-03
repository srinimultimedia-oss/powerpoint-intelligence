param([string]$Python)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
# Prevent a future literal from silently changing under an ANSI script decoder.
if(@([IO.File]::ReadAllBytes($PSCommandPath) | Where-Object {$_ -gt 127}).Count){throw 'Security bootstrap must remain ASCII; construct Unicode fixtures from character codes.'}
if(!$Python){$bundled=Join-Path $env:USERPROFILE '.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe';if(Test-Path $bundled){$Python=$bundled}else{$Python=(Get-Command python.exe -ErrorAction Stop).Source}}
if('PowerPointIntelligence.Tests.Security.SceneArtifactPath' -as [type]){throw 'Use a fresh PowerShell process to test the current source.'}
Add-Type -Path (Join-Path $repo 'tests/s3/SceneArtifactPath.cs')
$root=Join-Path $repo 'artifacts/security-remediation-tests'
New-Item -ItemType Directory -Force $root | Out-Null
$count=0
# Keep this bootstrap ASCII: Windows PowerShell 5.1 reads BOM-less scripts as ANSI.
# Build Unicode fixtures numerically so every host tests the same characters.
foreach($name in @('valid-scene','multiple.dots',('Unicode-'+[char]0x0394))){
    $actual=[PowerPointIntelligence.Tests.Security.SceneArtifactPath]::Resolve($root,$name)
    if($actual -ne (Join-Path $root ($name+'.v1.json'))){throw 'Valid basename changed.'};$count++
}
foreach($name in @('',' ','..','../escape','..\escape','/absolute','C:\absolute','C:relative','\\server\share','\\?\C:\extended','name:stream','trailing.','trailing ','CON','con.txt','PRN','AUX','NUL','COM1','LPT9',('LPT'+[char]0x00b9),('LPT'+[char]0x00b2),('LPT'+[char]0x00b3),('COM'+[char]0x00b9),('COM'+[char]0x00b2),('COM'+[char]0x00b3),'CON .txt',('a'*129),"bad`nname",('bad'+[char]0+'name'),'bad"name','bad|name','bad<name','bad>name')){
    $rejected=$false
    try{[PowerPointIntelligence.Tests.Security.SceneArtifactPath]::Resolve($root,$name) | Out-Null}catch{if($_.Exception.InnerException -is [IO.InvalidDataException]){$rejected=$true}else{throw}}
    if(!$rejected){$codes=($name.ToCharArray() | ForEach-Object {'U+{0:X4}' -f [int][char]$_}) -join ' ';throw ('Unsafe corpus basename accepted; fixture character codes: '+$codes)};$count++
}
$rejected=$false
try{[PowerPointIntelligence.Tests.Security.SceneArtifactPath]::ValidateNames($root,[string[]]@('same','SAME'))}catch{if($_.Exception.InnerException -is [IO.InvalidDataException]){$rejected=$true}else{throw}}
if(!$rejected){throw 'Colliding corpus names accepted.'};$count++
Write-Output "PASS $count corpus-path policy assertions against the actual C# helper; no cached runner replaced or executed."
Add-Type -Path @((Join-Path $repo 'src/Library.Core/PathPolicy.cs'),(Join-Path $repo 'tests/security/PathPolicyTests.cs'))
$null=[PowerPointIntelligence.Tests.Security.PathPolicyTests]::Run($repo)
& (Join-Path $PSScriptRoot 'Test-DiagnosticOutputPath.ps1')
& (Join-Path $PSScriptRoot 'Test-BuildIntegrity.ps1')
& (Join-Path $PSScriptRoot 'Test-PiidxRecovery.ps1')
& $Python -B (Join-Path $repo 'tests/security/test_repair_template_alignment.py')
if($LASTEXITCODE){throw 'Repair security regression tests failed.'}
