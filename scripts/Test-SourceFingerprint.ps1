$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'Source-Fingerprint.ps1')
$fixture=Join-Path $repo ('artifacts/source-fingerprint-tests/'+[Guid]::NewGuid().ToString('N'))
$count=0
function Check([bool]$Condition,[string]$Name){if(!$Condition){throw "FAIL source fingerprint: $Name"};$script:count++}
function Digest([string]$Text){$sha=[Security.Cryptography.SHA256]::Create();try{[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text))).Replace('-','')}finally{$sha.Dispose()}}
foreach($folder in @('apps/powerpoint-vsto','src/Chart.Domain','src/Chart.Renderers','src/Shared.Contracts','src/Library.Core','scripts')) {
    $null=New-Item -ItemType Directory -Force -Path (Join-Path $fixture $folder)
}
$marker=Join-Path $fixture 'must-not-exist.txt'
$probe="Set-Content -LiteralPath '"+$marker.Replace("'","''")+"' -Value 'executed'; throw 'Fingerprint must not execute source'"
$contents=[ordered]@{
    'apps\powerpoint-vsto\Feature.cs'='// data read only'
    'src\Chart.Domain\Model.cs'='// model fixture'
    'src\Library.Core\Probe.ps1'=$probe
}
foreach($relative in $contents.Keys){[IO.File]::WriteAllText((Join-Path $fixture $relative),$contents[$relative],[Text.Encoding]::UTF8)}
# WriteAllText may use a BOM on some framework hosts. Compute the expected
# content hashes from known written bytes, independently of file selection.
$records=@($contents.Keys | Sort-Object | ForEach-Object {$_+'='+(Get-FileHash -LiteralPath (Join-Path $fixture $_) -Algorithm SHA256).Hash})
$expected=Digest ($records -join "`n")
$addin=Get-PpiAddinFingerprint $fixture
Check ($addin -eq $expected) 'exact digest of known local file contents and relative names'
Check (!(Test-Path -LiteralPath $marker)) 'enumerated script body was not executed'
Check ($addin -match '\A[0-9A-F]{64}\z') 'only SHA256 digest returned'
Check ((Get-PpiAddinFingerprint $fixture) -eq $addin) 'deterministic unchanged inventory'
$null=New-Item -ItemType Directory -Force -Path (Join-Path $fixture 'apps/powerpoint-vsto/bin')
[IO.File]::WriteAllText((Join-Path $fixture 'apps/powerpoint-vsto/bin/ignored.cs'),'ignored output')
Check ((Get-PpiAddinFingerprint $fixture) -eq $addin) 'generated bin output excluded'
$all=Get-PpiSourceFingerprint $fixture
Check (!(Test-Path -LiteralPath $marker)) 'general source fingerprint also treats scripts as bytes'
[IO.File]::AppendAllText((Join-Path $fixture 'apps/powerpoint-vsto/Feature.cs'),' changed')
Check ((Get-PpiAddinFingerprint $fixture) -ne $addin) 'source mutation changes addin digest'
Check ((Get-PpiSourceFingerprint $fixture) -ne $all) 'source mutation changes general digest'
Check (!(Test-Path -LiteralPath $marker)) 'mutation detection invokes no selected source code'
Write-Output "PASS $count source-fingerprint assertions; inert source read as bytes, no download or selected-file execution."
