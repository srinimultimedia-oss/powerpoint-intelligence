$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'Build-Integrity.ps1')
$output=Join-Path $repo 'artifacts/build-integrity-tests'
$null=New-Item -ItemType Directory -Force -Path $output
$count=0
function Check([bool]$Value,[string]$Name){if(!$Value){throw "FAIL integrity: $Name"};$script:count++}
function Reject([scriptblock]$Action,[string]$Message){$rejected=$false;try{& $Action | Out-Null}catch{if($_.Exception.Message -notlike "*$Message*"){throw};$rejected=$true};Check $rejected 'fail-closed rejection'}
$candidate=Join-Path $output 'candidate.bin';$reference=Join-Path $output 'reference.bin'
[IO.File]::WriteAllBytes($candidate,[byte[]](1,2,3));[IO.File]::WriteAllBytes($reference,[byte[]](1,2,3))
Check ((Get-PpiContentHash $candidate) -eq (Get-FileHash -LiteralPath $candidate -Algorithm SHA256).Hash) 'stream hashing preserves fingerprint bytes'
$lease=Get-PpiVerifiedArtifact $candidate $reference
try {
    Check ($lease.Bytes.Length -eq 3 -and $lease.Bytes[2] -eq 3) 'verified immutable snapshot'
    $blocked=$false;try{[IO.File]::WriteAllBytes($candidate,[byte[]](9))}catch [IO.IOException]{$blocked=$true}
    Check $blocked 'write denied while execution lease is held'
} finally {$lease.Stream.Dispose()}
[IO.File]::WriteAllBytes($candidate,[byte[]](9,9,9))
Check ($lease.Bytes[0] -eq 1) 'verified bytes survive disk replacement'
$receipt=Join-Path $output 'forged-receipt.json'
@{SHA256=(Get-PpiContentHash $candidate)} | ConvertTo-Json | Set-Content -LiteralPath $receipt
Check ((Get-PpiContentHash $candidate) -eq (Get-Content $receipt -Raw | ConvertFrom-Json).SHA256) 'attacker-controlled cache and receipt agree'
Reject {Get-PpiVerifiedArtifact $candidate $reference} 'independently rebuilt'
Reject {Get-PpiReproducedArtifacts ([ordered]@{Fixture=$candidate}) {param($targets) [IO.File]::WriteAllBytes($targets.Fixture,[byte[]](1,2,3))}} 'independently rebuilt'
foreach($thumb in @('', ' ', '../certificate', ('A'*39), ('A'*41), (('A'*40)+';CustomAfterBuild=bad'), (('A'*40)+"`n"), ('G'*40))) {
    Reject {Get-PpiApprovedSigningCertificate $thumb} 'exactly 40 hexadecimal'
}
Reject {Assert-PpiSigningCertificate $null ('A'*40)} 'selected signing certificate'
$vswhere=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
Assert-PpiMicrosoftBuildTool $vswhere 'vswhere.exe';$count++
$vs=& $vswhere -latest -products '*' -property installationPath
$compiler=Join-Path $vs 'MSBuild/Current/Bin/Roslyn/csc.exe'
Assert-PpiMicrosoftBuildTool $compiler 'csc.exe';$count++
Reject {Assert-PpiMicrosoftBuildTool $compiler 'MSBuild.exe'} 'filename'
$source=Join-Path $output 'IntegrityFixture.cs';$image=Join-Path $output 'IntegrityFixture.dll'
[IO.File]::WriteAllText($source,'public static class IntegrityFixture { public static int Value() { return 42; } }',[Text.Encoding]::ASCII)
$compileOutput=& $compiler /nologo /deterministic+ /target:library "/out:$image" $source
if($LASTEXITCODE){throw ($compileOutput -join [Environment]::NewLine)}
$expected=Join-Path $output 'expected';$null=New-Item -ItemType Directory -Force $expected
$expectedImage=Join-Path $expected 'IntegrityFixture.dll';Copy-Item -LiteralPath $image -Destination $expectedImage -Force
$verified=Get-PpiVerifiedArtifact $image $expectedImage
try {$assembly=[Reflection.Assembly]::Load([byte[]]$verified.Bytes)} finally {$verified.Stream.Dispose()}
[IO.File]::WriteAllBytes($image,[byte[]](0,1,2))
Check ($assembly.GetType('IntegrityFixture').GetMethod('Value').Invoke($null,@()) -eq 42) 'execute verified bytes, never re-read changed DLL path'
Write-Output "PASS $count build-integrity assertions; only an inert constant-return fixture loaded. No cached PIIDX/scene/template runner or signing certificate was replaced."
