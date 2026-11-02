param([Parameter(Mandatory=$true)][string]$Compiler)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'Common.ps1')
$repo=Split-Path $PSScriptRoot -Parent
$root=Join-Path $repo ('artifacts/easySetup-tests/resource-reader-'+[Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force (Join-Path $root 'first'),(Join-Path $root 'second'),(Join-Path $root 'missing') | Out-Null
$source=Join-Path $root 'Inert.cs'
[IO.File]::WriteAllText($source,'public sealed class InertResourceFixture {}')
foreach($kind in @('first','second','missing')){
    $fixture=Join-Path $root "$kind/InertResourceFixture.dll"
    $data=Join-Path $root "$kind/data.txt"
    [IO.File]::WriteAllText($data,$kind)
    $resource=if($kind -eq 'missing'){'OtherResource'}else{'IntelligencePayload.zip'}
    & $Compiler /nologo /target:library "/out:$fixture" "/resource:$data,$resource" $source
    if($LASTEXITCODE){throw 'Inert resource fixture compilation failed.'}
}
$public=Join-Path $root 'Current.dll'
Copy-Item (Join-Path $root 'first/InertResourceFixture.dll') $public
$cached=[Reflection.Assembly]::LoadFile($public)
$stream=$cached.GetManifestResourceStream('IntelligencePayload.zip')
$sha=[Security.Cryptography.SHA256]::Create()
try{$firstHash=[BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-','')}finally{$sha.Dispose();$stream.Dispose()}
$replacement=Join-Path $root 'Next.dll'
Copy-Item (Join-Path $root 'second/InertResourceFixture.dll') $replacement
[IO.File]::Replace($replacement,$public,(Join-Path $root 'Prior.dll'))
$expected=(Get-FileHash (Join-Path $root 'second/data.txt') -Algorithm SHA256).Hash
$actual=Get-EasySetupEmbeddedPayloadHash ([IO.File]::ReadAllBytes($public))
if($actual -ne $expected -or $actual -eq $firstHash){throw 'Resource reader reused the previous path-loaded assembly.'}
# Repeated same-session validation must still read the current payload.
if((Get-EasySetupEmbeddedPayloadHash ([IO.File]::ReadAllBytes($public))) -ne $expected){throw 'Repeated resource validation returned stale bytes.'}
$rejected=$false
try{Get-EasySetupEmbeddedPayloadHash ([IO.File]::ReadAllBytes((Join-Path $root 'missing/InertResourceFixture.dll'))) | Out-Null}catch{if($_.Exception.Message -eq 'Embedded payload resource missing.'){$rejected=$true}else{throw}}
if(!$rejected){throw 'Missing embedded resource was accepted.'}
Write-Output 'PASS 3 resource-reader regression assertions: atomic replacement, repeated same-session validation and missing-resource refusal. Inert fixtures only; no installation.'
