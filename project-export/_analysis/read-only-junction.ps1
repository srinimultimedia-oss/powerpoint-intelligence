$ErrorActionPreference='Stop'
$repo=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
. (Join-Path $repo 'scripts/Artifact-Inputs.ps1')
. (Join-Path $repo 'scripts/Source-Fingerprint.ps1')
$fixture=Join-Path $repo 'artifacts/validation-20261103-120641-25027ceb/source/artifacts/checkmarx-113558/55cda39296314ab49e34d01b666a5251'
$valid=Join-Path $fixture 'evidence.json';$linked=Join-Path $fixture 'linked/evidence.json'
if(!(Test-Path -LiteralPath $valid)){throw 'Existing positive fixture unavailable; no fixture was created.'}
$text=Read-PpiArtifactJsonText $valid
if($text -ne '{"ReadOnly":true}'){throw 'Positive artifact fixture changed; no unverified control accepted.'}
Write-Output 'PASS existing ordinary JSON artifact read positive control'
$failed=$false;try{Read-PpiArtifactJsonText $linked | Out-Null}catch{if($_.Exception.Message -ne 'Reparse paths are not recovery inputs.'){throw};$failed=$true}
if(!$failed){throw 'Junction JSON was accepted'}
Write-Output 'PASS artifact JSON rejects existing junction before content read'
$failed=$false;try{Get-PpiFingerprintFileHash $fixture $linked | Out-Null}catch{if($_.Exception.Message -ne 'Fingerprint input cannot traverse a reparse point.'){throw};$failed=$true}
if(!$failed){throw 'Junction fingerprint was accepted'}
Write-Output 'PASS fingerprint rejects existing junction before hashing'
Write-Output ('UTC '+[DateTime]::UtcNow.ToString('o')+'; 3 read-only policy assertions passed; no fixtures created or changed')
