$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'Piidx-Recovery.ps1')
$root=Join-Path $repo ('artifacts/piidx-recovery-tests/'+[Guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Force $root
$receipt=Join-Path $root 'prior.json';$image=Join-Path $root 'Missing.dll';$scanner=Join-Path $root 'scanner.json'
$sources=@([pscustomobject]@{File='src/Library.Core/LibraryModels.cs';SHA256=('A'*64)})
$prior=@{InputSHA256=('B'*64);AssemblySHA256=('C'*64);Sources=$sources}
$prior | ConvertTo-Json -Depth 4 | Set-Content $receipt
function Evidence($Time,$Detections,$Errors){@{Time=$Time;ReadOnly=$true;McAfeeDetections=@($Detections);Errors=@($Errors)} | ConvertTo-Json -Depth 5 | Set-Content $scanner}
$count=0
function Check([bool]$Value){if(!$Value){throw 'PIIDX recovery assertion failed.'};$script:count++}
function Reject([scriptblock]$Action){$failed=$false;try{& $Action | Out-Null}catch{$failed=$true};Check $failed}
Evidence ([DateTimeOffset]::UtcNow.ToString('o')) @() @()
$before=[IO.File]::ReadAllText($receipt)
$review=Start-PpiPiidxSourceRecovery $receipt $image $scanner $sources
$audit=Get-Content $review.Audit -Raw | ConvertFrom-Json
Check ($audit.PriorReceipt -ceq $before)
Check ([IO.File]::ReadAllText($receipt) -ceq $before)
Check (!(Test-Path -LiteralPath $image) -and !$audit.Executed -and !$audit.RestoredOldBinary)
Evidence ([DateTimeOffset]::UtcNow.ToString('o')) @(@{SHA256=('C'*64);Target='unrelated'}) @()
Reject {Start-PpiPiidxSourceRecovery $receipt $image $scanner $sources}
Evidence ([DateTimeOffset]::UtcNow.ToString('o')) @(@{SHA256=('D'*64);Target=$image}) @()
Reject {Start-PpiPiidxSourceRecovery $receipt $image $scanner $sources}
Evidence ([DateTimeOffset]::UtcNow.ToString('o')) @() @('inaccessible')
Reject {Start-PpiPiidxSourceRecovery $receipt $image $scanner $sources}
Evidence ([DateTimeOffset]::UtcNow.AddDays(-2).ToString('o')) @() @()
Reject {Start-PpiPiidxSourceRecovery $receipt $image $scanner $sources}
Evidence ([DateTimeOffset]::UtcNow.ToString('o')) @() @()
$prior.Sources=@(@{File='../outside.cs';SHA256=('A'*64)});$prior | ConvertTo-Json -Depth 4 | Set-Content $receipt
Reject {Start-PpiPiidxSourceRecovery $receipt $image $scanner $sources}
[IO.File]::WriteAllBytes($image,[byte[]](1,2,3))
Reject {Start-PpiPiidxSourceRecovery $receipt $image $scanner $sources}
Check ([IO.File]::ReadAllBytes($image).Length -eq 3)
Write-Output "PASS $count PIIDX recovery assertions; inert receipts only, no compiler or recovered code executed."
