$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'Scene-RunnerCache.ps1')
$root=Join-Path ([IO.Path]::GetTempPath()) ('ppi-cache-policy-'+[Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $root | Out-Null
$receipt=Join-Path $root 'receipt.json'
$files=[ordered]@{Executable=(Join-Path $root 'runner.bin');Library=(Join-Path $root 'library.bin');Dependency=(Join-Path $root 'dependency.bin')}
$passed=0
function Check([bool]$Condition,[string]$Name){if(!$Condition){throw "FAIL $Name"};$script:passed++;Write-Output "PASS $Name"}
function Reject([scriptblock]$Action,[string]$Name){$rejected=$false;try{& $Action | Out-Null}catch{$rejected=$true};Check $rejected $Name}
Check ((Get-PpiSceneCacheState $receipt 'key' $files) -eq 'Build') 'Empty cache permits initial build'
[IO.File]::WriteAllText($files.Executable,'inert fixture; never executed')
Reject {Get-PpiSceneCacheState $receipt 'key' $files} 'Existing unreceipted runner is not overwritten'
foreach($name in $files.Keys){[IO.File]::WriteAllText($files[$name],"inert $name")}
$hashes=[ordered]@{};foreach($name in $files.Keys){$hashes[$name]=(Get-FileHash $files[$name] -Algorithm SHA256).Hash}
[pscustomobject]@{BuildKey='key';Hashes=$hashes} | ConvertTo-Json | Set-Content $receipt
$time=(Get-Item $files.Executable).LastWriteTimeUtc
Check ((Get-PpiSceneCacheState $receipt 'key' $files) -eq 'Reuse') 'Matching runner reused'
Check ((Get-PpiSceneCacheState $receipt 'key' $files) -eq 'Reuse') 'Repeated run reuses same file'
Check ((Get-Item $files.Executable).LastWriteTimeUtc -eq $time) 'Reuse leaves runner unchanged'
Reject {Get-PpiSceneCacheState $receipt 'different-source' $files} 'Changed build inputs fail closed'
[IO.File]::WriteAllText($files.Library,'changed')
Reject {Get-PpiSceneCacheState $receipt 'key' $files} 'Changed dependency hash fails closed'
[IO.File]::WriteAllText($files.Library,'inert Library')
Remove-Item -LiteralPath $files.Executable
Reject {Get-PpiSceneCacheState $receipt 'key' $files} 'Missing runner is not restored'
Check (!(Test-Path $files.Executable)) 'Missing file remains missing'
[pscustomobject]@{ExecutableSha256='legacy'} | ConvertTo-Json | Set-Content $receipt
Reject {Get-PpiSceneCacheState $receipt 'key' $files} 'Legacy receipt requires review, no implicit adoption'
Write-Output "$passed cache policy tests passed; no executable compiled or run. Fixtures: $root"

# Inert fixtures exercise explicit rebuild review without creating executable files.
foreach($name in $files.Keys){[IO.File]::WriteAllText($files[$name],"inert $name")}
$evidence=Join-Path $root 'scanner.json'
[pscustomobject]@{Errors=@();McAfeeDetections=@()} | ConvertTo-Json -Depth 5 | Set-Content $evidence
Check ((Start-PpiReviewedRunnerBuild $receipt $files $evidence) -eq 'Build') 'Explicit reviewed source rebuild retains stable files'
Check ((Get-Content $receipt -Raw) -match 'legacy') 'Review does not forge or replace old receipt'
$auditFile=Get-ChildItem -LiteralPath $root -Filter 'receipt.json.review-*.json' | Sort-Object Name -Descending | Select-Object -First 1
$auditText=[IO.File]::ReadAllText($auditFile.FullName)
$auditRecord=ConvertFrom-Json -InputObject $auditText
Check ($auditRecord.PriorReceipt -is [string] -and $auditRecord.PriorReceipt -eq [IO.File]::ReadAllText($receipt)) 'Audit preserves prior receipt as exact plain text'
Check ($auditText.Length -lt 8192 -and $auditText -notmatch 'PSProvider|PSDrive|PSParentPath') 'Audit excludes provider graphs and remains bounded'
Check (@($auditRecord.Before.PSObject.Properties).Count -eq 3 -and $auditRecord.Before.Executable.SHA256 -eq (Get-FileHash $files.Executable).Hash) 'Flat audit retains all original file hashes'
Check ($auditRecord.Executed -eq $false) 'Audit never claims runner execution'

[pscustomobject]@{Errors=@();McAfeeDetections=@(@{Target=('\\?\'+$files.Executable);SHA256='other-hash'})} | ConvertTo-Json -Depth 5 | Set-Content $evidence
Reject {Start-PpiReviewedRunnerBuild $receipt $files $evidence} 'Detected path cannot be rebuilt to evade scanner'
[pscustomobject]@{Errors=@();McAfeeDetections=@(@{Target='other-path';SHA256=(Get-FileHash $files.Executable).Hash})} | ConvertTo-Json -Depth 5 | Set-Content $evidence
Reject {Start-PpiReviewedRunnerBuild $receipt $files $evidence} 'Detected hash cannot be rebuilt at another path'
[pscustomobject]@{Errors=@('denied');McAfeeDetections=@()} | ConvertTo-Json -Depth 5 | Set-Content $evidence
Reject {Start-PpiReviewedRunnerBuild $receipt $files $evidence} 'Incomplete scanner review blocks rebuild'
Write-Output "$passed total cache policy assertions passed. All review fixtures are inert."

# Recovery is explicit, audited, and does not materialize a binary itself.
[pscustomobject]@{Errors=@();McAfeeDetections=@()} | ConvertTo-Json -Depth 5 | Set-Content $evidence
$priorHashes=@{};foreach($name in $files.Keys){$priorHashes[$name]=(Get-FileHash $files[$name]).Hash}
[pscustomobject]@{BuildKey='reviewed-inputs';Hashes=$priorHashes} | ConvertTo-Json -Depth 4 | Set-Content $receipt
Remove-Item -LiteralPath $files.Executable
Reject {Start-PpiReviewedRunnerBuild $receipt $files $evidence} 'Default review never recovers missing files'
Check ((Start-PpiReviewedRunnerBuild $receipt $files $evidence -RecoverMissing) -eq 'Build') 'Explicit recovery validates prior hashed receipt without restoring file'
Check (!(Test-Path $files.Executable)) 'Recovery review itself creates no executable'
[pscustomobject]@{Errors=@();McAfeeDetections=@(@{Target='other-path';SHA256=$priorHashes.Executable})} | ConvertTo-Json -Depth 5 | Set-Content $evidence
Reject {Start-PpiReviewedRunnerBuild $receipt $files $evidence -RecoverMissing} 'Recovery refuses a previously detected binary hash'
Write-Output "$passed cache policy assertions passed including explicit missing-file recovery; all fixtures inert."
# Explicit legacy recovery uses inert fixtures; no executable is compiled or launched.
$legacyRoot=Join-Path $root 'legacy-review';New-Item -ItemType Directory $legacyRoot | Out-Null
$legacyReceipt=Join-Path $legacyRoot 'compiled.json'
$legacyFiles=[ordered]@{Executable=(Join-Path $legacyRoot 'runner.bin');Library=(Join-Path $legacyRoot 'library.bin');Dependency=(Join-Path $legacyRoot 'dependency.bin')}
$legacyEvidence=Join-Path $legacyRoot 'scanner.json'
$legacy=[pscustomobject]@{Platform='x64';TargetFramework='.NET Framework 4.8';ExecutableSha256=('1'*64);LibrarySha256=('2'*64)}
[IO.File]::WriteAllText($legacyReceipt,($legacy|ConvertTo-Json));$legacyText=[IO.File]::ReadAllText($legacyReceipt)
[pscustomobject]@{ReadOnly=$true;Errors=@();McAfeeDetections=@()}|ConvertTo-Json|Set-Content $legacyEvidence
$auditPath=Start-PpiLegacySourceReview $legacyReceipt $legacyFiles $legacyEvidence ('3'*64) $legacyRoot
$legacyAudit=ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($auditPath))
Check ($legacyAudit.PriorReceipt -ceq $legacyText) 'Legacy review preserves exact prior receipt in plain audit'
Check ([IO.File]::ReadAllText($legacyReceipt) -ceq $legacyText) 'Legacy review does not overwrite prior receipt'
Check (@($legacyFiles.Values|Where-Object {Test-Path -LiteralPath $_}).Count -eq 0) 'Legacy review creates no runner outputs'
Check ($legacyAudit.CurrentBuildKey -eq ('3'*64) -and $legacyAudit.Executed -eq $false) 'Legacy audit binds current source/toolchain without claiming execution'
Reject {Start-PpiLegacySourceReview $legacyReceipt $legacyFiles $legacyEvidence 'bad-key' $legacyRoot} 'Legacy review requires bounded current build provenance'
[pscustomobject]@{ReadOnly=$true;Errors=@();McAfeeDetections=@(@{Target='old-location';SHA256=('1'*64)})}|ConvertTo-Json -Depth 4|Set-Content $legacyEvidence
Reject {Start-PpiLegacySourceReview $legacyReceipt $legacyFiles $legacyEvidence ('3'*64) $legacyRoot} 'Legacy recovery refuses a historically detected prior hash'
[pscustomobject]@{ReadOnly=$true;Errors=@();McAfeeDetections=@(@{Target='old-location';SHA256=('4'*64)})}|ConvertTo-Json -Depth 4|Set-Content $legacyEvidence
$reviewedReport=Read-PpiReviewedScannerEvidence $legacyEvidence
Reject {Assert-PpiRunnerScannerClear $reviewedReport $legacyFiles ([ordered]@{Executable=('4'*64)})} 'Post-build guard refuses a detected new hash independent of path'
[pscustomobject]@{ReadOnly=$true;Errors=@();McAfeeDetections=@(@{Target=('\\?\'+$legacyFiles.Executable);SHA256=('5'*64)})}|ConvertTo-Json -Depth 4|Set-Content $legacyEvidence
$reviewedReport=Read-PpiReviewedScannerEvidence $legacyEvidence
Reject {Assert-PpiRunnerScannerClear $reviewedReport $legacyFiles ([ordered]@{Executable=('6'*64)})} 'Post-build guard refuses a detected Windows-prefixed path'
[pscustomobject]@{ReadOnly=$true;Errors=@('unreadable');McAfeeDetections=@()}|ConvertTo-Json|Set-Content $legacyEvidence
Reject {Read-PpiReviewedScannerEvidence $legacyEvidence} 'Legacy recovery refuses incomplete scanner evidence'
[pscustomobject]@{ReadOnly=$true;Errors=@();McAfeeDetections=@()}|ConvertTo-Json|Set-Content $legacyEvidence
[IO.File]::WriteAllText($legacyFiles.Executable,'inert; never executed')
Reject {Start-PpiLegacySourceReview $legacyReceipt $legacyFiles $legacyEvidence ('3'*64) $legacyRoot} 'Legacy recovery never adopts or overwrites an existing binary'
Write-Output "$passed total cache/review assertions passed; all fixtures inert."
