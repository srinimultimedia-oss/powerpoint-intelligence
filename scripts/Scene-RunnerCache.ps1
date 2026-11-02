function Get-PpiSceneBuildKey([string]$Root,[string]$Platform,[string]$Compiler,[string[]]$References,[string[]]$Sources,[string]$CacheKind='scene') {
    if(!$Sources){$Sources=@('src/Chart.Renderers/SceneContracts.cs','src/Chart.Renderers/SceneCodec.cs','src/Chart.Renderers/StrictSceneJson.cs','src/Shared.Contracts/BoundedJson.cs','tests/s3/SceneContractTests.cs','tests/s3/SceneArtifactPath.cs')}
    $records=@("$CacheKind-runner-cache-v1",'net48','csharp73',$Platform)
    foreach($source in $sources){$records+=($source+'='+(Get-FileHash -LiteralPath (Join-Path $Root $source) -Algorithm SHA256).Hash)}
    $records+=('compiler='+(Get-FileHash -LiteralPath $Compiler -Algorithm SHA256).Hash)
    foreach($reference in $References){$records+=([IO.Path]::GetFileName($reference)+'='+(Get-FileHash -LiteralPath $reference -Algorithm SHA256).Hash)}
    $hash=[Security.Cryptography.SHA256]::Create()
    try{return ([BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes(($records -join "`n"))))).Replace('-','')}finally{$hash.Dispose()}
}
function Get-PpiSceneCacheState([string]$ReceiptPath,[string]$BuildKey,[System.Collections.IDictionary]$Files) {
    if(!(Test-Path -LiteralPath $ReceiptPath)){
        foreach($file in $Files.Values){if(Test-Path -LiteralPath $file){throw 'Unreceipted test runner files exist. No overwrite or automatic rebuild is allowed; review provenance first.'}}
        return 'Build'
    }
    $receipt=Get-Content -LiteralPath $ReceiptPath -Raw | ConvertFrom-Json
    if(!$receipt.BuildKey -or $receipt.BuildKey -ne $BuildKey){throw 'Test runner provenance is missing or sources/toolchain changed. Existing runner retained; a reviewed rebuild is required.'}
    foreach($name in $Files.Keys){
        if(!(Test-Path -LiteralPath $Files[$name])){throw 'A cached test runner file is missing. Check scanner evidence; no automatic restoration or rebuild.'}
        $expected=$receipt.Hashes.PSObject.Properties[$name]
        if(!$expected -or $expected.Value -ne (Get-FileHash -LiteralPath $Files[$name] -Algorithm SHA256).Hash){throw 'Cached test runner hash mismatch. File retained without execution; review provenance/scanner evidence.'}
    }
    return 'Reuse'
}
function Start-PpiReviewedRunnerBuild([string]$ReceiptPath,[System.Collections.IDictionary]$Files,[string]$ScannerEvidence,[switch]$RecoverMissing) {
    if(!$ScannerEvidence -or !(Test-Path -LiteralPath $ScannerEvidence)){throw 'A scanner evidence JSON report is required for reviewed runner rebuild.'}
    Write-Host "REVIEW read scanner evidence"
    $report=Get-Content -LiteralPath $ScannerEvidence -Raw | ConvertFrom-Json
    if($report.Errors.Count -gt 0){throw 'Scanner evidence contains read errors; review it before rebuilding.'}
    $before=[ordered]@{}
    foreach($name in $Files.Keys){
        $file=$Files[$name]
        $missing=!(Test-Path -LiteralPath $file)
        if($missing){
            if(!$RecoverMissing){throw 'Runner file missing; default review does not restore missing files. Explicit recovery requires a prior hashed receipt and scanner evidence.'}
            if(!(Test-Path -LiteralPath $ReceiptPath)){throw 'Missing-runner recovery requires a prior receipt.'}
            $prior=ConvertFrom-Json -InputObject ([IO.File]::ReadAllText([IO.Path]::GetFullPath($ReceiptPath)))
            if(!$prior.BuildKey -or !$prior.Hashes -or !$prior.Hashes.PSObject.Properties[$name]){throw 'Missing-runner recovery requires prior build inputs and this file hash; legacy/unreceipted files cannot be recovered.'}
            $hash=[string]$prior.Hashes.PSObject.Properties[$name].Value
            if($hash -notmatch '^[0-9a-fA-F]{64}$'){throw 'Invalid prior hash; recovery refused.'}
            Write-Host "REVIEW missing $name using prior hash; no executable restored"
        }else{Write-Host "REVIEW hash $name";$hash=(Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash}
        foreach($entry in $report.McAfeeDetections){
            $target=([string]$entry.Target).Replace('\\?\','')
            if($entry.SHA256 -eq $hash -or $target -eq [IO.Path]::GetFullPath($file)){throw 'Scanner records identify this runner hash/path. Vendor clearance is required before rebuilding or executing it; no file was replaced.'}
        }
        $before[$name]=[pscustomobject]@{Path=[IO.Path]::GetFullPath($file);SHA256=$hash;Missing=$missing}
    }
    $audit=$ReceiptPath+'.review-'+(Get-Date -Format 'yyyyMMdd-HHmmss-fff')+'.json'
    Write-Host "REVIEW read plain prior receipt"
    # Get-Content attaches PSDrive/PSProvider ETS properties to strings. Windows
    # PowerShell 5.1 ConvertTo-Json traverses them; PowerShell >=7.2 ignores them.
    # ReadAllText produces plain text without the provider object graph.
    $priorReceipt=$null
    if(Test-Path -LiteralPath $ReceiptPath){
        if((Get-Item -LiteralPath $ReceiptPath).Length -gt 1048576){throw 'Prior runner receipt exceeds the 1 MB review limit.'}
        $priorReceipt=[IO.File]::ReadAllText([IO.Path]::GetFullPath($ReceiptPath))
    }
    $evidencePath=[IO.Path]::GetFullPath($ScannerEvidence)
    $evidenceHash=(Get-FileHash -LiteralPath $evidencePath -Algorithm SHA256).Hash
    $record=[pscustomobject]@{Reason='Explicit reviewed source rebuild at existing stable paths';Time=([DateTime]::UtcNow.ToString('o'));ScannerEvidence=$evidencePath;ScannerEvidenceSha256=$evidenceHash;PriorReceipt=$priorReceipt;Before=$before;Executed=$false}
    Write-Host "REVIEW serialize plain audit fields"
    $json=ConvertTo-Json -InputObject $record -Depth 5
    Write-Host "REVIEW write audit receipt"
    [IO.File]::WriteAllText([IO.Path]::GetFullPath($audit),$json,[Text.UTF8Encoding]::new($false))
    Write-Host "REVIEW audit complete"
    # Compilation overwrites the same files. A new receipt is written only after successful compilation.
    return 'Build'
}
function Read-PpiReviewedScannerEvidence([string]$Path) {
    if(!$Path -or !(Test-Path -LiteralPath $Path)){throw 'Scanner report is required.'}
    if((Get-Item -LiteralPath $Path).Length -gt 16777216){throw 'Scanner report exceeds review bounds.'}
    $report=ConvertFrom-Json -InputObject ([IO.File]::ReadAllText([IO.Path]::GetFullPath($Path)))
    if(!$report.PSObject.Properties['Errors'] -or !$report.PSObject.Properties['McAfeeDetections'] -or $report.ReadOnly -ne $true -or @($report.Errors).Count -ne 0){throw 'Complete read-only scanner evidence without read errors is required.'}
    return $report
}
function Assert-PpiRunnerScannerClear($Report,[System.Collections.IDictionary]$Files,[System.Collections.IDictionary]$Hashes) {
    foreach($entry in $Report.McAfeeDetections){
        $target=([string]$entry.Target).Replace('\\?\','')
        foreach($name in $Files.Keys){
            if($target -eq [IO.Path]::GetFullPath($Files[$name]) -or ($Hashes.Contains($name) -and $entry.SHA256 -and [string]$entry.SHA256 -eq [string]$Hashes[$name])){throw 'Recorded McAfee detection matches this runner path/hash. Execution/recovery refused; vendor clearance required.'}
        }
    }
}
function Start-PpiLegacySourceReview([string]$ReceiptPath,[System.Collections.IDictionary]$Files,[string]$ScannerEvidence,[string]$BuildKey,[string]$SourceRoot) {
    if($BuildKey -notmatch '^[A-Fa-f0-9]{64}$'){throw 'Current source/toolchain build key is required.'}
    if(!(Test-Path -LiteralPath $ReceiptPath) -or (Get-Item -LiteralPath $ReceiptPath).Length -gt 1048576){throw 'Bounded legacy receipt is required.'}
    foreach($file in $Files.Values){if(Test-Path -LiteralPath $file){throw 'Legacy source review only supports absent outputs; no existing binary is adopted or overwritten.'}}
    $priorText=[IO.File]::ReadAllText([IO.Path]::GetFullPath($ReceiptPath));$prior=ConvertFrom-Json -InputObject $priorText
    if($prior.BuildKey -or $prior.Platform -ne 'x64' -or $prior.TargetFramework -ne '.NET Framework 4.8' -or $prior.ExecutableSha256 -notmatch '^[A-Fa-f0-9]{64}$' -or $prior.LibrarySha256 -notmatch '^[A-Fa-f0-9]{64}$'){throw 'Unsupported legacy receipt; not implicitly adopted.'}
    $report=Read-PpiReviewedScannerEvidence $ScannerEvidence
    $oldHashes=[ordered]@{Executable=[string]$prior.ExecutableSha256;Library=[string]$prior.LibrarySha256}
    Assert-PpiRunnerScannerClear $report $Files $oldHashes
    $audit=$ReceiptPath+'.legacy-source-review-'+[Guid]::NewGuid().ToString('N')+'.json'
    $record=[pscustomobject]@{Reason='Owner-approved fresh source build; no legacy binary adoption';Utc=[DateTime]::UtcNow.ToString('o');SourceRoot=[IO.Path]::GetFullPath($SourceRoot);CurrentBuildKey=$BuildKey;PriorReceipt=$priorText;PriorReceiptSha256=(Get-FileHash -LiteralPath $ReceiptPath -Algorithm SHA256).Hash;ScannerEvidence=[IO.Path]::GetFullPath($ScannerEvidence);ScannerEvidenceSha256=(Get-FileHash -LiteralPath $ScannerEvidence -Algorithm SHA256).Hash;StablePaths=$Files;Executed=$false}
    [IO.File]::WriteAllText($audit,(ConvertTo-Json -InputObject $record -Depth 5),[Text.UTF8Encoding]::new($false))
    return $audit
}
