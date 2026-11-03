. (Join-Path $PSScriptRoot 'Scene-RunnerCache.ps1')
function Start-PpiPiidxSourceRecovery([string]$ReceiptPath,[string]$AssemblyPath,[string]$ScannerEvidence,[object[]]$CurrentSources) {
    $ReceiptPath=Resolve-PpiArtifactInput $ReceiptPath
    $AssemblyPath=Resolve-PpiArtifactInput $AssemblyPath -AllowMissing
    $ScannerEvidence=Resolve-PpiArtifactInput $ScannerEvidence
    if(Test-Path -LiteralPath $AssemblyPath){throw 'PIIDX recovery only accepts an absent output; existing code is never adopted or overwritten.'}
    if(!(Test-Path -LiteralPath $ReceiptPath) -or (Get-Item -LiteralPath $ReceiptPath).Length -gt 1048576){throw 'A bounded prior PIIDX receipt is required.'}
    $priorText=Read-PpiArtifactJsonText $ReceiptPath
    $prior=ConvertFrom-Json -InputObject $priorText
    if($prior.InputSHA256 -notmatch '\A[A-Fa-f0-9]{64}\z' -or $prior.AssemblySHA256 -notmatch '\A[A-Fa-f0-9]{64}\z' -or @($prior.Sources).Count -eq 0 -or @($CurrentSources).Count -eq 0){throw 'PIIDX recovery requires historical and current source/hash provenance.'}
    foreach($source in @($prior.Sources)+@($CurrentSources)) {
        $name=([string]$source.File).Replace('\','/')
        if($name -notmatch '\A(src/Library.Core|tests/library/piidx)/[^/]+\.cs\z' -or $name.Contains('..') -or $source.SHA256 -notmatch '\A[A-Fa-f0-9]{64}\z'){throw 'PIIDX source receipt contains unsupported paths or hashes.'}
    }
    $report=Read-PpiReviewedScannerEvidence $ScannerEvidence
    if($report.Time -is [DateTime]){$time=[DateTimeOffset]$report.Time}
    else{$time=[DateTimeOffset]::Parse([string]$report.Time,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::RoundtripKind)}
    if($time -lt [DateTimeOffset]::UtcNow.AddDays(-1) -or $time -gt [DateTimeOffset]::UtcNow.AddMinutes(5)){throw 'PIIDX recovery requires scanner evidence collected within the last day.'}
    Assert-PpiRunnerScannerClear $report ([ordered]@{Assembly=$AssemblyPath}) ([ordered]@{Assembly=[string]$prior.AssemblySHA256})
    $audit=$ReceiptPath+'.source-recovery-'+[Guid]::NewGuid().ToString('N')+'.json'
    $record=[ordered]@{Reason='Owner-requested rebuild of missing PIIDX from current source at existing stable path';Utc=[DateTime]::UtcNow.ToString('o');StablePath=[IO.Path]::GetFullPath($AssemblyPath);PriorReceipt=$priorText;PriorReceiptSHA256=(Get-PpiContentHash $ReceiptPath);CurrentSources=$CurrentSources;ScannerEvidence=[IO.Path]::GetFullPath($ScannerEvidence);ScannerEvidenceSHA256=(Get-PpiContentHash $ScannerEvidence);RestoredOldBinary=$false;Executed=$false}
    [IO.File]::WriteAllText($audit,($record | ConvertTo-Json -Depth 5),[Text.UTF8Encoding]::new($false))
    return [pscustomobject]@{Audit=$audit;Report=$report}
}
