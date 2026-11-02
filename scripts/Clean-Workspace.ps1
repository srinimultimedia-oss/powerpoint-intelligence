param([switch]$Apply,[ValidateRange(0,20)][int]$KeepLatestSnapshots=2)
$ErrorActionPreference='Stop'
$repo=[IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent)).TrimEnd('\')
$artifacts=Join-Path $repo 'artifacts'
$candidates=New-Object System.Collections.Generic.List[string]
$runs=@(Get-ChildItem -LiteralPath $artifacts -Directory -Filter 'validation-*' | Sort-Object Name -Descending)
foreach($run in ($runs | Select-Object -Skip $KeepLatestSnapshots)) {
    $snapshot=Join-Path $run.FullName 'source'
    if(Test-Path -LiteralPath $snapshot) {
        Get-ChildItem -LiteralPath $snapshot -Directory -Recurse -ErrorAction Stop | Where-Object { $_.Name -eq 'node_modules' -and $_.FullName -notmatch '[\\/]node_modules[\\/]' } | ForEach-Object {$candidates.Add($_.FullName)}
    }
}
foreach($relative in @('artifacts/refs/compiler','artifacts/refs/compiler.zip','artifacts/refs/powerpoint','artifacts/refs/powerpoint.zip','artifacts/refs/office','artifacts/refs/newtonsoft.zip','apps/powerpoint-vsto/obj')) {
    $path=Join-Path $repo $relative;if(Test-Path -LiteralPath $path){$candidates.Add($path)}
}
$records=New-Object System.Collections.Generic.List[object]
foreach($candidate in ($candidates | Select-Object -Unique)) {
    $resolved=(Resolve-Path -LiteralPath $candidate).ProviderPath
    if(!$resolved.StartsWith($repo+'\',[StringComparison]::OrdinalIgnoreCase)) {throw 'Cleanup target escaped repository.'}
    if(!($resolved.StartsWith($artifacts+'\',[StringComparison]::OrdinalIgnoreCase) -or $resolved -eq (Join-Path $repo 'apps/powerpoint-vsto/obj'))) {throw 'Unexpected cleanup target.'}
    $item=Get-Item -LiteralPath $resolved
    if(($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0){throw 'Cleanup refuses reparse points.'}
    $files=if($item.PSIsContainer){@(Get-ChildItem -LiteralPath $resolved -Recurse -File)}else{@($item)}
    if($item.PSIsContainer -and @(Get-ChildItem -LiteralPath $resolved -Recurse -Attributes ReparsePoint).Count -gt 0){throw 'Cleanup refuses directories containing reparse points.'}
    $size=($files | Measure-Object Length -Sum).Sum
    $records.Add([pscustomobject]@{Path=$resolved;Files=$files.Count;Bytes=$size;Action=$(if($Apply){'Remove'}else{'Preview'})})
}
$report=Join-Path $artifacts ('cleanup-'+(Get-Date -Format 'yyyyMMdd-HHmmss')+'.json')
$records | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $report
if($Apply){foreach($record in $records){Remove-Item -LiteralPath $record.Path -Recurse -Force}}
[pscustomobject]@{Applied=[bool]$Apply;Targets=$records.Count;Files=($records | Measure-Object Files -Sum).Sum;MB=[math]::Round(($records | Measure-Object Bytes -Sum).Sum/1MB,1);Report=$report} | ConvertTo-Json
