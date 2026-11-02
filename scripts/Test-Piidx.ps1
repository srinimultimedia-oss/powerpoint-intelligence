param([string]$LibraryFolder,[string]$RunnerRoot)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'Piidx-TestSuite.ps1')
$suite=Import-PpiPiidxSuite $repo $RunnerRoot
if(!$LibraryFolder){$LibraryFolder=Join-Path $repo 'apps/powerpoint-vsto/templates'}
$output=Join-Path $repo ('artifacts/piidx/suite-'+[Guid]::NewGuid().ToString('N'));$null=New-Item -ItemType Directory -Path $output
$results=New-Object System.Collections.Generic.List[object]
foreach($name in @('foundation','sources','catalog','previews','edgecases','workspace','corpus')) {
    $case=Join-Path $output $name
    try {
        $message=switch($name) {
            'foundation' {[PiidxSchemaTests]::Run($case)}
            'sources' {[PiidxSourceTests]::Run($case,$LibraryFolder)}
            'catalog' {[PiidxCatalogTests]::Run($case)}
            'previews' {[PiidxPreviewTests]::Run($case)}
            'edgecases' {[PiidxEdgeTests]::Run($case)}
            'workspace' {[PiidxWorkspaceTests]::Run($case)}
            'corpus' {[PiidxCorpusTests]::Run($case,$LibraryFolder)}
        }
        $results.Add([pscustomobject]@{Name=$name;Status='PASS';Result=$message});Write-Host $message
    } catch {$results.Add([pscustomobject]@{Name=$name;Status='FAIL';Result=$_.Exception.ToString()});Write-Host "FAIL $name"}
}
Assert-PpiPiidxSources $repo $suite.Sources
[pscustomobject]@{Process64=[Environment]::Is64BitProcess;Runtime=[Environment]::Version.ToString();OS=[Environment]::OSVersion.VersionString;OfficeExecuted=$false;AssemblySHA256=$suite.AssemblySHA256;InputSHA256=$suite.InputSHA256;Sources=$suite.Sources;Results=$results.ToArray()} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $output 'results.json')
$results | Format-Table Name,Status
"Evidence: $output"
if(@($results | Where-Object Status -ne PASS).Count){throw 'PIIDX local suite incomplete; inspect results.json.'}