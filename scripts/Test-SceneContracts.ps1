param([ValidateSet('AnyCPU','x86','x64')][string]$Platform='AnyCPU', [switch]$CompileOnly,[string]$RunnerRoot,[switch]$RebuildReviewedRunner,[switch]$RecoverMissingRunner,[switch]$ReviewLegacyReceipt,[string]$ScannerEvidence)
$ErrorActionPreference='Stop'
if($RecoverMissingRunner -and !$RebuildReviewedRunner){throw 'RecoverMissingRunner requires RebuildReviewedRunner and a reviewed scanner report.'}
if($ReviewLegacyReceipt -and (!$RebuildReviewedRunner -or !$RecoverMissingRunner -or $Platform -ne 'x64')){throw 'Explicit legacy review requires x64, RebuildReviewedRunner and RecoverMissingRunner.'}
$repo=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'Scene-RunnerCache.ps1')
if(!$RunnerRoot){$RunnerRoot=$repo}
$RunnerRoot=(Resolve-Path -LiteralPath $RunnerRoot).ProviderPath
Push-Location $repo
try {
    $vswhere=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
    $vsPath=& $vswhere -latest -products '*' -property installationPath
    $compiler=Join-Path $vsPath 'MSBuild/Current/Bin/Roslyn/csc.exe'
    $refs=Join-Path ${env:ProgramFiles(x86)} 'Reference Assemblies/Microsoft/Framework/.NETFramework/v4.8'
    if(!(Test-Path (Join-Path $refs 'mscorlib.dll'))) { throw '.NET Framework 4.8 targeting pack is required.' }
    $dependency=Join-Path $repo 'artifacts/refs/newtonsoft/lib/net45/Newtonsoft.Json.dll'
    if(!(Test-Path $dependency)) { throw 'Run Test-S1.ps1 first to restore the pinned Newtonsoft.Json dependency.' }
    $output=Join-Path $RunnerRoot "artifacts/scene-contracts-$Platform"
    $results=Join-Path $repo "artifacts/scene-results-$Platform"
    $corpus=Join-Path $results 'corpus'
    $roundtrip=Join-Path $results 'roundtrip'
    $receiptPath=Join-Path $output 'compiled.json'
    $files=[ordered]@{Executable=(Join-Path $output 'SceneContractTests.exe');Library=(Join-Path $output 'Chart.Renderers.Contracts.dll');Dependency=(Join-Path $output 'Newtonsoft.Json.dll')}
    $key=Get-PpiSceneBuildKey $repo $Platform $compiler @("$refs/mscorlib.dll","$refs/System.dll","$refs/System.Core.dll",$dependency)
    $legacyAudit=$null
    if($ReviewLegacyReceipt){$legacyAudit=Start-PpiLegacySourceReview $receiptPath $files $ScannerEvidence $key $repo;$state='Build'}
    elseif($RebuildReviewedRunner){$state=Start-PpiReviewedRunnerBuild $receiptPath $files $ScannerEvidence -RecoverMissing:$RecoverMissingRunner}else{$state=Get-PpiSceneCacheState $receiptPath $key $files}
    $priorLaunch=Join-Path $output 'launch.json'
    if(Test-Path -LiteralPath $priorLaunch){$prior=Get-Content -LiteralPath $priorLaunch -Raw | ConvertFrom-Json;if($prior.Status -ne 'PASS' -and ($null -eq $prior.ExitCode -or $prior.PresentAfter -eq $false)){throw 'Previous scene runner failed to launch or disappeared. Review its launch/scanner evidence before running it again. No replacement was created.'}}
    if($state -eq 'Build') {
        New-Item -ItemType Directory -Force $output | Out-Null
        Copy-Item $dependency $files.Dependency
        $references=@('/nostdlib+',"/r:$refs/mscorlib.dll","/r:$refs/System.dll","/r:$refs/System.Core.dll","/r:$dependency")
        & $compiler /nologo /deterministic+ /langversion:7.3 "/platform:$Platform" /target:library "/out:$($files.Library)" @references src/Chart.Renderers/SceneContracts.cs src/Chart.Renderers/SceneCodec.cs src/Chart.Renderers/StrictSceneJson.cs src/Shared.Contracts/BoundedJson.cs
        if($LASTEXITCODE){throw 'Scene library build failed.'}
        & $compiler /nologo /deterministic+ /langversion:7.3 "/platform:$Platform" /target:exe "/out:$($files.Executable)" @references "/r:$($files.Library)" tests/s3/SceneContractTests.cs tests/s3/SceneArtifactPath.cs
        if($LASTEXITCODE){throw 'Scene runner build failed.'}
        $hashes=[ordered]@{};foreach($name in $files.Keys){$hashes[$name]=(Get-FileHash -LiteralPath $files[$name] -Algorithm SHA256).Hash}
        if($ScannerEvidence){$report=Read-PpiReviewedScannerEvidence $ScannerEvidence;Assert-PpiRunnerScannerClear $report $files $hashes}
        [pscustomobject]@{BuildKey=$key;Platform=$Platform;TargetFramework='.NET Framework 4.8';Hashes=$hashes;ExecutableSha256=$hashes.Executable;LibrarySha256=$hashes.Library;LegacySourceReview=$legacyAudit;ScannerEvidence=$(if($legacyAudit){[IO.Path]::GetFullPath($ScannerEvidence)}else{$null});ScannerEvidenceSha256=$(if($legacyAudit){(Get-FileHash -LiteralPath $ScannerEvidence -Algorithm SHA256).Hash}else{$null});Time=(Get-Date).ToString('o')} | ConvertTo-Json -Depth 5 | Set-Content $receiptPath
    }
    $currentReceipt=ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($receiptPath))
    if($currentReceipt.LegacySourceReview){
        if(!(Test-Path -LiteralPath $currentReceipt.LegacySourceReview) -or !(Test-Path -LiteralPath $currentReceipt.ScannerEvidence)){throw 'Reviewed recovery audit/evidence is missing; execution refused.'}
        if((Get-FileHash -LiteralPath $currentReceipt.ScannerEvidence -Algorithm SHA256).Hash -ne $currentReceipt.ScannerEvidenceSha256){throw 'Reviewed scanner evidence changed; execution refused.'}
        $report=Read-PpiReviewedScannerEvidence $currentReceipt.ScannerEvidence
        $launchHashes=[ordered]@{};foreach($name in $files.Keys){$launchHashes[$name]=(Get-FileHash -LiteralPath $files[$name] -Algorithm SHA256).Hash}
        Assert-PpiRunnerScannerClear $report $files $launchHashes
    }
    $executable=$files.Executable
    Write-Output "$state scene runner at stable path: $executable"
    New-Item -ItemType Directory -Force $results | Out-Null
    & node.exe tests/s3/scene-fixtures.mjs --write-corpus $corpus
    if($LASTEXITCODE){throw 'Scene corpus generation failed.'}
    if($CompileOnly) { Write-Output "COMPILE ONLY: $Platform scene runner prepared ($state); execution not validated."; return }
    $launch=[ordered]@{Time=(Get-Date).ToString('o');Platform=$Platform;Executable=$executable;CompiledReceipt=(Join-Path $output 'compiled.json');PresentBefore=(Test-Path -LiteralPath $executable);Status='FAIL';ExitCode=$null;PresentAfter=$null;Errors=@()}
    try {
        & $executable tests/s3/fixtures/all-primitives.v1.json $roundtrip $corpus 2>&1 | Tee-Object -FilePath (Join-Path $results 'execution.log')
        $launch.ExitCode=$LASTEXITCODE
        if($LASTEXITCODE) { throw 'Scene .NET contract tests failed.' }
        $launch.Status='PASS'
    } catch {
        $exception=$_.Exception
        $chain=New-Object System.Collections.Generic.List[object]
        while($exception) {
            $chain.Add([pscustomobject]@{Type=$exception.GetType().FullName;HResult=('0x{0:X8}' -f $exception.HResult);NativeErrorCode=$(if($exception -is [ComponentModel.Win32Exception]){$exception.NativeErrorCode});Message=$exception.Message})
            $exception=$exception.InnerException
        }
        $launch.Errors=@($chain.ToArray())
        throw
    } finally {
        $launch.PresentAfter=Test-Path -LiteralPath $executable
        $launch | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $output 'launch.json')
        $launch | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $results 'launch.json')
        if(!$launch.PresentAfter) { Write-Warning 'The compiled runner is now missing. Run Collect-SceneThreatEvidence.ps1 in an approved desktop session before restoring or rebuilding it; no scanner cause is inferred here.' }
    }
    & node.exe tests/s3/validate-roundtrip.mjs $corpus $roundtrip
    if($LASTEXITCODE) { throw 'Scene cross-language round-trip validation failed.' }
    # Exercise the real, already-provenance-checked runner with a data-derived
    # traversal name. Everything, including the outside-output sentinel, is owned test data.
    $negative=Join-Path $results ('security-negative-'+[Guid]::NewGuid().ToString('N'))
    $negativeOutput=Join-Path $negative 'output'
    New-Item -ItemType Directory -Force $negativeOutput | Out-Null
    $sentinel=Join-Path $negative 'sentinel.v1.json'
    [IO.File]::WriteAllText($sentinel,'owned sentinel')
    $golden=(Resolve-Path 'tests/s3/fixtures/all-primitives.v1.json').ProviderPath
    ConvertTo-Json -InputObject @([ordered]@{name='../sentinel';valid=$true;json=[IO.File]::ReadAllText($golden)}) -Depth 4 | Set-Content (Join-Path $negative 'cases.json')
    if((Get-FileHash -LiteralPath $executable -Algorithm SHA256).Hash -ne (Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json).Hashes.Executable){throw 'Runner changed before the security regression; execution refused.'}
    $negativeRun=Start-Process -FilePath $executable -ArgumentList @(('"'+$golden+'"'),('"'+$negativeOutput+'"'),('"'+$negative+'"')) -WindowStyle Hidden -Wait -PassThru -RedirectStandardOutput (Join-Path $negative 'stdout.log') -RedirectStandardError (Join-Path $negative 'stderr.log')
    if($negativeRun.ExitCode -ne 1 -or [IO.File]::ReadAllText($sentinel) -ne 'owned sentinel' -or !(Select-String -LiteralPath (Join-Path $negative 'stderr.log') -Pattern 'Corpus name' -Quiet)){throw 'Real scene runner did not reject the unsafe corpus name without altering the sentinel.'}
    Write-Output 'PASS real scene runner rejects corpus traversal before outside-output mutation.'
} finally { Pop-Location }
