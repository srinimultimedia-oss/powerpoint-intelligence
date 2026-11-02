param([switch]$InteractiveOffice, [ValidateRange(1,10)][int]$OfficeRuns = 3,[ValidateRange(30,3600)][int]$LocalTimeoutSeconds=300,[ValidateRange(60,7200)][int]$OfficeTimeoutSeconds=3600)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$runDirectory = Join-Path $repo ('artifacts/validation-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0,8))
New-Item -ItemType Directory -Path $runDirectory | Out-Null
. (Join-Path $PSScriptRoot 'Source-Fingerprint.ps1')
. (Join-Path $PSScriptRoot 'Check-Process.ps1')
$originalRepo=$repo
$before=Get-PpiSourceFingerprint $repo
$snapshot=Join-Path $runDirectory 'source'
New-Item -ItemType Directory -Path $snapshot | Out-Null
foreach($folder in @('apps','src','scripts','tests','packages','docs')) {
    Get-ChildItem (Join-Path $repo $folder) -File -Recurse | Where-Object {
        $relative=$_.FullName.Substring($repo.Length+1)
        ($relative -match '(^|[\\/])node_modules[\\/]' -or $relative -notmatch '(^|[\\/])(bin|obj|artifacts)[\\/]') -and
        $_.Name -notlike '.env*' -and $_.Extension -notin @('.pfx','.p12','.snk','.key','.pem')
    } | ForEach-Object {
        $destination=Join-Path $snapshot $_.FullName.Substring($repo.Length+1)
        $parent=Split-Path $destination -Parent
        if(!(Test-Path $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
        Copy-Item -LiteralPath $_.FullName -Destination $destination
    }
}
$snapshotArtifacts=Join-Path $snapshot 'artifacts'
New-Item -ItemType Directory -Path $snapshotArtifacts -Force | Out-Null
$dependency=Join-Path $repo 'artifacts/refs/newtonsoft/lib/net45/Newtonsoft.Json.dll'
if(Test-Path $dependency) {
    $dependencyFolder=Join-Path $snapshotArtifacts 'refs/newtonsoft/lib/net45'
    New-Item -ItemType Directory -Force -Path $dependencyFolder | Out-Null
    Copy-Item $dependency $dependencyFolder
}
$receipt=Join-Path $repo 'artifacts/installed-demo-receipt.json'
if(Test-Path $receipt) { Copy-Item $receipt $snapshotArtifacts }
$after=Get-PpiSourceFingerprint $repo
$captured=Get-PpiSourceFingerprint $snapshot
if($before -ne $after -or $captured -ne $after) { throw 'Source changed while capturing validation snapshot. No test result is valid; retry after edits finish.' }
[pscustomobject]@{Original=$repo;Snapshot=$snapshot;SourceSha256=$captured} | ConvertTo-Json | Set-Content (Join-Path $runDirectory 'source-snapshot.json')
$repo=$snapshot
$scriptRoot=Join-Path $repo 'scripts'
Write-Host "Source snapshot ready: $captured"
$results = New-Object System.Collections.Generic.List[object]
$validationStarted = Get-Date
$shellPath = (Get-Process -Id $PID).Path
function Invoke-Check([string]$Name, [string]$Executable, [string[]]$Arguments) {
    if($script:CheckTimedOut){throw 'Validation stopped after a timed-out owned check. Inspect its stdout/stderr before retrying; no user Office process was closed.'}
    $log = Join-Path $runDirectory ($Name + '.log')
    Write-Host "Running $Name"
    try {
        $null = Get-Command $Executable -ErrorAction Stop
        $timeout=if($Name -match 'office|installed-ribbon|demo-feature-runtime'){$OfficeTimeoutSeconds}else{$LocalTimeoutSeconds}
        $execution=Invoke-PpiCheckProcess $Executable $Arguments $log $timeout
        $code=$execution.ExitCode
        if($execution.TimedOut){$script:CheckTimedOut=$true}
        $results.Add([pscustomobject]@{ Name=$Name; ExitCode=$code; Log=$log; Status=$(if ($code -eq 0) {'PASS'} else {'FAIL'}) })
    } catch {
        $_ | Out-String | Add-Content $log
        $results.Add([pscustomobject]@{ Name=$Name; ExitCode=1; Log=$log; Status='FAIL' })
    }
}
Push-Location $repo
try {
    Invoke-Check 's0-unit' $shellPath @('-NoProfile','-File',"$scriptRoot/Test-S0.ps1")
    Invoke-Check 'piidx-engine' $shellPath @('-NoProfile','-File',"$scriptRoot/Test-Piidx.ps1",'-RunnerRoot',$originalRepo)
    Invoke-Check 'piidx-isolation' $shellPath @('-NoProfile','-File',"$scriptRoot/Test-PiidxIsolation.ps1")
    Invoke-Check 'ribbon-definition' $shellPath @('-NoProfile','-File',"$scriptRoot/Test-Ribbon.ps1")
    Invoke-Check 'demo-feature-compile' $shellPath @('-NoProfile','-File',"$scriptRoot/Test-DemoFeatures.ps1")
    Invoke-Check 's1-domain' $shellPath @('-NoProfile','-File',"$scriptRoot/Test-S1.ps1")
    foreach ($testPlatform in @('AnyCPU','x86','x64')) {
        Invoke-Check "compatibility-$testPlatform" $shellPath @('-NoProfile','-File',"$scriptRoot/Test-OfficeCompatibility.ps1",'-Platform',$testPlatform)
    }
    Invoke-Check 'json-schema' 'npm.cmd' @('test','--prefix','packages/chart-domain')
    Invoke-Check 'calculations' 'npm.cmd' @('test','--prefix','packages/chart-core')
    Invoke-Check 's2-chart-engine' 'npm.cmd' @('test','--prefix','packages/chart-engine')
    Invoke-Check 's2-typecheck' 'npm.cmd' @('run','typecheck','--prefix','packages/chart-engine')
    foreach ($testPlatform in @('AnyCPU','x86','x64')) {
        Invoke-Check "scene-contracts-$testPlatform" $shellPath @('-NoProfile','-File',"$scriptRoot/Test-SceneContracts.ps1",'-Platform',$testPlatform,'-RunnerRoot',$originalRepo)
    }
    Invoke-Check 'template-library' $shellPath @('-NoProfile','-File',"$scriptRoot/Test-TemplateLibrary.ps1",'-RunnerRoot',$originalRepo)
    Invoke-Check 'native-renderer-core' $shellPath @('-NoProfile','-File',"$scriptRoot/Test-NativeRenderer.ps1")
    if (($results | Where-Object Name -eq 's1-domain').Status -eq 'PASS') {
        Invoke-Check 'cross-language' 'node.exe' @('packages/chart-domain/test/validate-roundtrip.mjs','artifacts/domain-roundtrip.v2.json')
    }
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
    $vsPath = & $vswhere -latest -products '*' -property installationPath
    foreach ($testPlatform in @('AnyCPU','x86','x64')) {
        Invoke-Check "compile-$testPlatform" (Join-Path $vsPath 'MSBuild/Current/Bin/MSBuild.exe') @('apps/powerpoint-vsto/PowerPointIntelligence.csproj','/t:Compile','/p:Configuration=Release',"/p:Platform=$testPlatform",'/v:minimal','/nologo')
    }
    if ($InteractiveOffice) {
        Invoke-Check 'template-library-office' $shellPath @('-NoProfile','-File',"$scriptRoot/Test-TemplateLibrary.ps1",'-InteractiveOffice','-RunnerRoot',$originalRepo)
        Invoke-Check 'installed-ribbon' $shellPath @('-NoProfile','-File',"$scriptRoot/Test-InstalledRibbon.ps1")
        Invoke-Check 'demo-feature-runtime' $shellPath @('-NoProfile','-File',"$scriptRoot/Test-DemoFeatures.ps1",'-InteractiveOffice')
        Invoke-Check 'native-shapes-office' $shellPath @('-NoProfile','-File',"$scriptRoot/Test-NativeRenderer.ps1",'-InteractiveOffice')
        for ($iteration = 1; $iteration -le $OfficeRuns; $iteration++) {
            Invoke-Check "office-$iteration" $shellPath @('-NoProfile','-File',"$scriptRoot/Test-S0.ps1",'-InteractiveOffice')
            # Do not repeat an activation denial in an execution context that cannot run Office.
            if (Select-String -Path (Join-Path $runDirectory "office-$iteration.log") -Pattern '80070005|Save and close user presentations before runtime validation' -Quiet) { break }
        }
        # Crash metadata only: never save event messages, document paths or content.
        $crashLog = Join-Path $runDirectory 'office-crashes.json'
        try {
            $events = Get-WinEvent -FilterHashtable @{LogName='Application'; Id=1000; StartTime=$validationStarted} -ErrorAction Stop
            $crashes = @($events | ForEach-Object {
                $eventXml = [xml]$_.ToXml()
                $fields = @{}
                foreach ($field in $eventXml.Event.EventData.Data) { $fields[$field.Name] = $field.'#text' }
                if ($fields['AppName'] -match '^(POWERPNT|EXCEL)\.EXE$') {
                    [pscustomobject]@{Time=$_.TimeCreated; Application=$fields['AppName']; Module=$fields['ModuleName']; ExceptionCode=$fields['ExceptionCode']}
                }
            })
            ConvertTo-Json -InputObject $crashes | Set-Content $crashLog
        } catch { ('Crash evidence unavailable: ' + $_.Exception.GetType().Name) | Set-Content $crashLog }
    }
    $gitCommand = Get-Command git.exe -ErrorAction SilentlyContinue
    $gitPath = if ($gitCommand) { $gitCommand.Source } else {
        @((Join-Path $env:ProgramFiles 'Git/cmd/git.exe'),
          (Join-Path $env:USERPROFILE '.cache/codex-runtimes/codex-primary-runtime/dependencies/native/git/cmd/git.exe')) |
            Where-Object { Test-Path $_ } | Select-Object -First 1
    }
    if (!$gitPath) { $gitPath = 'git.exe' }
    Push-Location $originalRepo
    try { Invoke-Check 'diff-check' $gitPath @('diff','--check') } finally { Pop-Location }
} finally {
    Pop-Location
    $results | ConvertTo-Json -Depth 4 | Set-Content (Join-Path $runDirectory 'results.json')
    $lines = @('# Validation results', '', '| Check | Result | Log |', '|---|---|---|')
    foreach ($result in $results) { $lines += "| $($result.Name) | $($result.Status) | $([IO.Path]::GetFileName($result.Log)) |" }
    $lines += @('', 'Compile targets do not validate signed deployment, VSTO loading, ribbon/task pane, or Office runtime behavior.',
        'Compatibility-policy checks do not certify Office 2016/2019/365 installations.',
        'Office logs stop at the first failed assertion; later scenarios are not passes.',
        'Office checks run only with -InteractiveOffice. Full process restart and add-in UI require separate validation.')
    $lines | Set-Content (Join-Path $runDirectory 'results.md')
    $results | Select-Object Name,Status | Format-Table -AutoSize | Out-Host
    Write-Host "Results saved to $runDirectory"
}
if ($results | Where-Object Status -eq 'FAIL') { exit 1 }
