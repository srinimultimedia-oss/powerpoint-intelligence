param([switch]$InteractiveOffice, [switch]$CompileOnly, [ValidateSet('AnyCPU','x86','x64')][string]$Platform='AnyCPU')
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
Push-Location $repo
try {
    $vswhere=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
    $vsPath=& $vswhere -latest -products '*' -property installationPath
    $compiler=Join-Path $vsPath 'MSBuild/Current/Bin/Roslyn/csc.exe'
    $refs=Join-Path ${env:ProgramFiles(x86)} 'Reference Assemblies/Microsoft/Framework/.NETFramework/v4.8'
    $pia=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Shared/Visual Studio Tools for Office/PIA/Office15'
    $dependency=Join-Path $repo 'artifacts/refs/newtonsoft/lib/net45/Newtonsoft.Json.dll'
    if(!(Test-Path $dependency)) { throw 'Run Test-S1.ps1 first to restore Newtonsoft.Json.' }
    $output=Join-Path $repo $(if($Platform -eq 'AnyCPU') {'artifacts/native-renderer'} else {"artifacts/native-renderer-$Platform"})
    New-Item -ItemType Directory -Force $output | Out-Null
    Copy-Item $dependency (Join-Path $output 'Newtonsoft.Json.dll') -Force
    & node.exe packages/chart-engine/tools/render-chart-v2.mjs tests/s1/fixtures/clustered-column.v2.json "$output/scene.v1.json"
    if($LASTEXITCODE) { throw 'Canonical scene fixture generation failed.' }
    $references=@('/nostdlib+',"/r:$refs/mscorlib.dll","/r:$refs/System.dll","/r:$refs/System.Core.dll","/r:$refs/System.Xml.dll","/r:$dependency","/link:$pia/office.dll","/link:$pia/Microsoft.Office.Interop.PowerPoint.dll")
    $resources=@('/resource:apps/powerpoint-vsto/RendererSamples/canonical.json,PowerPointIntelligence.RendererSamples.canonical.json','/resource:apps/powerpoint-vsto/RendererSamples/scene.json,PowerPointIntelligence.RendererSamples.scene.json')
    & $compiler /nologo /langversion:7.3 "/platform:$Platform" /target:exe "/out:$output/NativeRendererTests.exe" @references @resources src/Chart.Domain/ChartDocument.cs src/Chart.Domain/ChartCodec.cs src/Chart.Renderers/SceneContracts.cs src/Chart.Renderers/SceneCodec.cs src/Chart.Renderers/StrictSceneJson.cs src/Shared.Contracts/BoundedJson.cs src/Chart.Renderers/NativeRenderingCore.cs src/Chart.Renderers/PowerPointSceneSurface.cs apps/powerpoint-vsto/RendererRibbonCommand.cs tests/s3/RendererRibbonTests.cs tests/s3/NativeRendererTests.cs src/Library.Core/PathPolicy.cs
    if($LASTEXITCODE) { throw 'Native renderer compilation failed.' }
    if($CompileOnly) { Write-Output 'COMPILE ONLY: native renderer and tests; no unit or Office execution.'; return }
    $executionLog=Join-Path $output 'unit-execution.log'
    & "$output/NativeRendererTests.exe" tests/s1/fixtures/clustered-column.v2.json "$output/scene.v1.json" $output 2>&1 | Tee-Object -FilePath $executionLog
    if($LASTEXITCODE) { throw 'Native renderer core or ribbon preparation tests failed.' }
    if($Platform -ne 'AnyCPU' -and !(Select-String -LiteralPath $executionLog -SimpleMatch "Native renderer runner architecture=$Platform" -Quiet)) { throw 'Native renderer runner did not report the requested process architecture.' }
    . (Join-Path $PSScriptRoot 'Source-Fingerprint.ps1')
    [pscustomobject]@{Time=(Get-Date).ToString('o');Platform=$Platform;SourceSha256=(Get-PpiSourceFingerprint $repo);ExecutableSha256=(Get-FileHash "$output/NativeRendererTests.exe" -Algorithm SHA256).Hash;UnitExitCode=$LASTEXITCODE;Log=$executionLog} | ConvertTo-Json | Set-Content (Join-Path $output 'execution.json')
    if($InteractiveOffice) {
        & "$output/NativeRendererTests.exe" tests/s1/fixtures/clustered-column.v2.json "$output/scene.v1.json" $output --interactive-office 2>&1 | Tee-Object -FilePath (Join-Path $output 'office-execution.log')
        if($LASTEXITCODE) { throw 'Native renderer interactive Office tests failed.' }
    } else { Write-Output 'SKIP native shape Office round trip without -InteractiveOffice.' }
} finally { Pop-Location }
