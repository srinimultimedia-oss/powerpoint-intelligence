param([switch]$InteractiveOffice,[switch]$ProceduralOffice,[ValidateSet('AnyCPU','x86','x64')][string]$Platform='x64',[string]$RunnerRoot,[switch]$CompileOnly,[switch]$RebuildReviewedRunner,[switch]$RecoverMissingRunner,[string]$ScannerEvidence)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'Scene-RunnerCache.ps1')
if(!$RunnerRoot){$RunnerRoot=$repo}
$RunnerRoot=(Resolve-Path -LiteralPath $RunnerRoot).ProviderPath
Push-Location $repo
$executionLeases=@()
try {
    $vswhere=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
    Assert-PpiMicrosoftBuildTool $vswhere 'vswhere.exe'
    $vsPath=& $vswhere -latest -products '*' -property installationPath
    $compiler=Join-Path $vsPath 'MSBuild/Current/Bin/Roslyn/csc.exe'
    Assert-PpiMicrosoftBuildTool $compiler 'csc.exe'
    $pia=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Shared/Visual Studio Tools for Office/PIA/Office15'
    $framework=Join-Path ${env:ProgramFiles(x86)} 'Reference Assemblies/Microsoft/Framework/.NETFramework/v4.8'
    $output=Join-Path $repo "artifacts/template-library-$Platform"
    $cache=Join-Path $RunnerRoot "artifacts/template-library-$Platform"
    $executable=Join-Path $cache 'TemplateLibraryTests.exe'
    $receipt=Join-Path $cache 'compiled.json'
    $files=[ordered]@{Executable=$executable}
    $sources=@('apps/powerpoint-vsto/TemplateCatalog.cs','apps/powerpoint-vsto/TemplateLibrary.cs','apps/powerpoint-vsto/LocalLibraryGallery.cs','apps/powerpoint-vsto/LocalLibraryWorkspace.cs','tests/library/LegacyLibraryCacheFixture.cs','tests/library/TemplateFidelity.cs','tests/library/TemplateFidelityTests.cs','tests/library/TemplateLibraryTests.cs')
    $sources+=@(Get-ChildItem -LiteralPath (Join-Path $repo 'src/Library.Core') -File -Filter '*.cs' | ForEach-Object {$_.FullName.Substring($repo.Length+1).Replace('\','/')})
    $refs=@("$framework/System.IO.Compression.dll","$framework/System.IO.Compression.FileSystem.dll","$framework/mscorlib.dll","$framework/System.dll","$framework/System.Core.dll","$framework/System.Xml.dll","$framework/System.Xml.Linq.dll","$framework/Microsoft.CSharp.dll","$framework/System.Windows.Forms.dll","$framework/System.Drawing.dll","$pia/office.dll","$pia/Microsoft.Office.Interop.PowerPoint.dll")
    $sourceBuild={param($targets)
        $compilerOutput=& $compiler /nologo "/r:$framework/System.IO.Compression.dll" "/r:$framework/System.IO.Compression.FileSystem.dll" /deterministic+ "/pathmap:$repo=/_/PowerPointIntelligence" /nostdlib+ /langversion:7.3 "/platform:$Platform" /target:exe "/out:$($targets.Executable)" "/link:$pia/office.dll" "/link:$pia/Microsoft.Office.Interop.PowerPoint.dll" "/r:$framework/mscorlib.dll" "/r:$framework/System.dll" "/r:$framework/System.Core.dll" "/r:$framework/System.Xml.dll" "/r:$framework/Microsoft.CSharp.dll" "/r:$framework/System.Windows.Forms.dll" "/r:$framework/System.Drawing.dll" @sources
        if($LASTEXITCODE){throw ('Template library compilation failed. '+($compilerOutput -join [Environment]::NewLine))}
    }.GetNewClosure()
    $key=Get-PpiSceneBuildKey -Root $repo -Platform $Platform -Compiler $compiler -References $refs -Sources $sources -CacheKind 'template'
    if($RebuildReviewedRunner){$state=Start-PpiReviewedRunnerBuild $receipt $files $ScannerEvidence -RecoverMissing:$RecoverMissingRunner}else{$state=Get-PpiSceneCacheState $receipt $key $files}
    $launchFile=Join-Path $cache 'launch.json'
    if(Test-Path -LiteralPath $launchFile){$prior=Get-Content -LiteralPath $launchFile -Raw | ConvertFrom-Json;if($prior.Status -ne 'PASS' -and ($null -eq $prior.ExitCode -or $prior.PresentAfter -eq $false)){throw 'Previous template runner failed to launch or disappeared. Review execution/scanner evidence before retrying; no replacement was created.'}}
    if($state -eq 'Build') {
        New-Item -ItemType Directory -Force $cache | Out-Null
        & $sourceBuild $files
        [pscustomobject]@{BuildKey=$key;Platform=$Platform;Hashes=@{Executable=(Get-FileHash -LiteralPath $executable -Algorithm SHA256).Hash};Time=(Get-Date).ToString('o')} | ConvertTo-Json -Depth 4 | Set-Content $receipt
    }
    # The stable runner resolves assets beside its own assembly, not the source snapshot.
    # Supply an explicit absolute path without altering or relocating the executable.
    $templateRoot=Join-Path $repo 'apps/powerpoint-vsto/templates'
    if(Test-Path -LiteralPath $templateRoot){
        $configuration=New-Object System.Xml.XmlDocument
        $configuration.LoadXml('<configuration><appSettings><add key="TemplateLibraryFolder" value="" /></appSettings></configuration>')
        $configuration.configuration.appSettings.add.SetAttribute('value',(Resolve-Path -LiteralPath $templateRoot).ProviderPath)
        $configuration.Save($executable+'.config')
    }else{throw "Template assets are unavailable in the validation source: $templateRoot"}
    Write-Output "$state template runner at stable path: $executable"
    $executionLeases=Get-PpiReproducedArtifacts $files $sourceBuild
    Write-Output 'PASS cached template code matches independently rebuilt current source; execution file locked.'
    if($CompileOnly){Write-Output 'Runner prepared; execution not validated.';return}
    New-Item -ItemType Directory -Force $output | Out-Null
    $launch=[ordered]@{Executable=$executable;Status='FAIL';ExitCode=$null;PresentBefore=(Test-Path -LiteralPath $executable);PresentAfter=$null;Error=$null;InteractiveOffice=[bool]$InteractiveOffice;Time=(Get-Date).ToString('o')}
    try {
        if($InteractiveOffice){$originalRoot=Join-Path $repo 'apps/powerpoint-vsto/templates';if(!$ProceduralOffice -and (Test-Path -LiteralPath $originalRoot)){& $executable $output --source-originals}else{& $executable $output --interactive-office}}
        else {& $executable $output}
        $launch.ExitCode=$LASTEXITCODE
        if($LASTEXITCODE){throw 'Template library validation failed.'}
        $launch.Status='PASS'
    }catch{$launch.Error=$_.Exception.Message;throw}
    finally {$launch.PresentAfter=Test-Path -LiteralPath $executable;$launch | ConvertTo-Json -Depth 4 | Set-Content $launchFile;$launch | ConvertTo-Json -Depth 4 | Set-Content (Join-Path $output 'launch.json')}
} finally {foreach($lease in $executionLeases){$lease.Stream.Dispose()};Pop-Location}
