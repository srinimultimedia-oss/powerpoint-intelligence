param([string]$CertificateThumbprint,[string]$DeploymentVersion)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'Common.ps1')
. (Join-Path $repo 'scripts/Safe-Xml.ps1')
$buildLock=New-Object Threading.Mutex($false,'Local\PowerPointIntelligence.EasySetup.Build.v1')
if(!$buildLock.WaitOne(0)){$buildLock.Dispose();throw 'Another easySetup build is running.'}
try {
$source=Get-EasySetupFingerprint $repo
$vswhere=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
Assert-PpiMicrosoftBuildTool $vswhere 'vswhere.exe'
$vs=& $vswhere -latest -products '*' -property installationPath
if(!$vs){throw 'Visual Studio with Office development tools is required on the build computer.'}
$compiler=Join-Path $vs 'MSBuild/Current/Bin/Roslyn/csc.exe'
$msbuild=Join-Path $vs 'MSBuild/Current/Bin/MSBuild.exe'
Assert-PpiMicrosoftBuildTool $compiler 'csc.exe'
Assert-PpiMicrosoftBuildTool $msbuild 'MSBuild.exe'
if(!$CertificateThumbprint){$certificate=Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert | Where-Object {$_.Subject -eq 'CN=PowerPoint Intelligence Local Development' -and $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date).AddDays(1)} | Sort-Object NotAfter -Descending | Select-Object -First 1;if(!$certificate){throw 'No signing certificate is available. Supply -CertificateThumbprint for an approved code-signing certificate, or run Install-Local.ps1 in normal desktop PowerShell first.'};$CertificateThumbprint=$certificate.Thumbprint}
$certificate=Get-PpiApprovedSigningCertificate $CertificateThumbprint
$CertificateThumbprint=$certificate.Thumbprint
if(!$certificate.HasPrivateKey -or $certificate.NotAfter -le (Get-Date)){throw 'An unexpired signing certificate with its private key is required.'}
if(!$DeploymentVersion){$versions=@([Version]'1.0.0.0');foreach($platform in @('x86','x64')){$file=Join-Path $repo "apps/powerpoint-vsto/bin/$platform/Release/PowerPointIntelligence.vsto";if(Test-Path $file){$xml=Read-PpiSafeXml $file;$versions+= [Version]$xml.SelectSingleNode("/*[local-name()='assembly']/*[local-name()='assemblyIdentity']").GetAttribute('version')}};$latest=$versions | Sort-Object -Descending | Select-Object -First 1;if($latest.Revision -ge 65535){throw 'Deployment revision exhausted.'};$DeploymentVersion=[Version]::new($latest.Major,$latest.Minor,$latest.Build,$latest.Revision+1).ToString()}
$priorReceipt=Join-Path $PSScriptRoot 'build.json'
if(Test-Path $priorReceipt){$prior=Get-Content $priorReceipt -Raw | ConvertFrom-Json;$priorVersion=[Version]$prior.Version;if([Version]$DeploymentVersion -le $priorVersion){if($priorVersion.Revision -ge 65535){throw 'Deployment revision exhausted.'};$DeploymentVersion=[Version]::new($priorVersion.Major,$priorVersion.Minor,$priorVersion.Build,$priorVersion.Revision+1).ToString()}}
$DeploymentVersion=([Version]$DeploymentVersion).ToString()
$work=Join-Path $repo ('artifacts/easySetup-build-'+[Guid]::NewGuid().ToString('N'))
$payload=Join-Path $work 'payload';New-Item -ItemType Directory -Force $payload | Out-Null
$config=Read-PpiSafeXml (Join-Path $repo 'apps/powerpoint-vsto/App.config')
$config.SelectSingleNode("/configuration/appSettings/add[@key='TemplateLibraryFolder']").SetAttribute('value','templates')
$configFile=Join-Path $work 'App.config';$config.Save($configFile)
foreach($platform in @('x86','x64')){
    $output=Join-Path $work $platform
    & $msbuild (Join-Path $repo 'apps/powerpoint-vsto/PowerPointIntelligence.csproj') /t:Build /p:Configuration=Release "/p:Platform=$platform" /p:SignManifests=true /p:UseSharedCompilation=false /p:PpiForcePayloadCompile=true /p:PpiEasySetupBuild=true "/p:OutputPath=$output/" "/p:AppConfig=$configFile" "/p:ManifestCertificateThumbprint=$CertificateThumbprint" "/p:ApplicationVersion=$DeploymentVersion" /v:minimal /nologo
    if($LASTEXITCODE){throw "Signed $platform build failed; no installer published."}
    & (Join-Path $repo 'scripts/Test-DemoPackage.ps1') -Platform $platform -OutputDirectory $output
    $manifest=Read-PpiSafeXml (Join-Path $output 'PowerPointIntelligence.dll.manifest')
    $templates=@(Get-ChildItem (Join-Path $repo 'apps/powerpoint-vsto/templates') -Recurse -File -Filter '*.potx')
    if(!$templates.Count){throw 'Licensed local template assets are missing.'}
    foreach($template in $templates){$relative=$template.FullName.Substring((Join-Path $repo 'apps/powerpoint-vsto').Length+1).Replace('\','/');$included=@($manifest.SelectNodes("//*[local-name()='file']") | Where-Object {$_.GetAttribute('name').Replace('\','/') -eq $relative});if($included.Count -ne 1){throw "Template is not covered by the signed application manifest: $relative"}}
    $target=Join-Path $payload $platform;New-Item -ItemType Directory $target | Out-Null
    Get-ChildItem $output -File | Where-Object {$_.Extension -in @('.dll','.config','.manifest','.vsto')} | Copy-Item -Destination $target
    Copy-Item (Join-Path $output 'templates') -Destination $target -Recurse
}
if($source -ne (Get-EasySetupFingerprint $repo)){throw 'Source changed during packaging; no installer published.'}
$labels=@([regex]::Matches((Get-Content (Join-Path $repo 'apps/powerpoint-vsto/DemoTools.cs') -Raw),'add\("[^"]+","[^"]+","([^"]+)"') | ForEach-Object {$_.Groups[1].Value})
$readme=@"
# NexWave — easySetup

Version: $DeploymentVersion
Source SHA-256: $source
Generated UTC: $([DateTime]::UtcNow.ToString('o'))

Run PowerPointIntelligenceSetup.exe after saving and closing PowerPoint. Setup detects Office x86/x64, verifies/extracts its embedded payload to the current user's LocalAppData, and opens the normal VSTO publisher prompt. Keep extracted version folders: VSTO may depend on them. Installation is per Windows user. Updates use a newer signed deployment version: the prior easySetup package is uninstalled through VSTO and the previous package is offered for rollback if installation fails. A first migration from an older repository-based installation may require uninstalling that old package normally if VSTO reports a different installation location. Uninstall through Windows Apps / Programs and Features. No administrator rights, certificate trust stores, antivirus exclusions or security policies are changed by this bootstrapper. A company policy may still require administrator assistance.

## Requirements
Windows desktop, .NET Framework 4.8, supported x86/x64 PowerPoint 2016/2019/Microsoft 365, and Microsoft VSTO Runtime already installed. Prerequisites are not downloaded or silently installed. These product profiles do not establish certification on every Office build. The build computer also requires Visual Studio Office tools and a usable signing certificate. Self-signed development publishers may not be trusted on another computer; use an approved publisher certificate for distribution. Private keys are never included.

## Features in this source
Editable native charts and grouped-shape rendering; Excel/CSV/JSON/XML data sources; explicit Excel worksheet/range/column mapping, refresh and confirmed relinking; thumbnail template library and palette/alignment controls; agendas; text elements, process layouts, universal connectors, Harvey balls and checkboxes; contextual properties; layout, formatting and quality tools. External services need separate endpoints and credentials.

Command labels discovered from the current source:
$($labels -join ', ')

## Included assets and validation limits
Included template packages: $($templates.Count), licensed-computer use only. Third-party templates are not independently authored assets and must not be redistributed outside your licenses. Native fidelity evidence covers 16 tested original slides, not the complete library. The focused desktop mapped Excel test passed six assertions and the focused linked-chart test passed refresh/edit/save-reopen. Full native suite acceptance remains open; packaging does not prove runtime acceptance. Review docs/implementation for detailed evidence.

## Rebuild and automatic refresh
Run .\easySetup\Build-EasySetup.ps1 from normal Windows PowerShell. Optionally supply -CertificateThumbprint. Signed Release builds call this script automatically; recursive packaging is disabled. Start .\easySetup\Watch-EasySetup.ps1 for rebuilding after source edits. Saving source alone does not produce a signed installer unless the watcher is running. A failed build leaves the previous EXE unchanged; use Test-EasySetup.ps1 to reject stale packages before sharing.

## Logs and deployment files
Setup logs: %LOCALAPPDATA%\PowerPointIntelligence\easySetup. The EXE contains both architecture payloads, signed manifests and template assets. easySetup/build.json records source/package hashes and signer details. README.md is regenerated with each successful package. Build work/evidence remains in artifacts. Do not copy source credentials into the package.
"@
$readme | Set-Content -Encoding UTF8 (Join-Path $payload 'README.md')
$inventory=@(Get-ChildItem $payload -Recurse -File | Sort-Object FullName | ForEach-Object {(Get-FileHash $_.FullName -Algorithm SHA256).Hash+"`t"+$_.FullName.Substring($payload.Length+1).Replace('\','/')})
$inventory | Set-Content -Encoding UTF8 (Join-Path $payload 'payload.sha256')
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip=Join-Path $work 'payload.zip';[IO.Compression.ZipFile]::CreateFromDirectory($payload,$zip,[IO.Compression.CompressionLevel]::Optimal,$false)
$bundleHash=(Get-FileHash $zip -Algorithm SHA256).Hash
$bootstrap=(Get-Content (Join-Path $PSScriptRoot 'Bootstrapper.cs') -Raw).Replace('__BUNDLE_HASH__',$bundleHash)
$bootstrapFile=Join-Path $work 'Bootstrapper.cs';$bootstrap | Set-Content -Encoding UTF8 $bootstrapFile
$pending=Join-Path $work 'PowerPointIntelligenceSetup.exe'
& $compiler /nologo /target:winexe /platform:anycpu /langversion:7.3 /r:System.Windows.Forms.dll /r:System.Core.dll /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll "/resource:$zip,IntelligencePayload.zip" "/out:$pending" $bootstrapFile
if($LASTEXITCODE){throw 'Bootstrapper compilation failed.'}
$signature=Set-AuthenticodeSignature -FilePath $pending -Certificate $certificate -HashAlgorithm SHA256
if($signature.Status -ne 'Valid'){throw "Installer signing/trust validation failed: $($signature.Status). Previous published installer was not replaced. Use an approved trusted signing certificate; no trust policy was changed."}
if($source -ne (Get-EasySetupFingerprint $repo)){throw 'Source changed during signing; no installer published.'}
$publicExe=Join-Path $PSScriptRoot 'PowerPointIntelligenceSetup.exe'
if(Test-Path $publicExe){[IO.File]::Replace($pending,$publicExe,(Join-Path $work 'previous-setup.exe'))}else{[IO.File]::Move($pending,$publicExe)}
Copy-Item (Join-Path $payload 'README.md') (Join-Path $PSScriptRoot 'README.md') -Force
[ordered]@{Version=$DeploymentVersion;SourceSha256=$source;ExeSha256=(Get-FileHash (Join-Path $PSScriptRoot 'PowerPointIntelligenceSetup.exe') -Algorithm SHA256).Hash;BundleSha256=$bundleHash;CertificateThumbprint=$certificate.Thumbprint;Architectures=@('x86','x64');TemplateCount=$templates.Count;GeneratedUtc=[DateTime]::UtcNow.ToString('o')} | ConvertTo-Json | Set-Content (Join-Path $PSScriptRoot 'build.json')
Write-Output "Installer published: easySetup/PowerPointIntelligenceSetup.exe; version=$DeploymentVersion"

} finally {$buildLock.ReleaseMutex();$buildLock.Dispose()}
