param([switch]$SourceOnly)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'Common.ps1')
$errors=@();Get-ChildItem $PSScriptRoot -Filter '*.ps1' | ForEach-Object {$tokens=$null;$parseErrors=$null;[Management.Automation.Language.Parser]::ParseFile($_.FullName,[ref]$tokens,[ref]$parseErrors)|Out-Null;$errors+=@($parseErrors)}
if($errors.Count){throw ($errors -join '; ')}
[xml]$project=Get-Content (Join-Path $repo 'apps/powerpoint-vsto/PowerPointIntelligence.csproj') -Raw
if(!$project.SelectSingleNode("//*[local-name()='Target' and @Name='RefreshEasySetup']")){throw 'Automatic signed Release packaging hook missing.'}
$vswhere=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe';$vs=& $vswhere -latest -products '*' -property installationPath
$output=Join-Path $repo 'artifacts/easySetup-tests';New-Item -ItemType Directory -Force $output | Out-Null
& (Join-Path $vs 'MSBuild/Current/Bin/Roslyn/csc.exe') /nologo /target:exe /main:SetupPolicyTests /langversion:7.3 /r:System.Windows.Forms.dll /r:System.Core.dll /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll (Join-Path $PSScriptRoot 'Bootstrapper.cs') (Join-Path $PSScriptRoot 'SetupPolicyTests.cs') "/out:$output/SetupPolicyTests.exe"
if($LASTEXITCODE){throw 'Setup policy compilation failed.'}
& (Join-Path $output 'SetupPolicyTests.exe');if($LASTEXITCODE){throw 'Setup policy tests failed.'}
if($SourceOnly){Write-Output 'PASS packaging source checks; signed installer generation and installation not validated.';return}
$exe=Join-Path $PSScriptRoot 'PowerPointIntelligenceSetup.exe';$receipt=Join-Path $PSScriptRoot 'build.json'
if(!(Test-Path $exe) -or !(Test-Path $receipt)){throw 'Installer is not built. Run easySetup/Build-EasySetup.ps1.'}
$build=Get-Content $receipt -Raw | ConvertFrom-Json
if($build.SourceSha256 -ne (Get-EasySetupFingerprint $repo)){throw 'Installer is stale relative to current plugin, template or packaging sources.'}
if($build.ExeSha256 -ne (Get-FileHash $exe -Algorithm SHA256).Hash){throw 'Installer checksum mismatch.'}
$assembly=[Reflection.Assembly]::LoadFile($exe)
$stream=$assembly.GetManifestResourceStream('IntelligencePayload.zip');if(!$stream){throw 'Embedded payload resource missing.'}
$sha=[Security.Cryptography.SHA256]::Create();try{$actualBundle=[BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-','')}finally{$sha.Dispose();$stream.Dispose()}
if($actualBundle -ne $build.BundleSha256){throw 'Embedded payload checksum differs from build receipt.'}
$signature=Get-AuthenticodeSignature $exe
if($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Thumbprint -ne $build.CertificateThumbprint){throw 'Installer signature/trust validation failed.'}
Write-Output "PASS fresh signed installer $($build.Version); no installation was performed."
