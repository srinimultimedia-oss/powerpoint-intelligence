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
& (Join-Path $PSScriptRoot 'Test-ResourceReader.ps1') -Compiler (Join-Path $vs 'MSBuild/Current/Bin/Roslyn/csc.exe')
if($SourceOnly){Write-Output 'PASS packaging source checks; signed installer generation and installation not validated.';return}
$exe=Join-Path $PSScriptRoot 'PowerPointIntelligenceSetup.exe';$receipt=Join-Path $PSScriptRoot 'build.json'
if(!(Test-Path $exe) -or !(Test-Path $receipt)){throw 'Installer is not built. Run easySetup/Build-EasySetup.ps1.'}
$validationLock=New-Object Threading.Mutex($false,'Local\PowerPointIntelligence.EasySetup.Build.v1')
$locked=$false
$file=$null
try{
    try{$locked=$validationLock.WaitOne(0)}catch [Threading.AbandonedMutexException]{$locked=$true}
    if(!$locked){throw 'An installer build is in progress. Retry validation after it publishes.'}
    $build=Get-Content $receipt -Raw | ConvertFrom-Json
    if($build.SourceSha256 -ne (Get-EasySetupFingerprint $repo)){throw 'Installer is stale relative to current plugin, template or packaging sources.'}
    # Deny writes/replacement during hash/signature/resource checks. Never load
    # a path-cached assembly, and verify publisher trust before loading metadata.
    $file=[IO.File]::Open($exe,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
    if($file.Length -gt 1073741824){throw 'Installer exceeds the validation size bound.'}
    $memory=New-Object IO.MemoryStream
    try{$file.CopyTo($memory);$bytes=$memory.ToArray()}finally{$memory.Dispose()}
    $sha=[Security.Cryptography.SHA256]::Create()
    try{$actualExe=[BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-','')}finally{$sha.Dispose()}
    if($build.ExeSha256 -ne $actualExe){throw 'Installer checksum mismatch.'}
    $signature=Get-AuthenticodeSignature $exe
    if($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Thumbprint -ne $build.CertificateThumbprint){throw 'Installer signature/trust validation failed.'}
    $actualBundle=Get-EasySetupEmbeddedPayloadHash $bytes
    if($actualBundle -ne $build.BundleSha256){throw 'Embedded payload checksum differs from build receipt.'}
    Write-Output "PASS fresh signed installer $($build.Version); no installation was performed."
}finally{
    if($file){$file.Dispose()}
    if($locked){$validationLock.ReleaseMutex()}
    $validationLock.Dispose()
}
