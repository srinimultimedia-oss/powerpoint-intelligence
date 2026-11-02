param([ValidateSet('x86','x64')][string]$Platform='x64',[string]$OutputDirectory)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'Package-Content.ps1')
$repo=Split-Path $PSScriptRoot -Parent
$output=if($OutputDirectory){[IO.Path]::GetFullPath($OutputDirectory)}else{Join-Path $repo "apps/powerpoint-vsto/bin/$Platform/Release"}
$assembly=Join-Path $output 'PowerPointIntelligence.dll'
$appManifest=Join-Path $output 'PowerPointIntelligence.dll.manifest'
$deployment=Join-Path $output 'PowerPointIntelligence.vsto'
foreach($file in @($assembly,$appManifest,$deployment)) { if(!(Test-Path $file)){throw "Package file missing: $file"} }
$source=Get-Content (Join-Path $repo 'apps/powerpoint-vsto/DemoRibbon.cs') -Raw
$marker=[regex]::Match($source,'BuildMarker = "([^"]+)"').Groups[1].Value
if(!(Test-PpiPayloadMarker -Bytes ([IO.File]::ReadAllBytes($assembly)) -Marker $marker)) {
    throw 'STALE PACKAGE: deployed DLL does not contain the current demo ribbon. Run Install-Local.ps1 WITHOUT -SkipBuild.'
}
$newest=Get-ChildItem (Join-Path $repo 'apps/powerpoint-vsto') -File | Where-Object { $_.Extension -in @('.cs','.csproj','.config') } | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
if($newest.LastWriteTimeUtc -gt (Get-Item $assembly).LastWriteTimeUtc) { throw "STALE PACKAGE: source $($newest.Name) is newer than deployed DLL. Rebuild required." }
function Assert-IdentityHash([xml]$Manifest,[string]$Codebase) {
    $dependency=$Manifest.SelectSingleNode("//*[local-name()='dependentAssembly' and @codebase='$Codebase']")
    if(!$dependency) { throw "Manifest dependency missing: $Codebase" }
    $method=$dependency.SelectSingleNode(".//*[local-name()='DigestMethod']").GetAttribute('Algorithm')
    if($method -notmatch 'sha256$') { throw "Unsupported package digest method: $method" }
    $expected=$dependency.SelectSingleNode(".//*[local-name()='DigestValue']").InnerText
    $hash=[Security.Cryptography.SHA256]::Create()
    try { $actual=[Convert]::ToBase64String($hash.ComputeHash([IO.File]::ReadAllBytes((Join-Path $output $Codebase)))) } finally { $hash.Dispose() }
    if($actual -ne $expected) { throw "Package hash mismatch: $Codebase. Rebuild signed manifests required." }
}
[xml]$applicationXml=Get-Content $appManifest -Raw
[xml]$deploymentXml=Get-Content $deployment -Raw
Assert-IdentityHash $applicationXml 'PowerPointIntelligence.dll'
Assert-IdentityHash $deploymentXml 'PowerPointIntelligence.dll.manifest'
foreach($xml in @($applicationXml,$deploymentXml)) { if(!$xml.SelectSingleNode("//*[local-name()='Signature']")) { throw 'Manifest signature missing.' } }
$version=$deploymentXml.SelectSingleNode("/*[local-name()='assembly']/*[local-name()='assemblyIdentity']").GetAttribute('version')
Write-Output "PASS current demo payload and manifest hashes: $marker; deployment=$version; platform=$Platform"
Write-Output 'Publisher trust, installed loading and visible ribbon are verified separately.'
