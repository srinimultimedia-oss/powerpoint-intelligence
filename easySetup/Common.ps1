. (Join-Path (Split-Path $PSScriptRoot -Parent) 'scripts/Source-Fingerprint.ps1')
function Get-EasySetupEmbeddedPayloadHash([byte[]]$AssemblyBytes) {
    # A path-loaded assembly may remain cached after an atomic EXE replacement.
    # Load the verified immutable bytes, never a previously loaded path identity.
    $assembly=[Reflection.Assembly]::Load($AssemblyBytes)
    $stream=$assembly.GetManifestResourceStream('IntelligencePayload.zip')
    if(!$stream){throw 'Embedded payload resource missing.'}
    $sha=[Security.Cryptography.SHA256]::Create()
    try{return [BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-','')}
    finally{$sha.Dispose();$stream.Dispose()}
}
function Get-EasySetupFingerprint([string]$Root){
    $records=@((Get-PpiAddinFingerprint $Root),(Get-PpiSourceFingerprint $Root))
    $files=@(Get-ChildItem (Join-Path $Root 'easySetup') -File | Where-Object {$_.Extension -in @('.ps1','.cs')};Get-ChildItem (Join-Path $Root 'apps/powerpoint-vsto/templates') -File -Recurse -Filter '*.potx') | Sort-Object FullName
    $records+=@($files | ForEach-Object {$_.FullName.Substring($Root.Length+1).Replace('\','/')+'='+(Get-FileHash $_.FullName -Algorithm SHA256).Hash})
    $hash=[Security.Cryptography.SHA256]::Create();try{return [BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes(($records -join "`n")))).Replace('-','')}finally{$hash.Dispose()}
}

function New-EasySetupPayloadZip([string]$Payload,[string]$ZipPath) {
    # Windows PowerShell 5.1 does not necessarily load the enum's assembly when
    # ZipFile's FileSystem assembly is loaded. Resolve both before any ZIP types.
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $payloadRoot=[IO.Path]::GetFullPath($Payload).TrimEnd('\','/')
    $archive=[IO.Compression.ZipFile]::Open($ZipPath,[IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach($file in (Get-ChildItem -LiteralPath $payloadRoot -Recurse -File | Sort-Object FullName)) {
            $name=$file.FullName.Substring($payloadRoot.Length+1).Replace('\','/')
            $null=[IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive,$file.FullName,$name,[IO.Compression.CompressionLevel]::Optimal)
        }
    } finally {$archive.Dispose()}
}
