param([string]$CertificateThumbprint,[ValidateRange(2,300)][int]$PollSeconds=5)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'Common.ps1')
$built=$null
Write-Output 'Watching plugin and packaging sources. Ctrl+C stops watching. Signing and normal Windows security policies apply.'
while($true){
    $current=Get-EasySetupFingerprint $repo
    if($current -ne $built){
        Start-Sleep -Seconds $PollSeconds
        if($current -ne (Get-EasySetupFingerprint $repo)){continue}
        try{& (Join-Path $PSScriptRoot 'Build-EasySetup.ps1') -CertificateThumbprint $CertificateThumbprint;$built=$current}
        catch{Write-Warning $_.Exception.Message;Write-Warning 'Previous installer is unchanged. Waiting for another source edit before retrying.';$built=$current}
    }
    Start-Sleep -Seconds $PollSeconds
}
