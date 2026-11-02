. (Join-Path (Split-Path $PSScriptRoot -Parent) 'scripts/Source-Fingerprint.ps1')
function Get-EasySetupFingerprint([string]$Root){
    $records=@((Get-PpiAddinFingerprint $Root))
    $files=@(Get-ChildItem (Join-Path $Root 'easySetup') -File | Where-Object {$_.Extension -in @('.ps1','.cs')};Get-ChildItem (Join-Path $Root 'apps/powerpoint-vsto/templates') -File -Recurse -Filter '*.potx') | Sort-Object FullName
    $records+=@($files | ForEach-Object {$_.FullName.Substring($Root.Length+1).Replace('\','/')+'='+(Get-FileHash $_.FullName -Algorithm SHA256).Hash})
    $hash=[Security.Cryptography.SHA256]::Create();try{return [BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes(($records -join "`n")))).Replace('-','')}finally{$hash.Dispose()}
}
