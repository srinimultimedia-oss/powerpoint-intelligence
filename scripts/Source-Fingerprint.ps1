. (Join-Path $PSScriptRoot 'Build-Integrity.ps1')
function Get-PpiSourceFingerprint([string]$Root) {
    $records=New-Object System.Collections.Generic.List[string]
    foreach($folder in @('apps','src','scripts','tests','packages')) {
        $path=Join-Path $Root $folder
        if(!(Test-Path $path)) { continue }
        Get-ChildItem $path -File -Recurse | Where-Object {
            $_.FullName.Substring($Root.Length+1) -notmatch '(^|[\\/])(bin|obj|node_modules|artifacts)[\\/]' -and $_.Extension -in @('.cs','.csproj','.ps1','.py','.json','.mjs','.ts','.tsx','.config','.xsd')
        } | Sort-Object FullName | ForEach-Object {
            $relative=$_.FullName.Substring($Root.Length+1).Replace('\','/')
            $records.Add($relative+'='+(Get-PpiContentHash $_.FullName))
        }
    }
    $hash=[Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes(($records -join "`n"))))).Replace('-','') }
    finally { $hash.Dispose() }
}
function Get-PpiAddinFingerprint([string]$Root) {
    $files=@(Get-ChildItem (Join-Path $Root 'apps/powerpoint-vsto') -File -Recurse | Where-Object { $_.FullName -notmatch '[\\/](bin|obj)[\\/]' -and $_.Extension -in @('.cs','.csproj','.config','.json') }; Get-ChildItem (Join-Path $Root 'src/Chart.Domain') -File; Get-ChildItem (Join-Path $Root 'src/Chart.Renderers') -File; Get-ChildItem (Join-Path $Root 'src/Shared.Contracts') -File; Get-ChildItem (Join-Path $Root 'src/Library.Core') -File) | Sort-Object FullName
    $lines=@($files | ForEach-Object { ($_.FullName.Substring($Root.Length+1))+'='+(Get-PpiContentHash $_.FullName) })
    $hash=[Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes(($lines -join "`n"))))).Replace('-','') }
    finally { $hash.Dispose() }
}
