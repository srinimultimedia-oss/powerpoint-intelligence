$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
$files=@('src/Library.Core/PresentationSource.cs','tests/library/piidx/PiidxSourceTests.cs') | ForEach-Object {Join-Path $repo $_}
Add-Type -Path $files
$output=Join-Path $repo ('artifacts/piidx/source-'+[Guid]::NewGuid().ToString('N'))
$result=[PiidxSourceTests]::Run($output,(Join-Path $repo 'apps/powerpoint-vsto/templates'))
[pscustomobject]@{Status='PASS';Result=$result;Process64=[Environment]::Is64BitProcess;Runtime=[Environment]::Version.ToString();OfficeExecuted=$false;Sources=@($files | ForEach-Object {[pscustomobject]@{File=$_.Substring($repo.Length+1);SHA256=(Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash}})} | ConvertTo-Json -Depth 4 | Set-Content (Join-Path $output 'result.json')
$result
"Evidence: $output"