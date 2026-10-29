$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
$code=@('src/Library.Core/NativeSqlite.cs','src/Library.Core/PiidxSchema.cs','tests/library/piidx/PiidxSchemaTests.cs') | ForEach-Object {Join-Path $repo $_}
Add-Type -Path $code
$output=Join-Path $repo ('artifacts/piidx/schema-'+[Guid]::NewGuid().ToString('N'))
$result=[PiidxSchemaTests]::Run($output)
[pscustomobject]@{Status='PASS';Result=$result;OS=[Environment]::OSVersion.VersionString;Process64=[Environment]::Is64BitProcess;Runtime=[Environment]::Version.ToString();OfficeExecuted=$false;Sources=@($code | ForEach-Object {[pscustomobject]@{File=$_.Substring($repo.Length+1);SHA256=(Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash}})} | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $output 'result.json')
$result
"Evidence: $output"