$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'Piidx-TestSuite.ps1')
$suite=Import-PpiPiidxSuite $repo
$output=Join-Path $repo ('artifacts/piidx/schema-'+[Guid]::NewGuid().ToString('N'))
$result=[PiidxSchemaTests]::Run($output)
Assert-PpiPiidxSources $repo $suite.Sources
[pscustomobject]@{Status='PASS';Result=$result;OS=[Environment]::OSVersion.VersionString;Process64=[Environment]::Is64BitProcess;Runtime=[Environment]::Version.ToString();OfficeExecuted=$false;Sources=$suite.Sources;AssemblySHA256=$suite.AssemblySHA256;InputSHA256=$suite.InputSHA256} | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $output 'result.json')
$result
"Evidence: $output"
