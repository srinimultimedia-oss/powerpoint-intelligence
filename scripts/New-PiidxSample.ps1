$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'Piidx-TestSuite.ps1')
$suite=Import-PpiPiidxSuite $repo
$source=Join-Path $repo 'artifacts/demo-features-97c75dce36974d3ea08640c33105af95/context-properties-native.pptx'
if(!(Test-Path -LiteralPath $source)){throw 'Previously generated native test source is unavailable; no replacement or native acceptance is fabricated.'}
$output=Join-Path $repo ('artifacts/piidx/sample-native-'+[Guid]::NewGuid().ToString('N'))
$result=[PiidxSampleBuilder]::Build($output,$source)
Assert-PpiPiidxSources $repo $suite.Sources
[pscustomobject]@{Status='PASS';Result=$result;Source=$source;SourceSHA256=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash;IndexSHA256=(Get-FileHash -LiteralPath (Join-Path $output 'sample-native.piidx') -Algorithm SHA256).Hash;OfficeExecutedNow=$false;AssemblySHA256=$suite.AssemblySHA256;Sources=$suite.Sources} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $output 'evidence.json')
$result
"Evidence: $output"