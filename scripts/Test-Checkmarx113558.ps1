$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
. (Join-Path $repo 'scripts/Artifact-Inputs.ps1')
. (Join-Path $repo 'scripts/Safe-Xml.ps1')
. (Join-Path $repo 'scripts/Source-Fingerprint.ps1')
$root=Join-Path $repo ('artifacts/checkmarx-113558/'+[Guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Force $root
$count=0
function Check($ok,$name){if(!$ok){throw "FAIL $name"};$script:count++;Write-Output "PASS $name"}
function Reject([scriptblock]$action,$name){$failed=$false;try{& $action | Out-Null}catch{$failed=$true};Check $failed $name}
$json=Join-Path $root 'evidence.json';[IO.File]::WriteAllText($json,'{"ReadOnly":true}')
Check ((Read-PpiArtifactJsonText $json) -eq '{"ReadOnly":true}') 'valid artifact JSON preserved'
Reject {Read-PpiArtifactJsonText (Join-Path $repo 'AGENTS.md')} 'outside artifacts rejected'
Reject {Read-PpiArtifactJsonText (Join-Path $root '../evidence.json')} 'parent traversal rejected'
Reject {Read-PpiArtifactJsonText ($json+':stream')} 'alternate stream rejected'
Reject {Read-PpiArtifactJsonText $json 2} 'bounded JSON read'
Reject {Read-PpiArtifactJsonText (Join-Path $root 'missing.json')} 'missing JSON rejected'
$xml=Join-Path $root 'manifest.xml'
[IO.File]::WriteAllText($xml,'<assembly xmlns="urn:example"><assemblyIdentity version="1.0.0.12"/></assembly>')
$doc=Read-PpiSafeXml $xml
Check ($doc.SelectSingleNode("/*[local-name()='assembly']/*[local-name()='assemblyIdentity']").GetAttribute('version') -eq '1.0.0.12') 'namespaced manifest version preserved'
[IO.File]::WriteAllText($xml,'<configuration><appSettings><add key="TemplateLibraryFolder" value="local"/></appSettings></configuration>')
$doc=Read-PpiSafeXml $xml;$doc.SelectSingleNode('/configuration/appSettings/add').SetAttribute('value','templates');$doc.Save($xml)
Check ((Read-PpiSafeXml $xml).SelectSingleNode('/configuration/appSettings/add').GetAttribute('value') -eq 'templates') 'config edit and save preserved'
foreach($bad in @('<!DOCTYPE x [<!ENTITY e SYSTEM "file:///C:/Windows/win.ini">]><x>&e;</x>','<!DOCTYPE x SYSTEM "https://invalid.example/test.dtd"><x/>','<!DOCTYPE x [<!ENTITY a "x"><!ENTITY b "&a;&a;">]><x>&b;</x>','<x>')){
    [IO.File]::WriteAllText($xml,$bad);Reject {Read-PpiSafeXml $xml} 'DTD/entity/malformed XML rejected'
}
[IO.File]::WriteAllText($xml,('x'*16777217));Reject {Read-PpiSafeXml $xml} 'oversize XML rejected'
Reject {Get-PpiFingerprintFileHash $root (Join-Path $repo 'AGENTS.md')} 'fingerprint outside selected root rejected'
Check ((Get-PpiFingerprintFileHash $root $json) -eq (Get-FileHash -LiteralPath $json -Algorithm SHA256).Hash) 'bounded local fingerprint hash unchanged'
[IO.File]::WriteAllText($xml,'<fixture/>')
$junction=Join-Path $root 'linked'
try{$null=New-Item -ItemType Junction -Path $junction -Target $root -ErrorAction Stop}catch{Write-Output ('NOT RUN junction fixture: '+$_.Exception.GetType().Name)}
if(Test-Path -LiteralPath $junction){
Reject {Read-PpiArtifactJsonText (Join-Path $junction 'evidence.json')} 'junction JSON input rejected'
Reject {Get-PpiFingerprintFileHash $root (Join-Path $junction 'evidence.json')} 'junction fingerprint input rejected'
}
Write-Output "$count Checkmarx regression assertions passed; inert data only, no Office/signing/network operations."
