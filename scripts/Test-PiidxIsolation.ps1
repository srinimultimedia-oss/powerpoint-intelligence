$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
$violations=New-Object System.Collections.Generic.List[string]
$roots=@('apps/powerpoint-vsto','src')
foreach($root in $roots){Get-ChildItem -LiteralPath (Join-Path $repo $root) -Recurse -File | Where-Object {$_.FullName -notmatch '[\\/](bin|obj|templates)[\\/]' -and $_.Extension -in @('.cs','.csproj','.config')} | ForEach-Object {
    $text=[IO.File]::ReadAllText($_.FullName)
    if($text -match '(?i)\.tcidx|tcidx|LocalLibraryCache|library-index-v1\.xml' -and $_.Name -ne 'LibraryConfiguration.cs'){$violations.Add($_.FullName+': obsolete indexing reference')}
    if($text -match '(?i)ProgramFiles[^\r\n]*think.?cell|Path\.Combine[^\r\n]*"think-cell"|(?:DllImport|Reference Include|GetTypeFromProgID)[^\r\n]*think.?cell'){$violations.Add($_.FullName+': vendor loading/path dependency')}
}}
[xml]$project=Get-Content -LiteralPath (Join-Path $repo 'apps/powerpoint-vsto/PowerPointIntelligence.csproj') -Raw
$ns=New-Object Xml.XmlNamespaceManager($project.NameTable);$ns.AddNamespace('m','http://schemas.microsoft.com/developer/msbuild/2003')
if(!$project.SelectSingleNode('//m:Compile[@Include="..\..\src\Library.Core\*.cs"]',$ns)){$violations.Add('PIIDX core is not in the production compilation graph')}
if($project.SelectNodes('//m:Compile|//m:EmbeddedResource|//m:Content',$ns) | Where-Object {$_.Include -match '(?i)research|LegacyLibraryCacheFixture|tcidx'}){$violations.Add('Research/legacy fixture included in deployment graph')}
if($violations.Count){$violations | Write-Output;throw 'PIIDX production dependency isolation failed.'}
'PASS production dependency isolation: no proprietary index access, vendor path/loading dependency or legacy cache; core and gallery compiled in the production graph. Descriptive inventory/history references are not runtime dependencies.'