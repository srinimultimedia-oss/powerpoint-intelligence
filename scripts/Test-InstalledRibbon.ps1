$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'Source-Fingerprint.ps1')
$report=Join-Path $repo 'artifacts/installed-ribbon-status.json'
$fixture=$null
$result=[ordered]@{Status='FAIL';LoadedBuild=$null;RibbonOnLoad=$false;ToolBindings=0;AssemblySha256=$null;Error=$null}
try {
    $app=New-Object -ComObject PowerPoint.Application
    $app.Visible=-1
    if($app.Presentations.Count -gt 0) { throw 'Save and close user presentations before runtime validation. No presentation was modified.' }
    $fixture=$app.Presentations.Add(-1)
    $fixture.Slides.Add(1,12).Select()
    $matches=@($app.COMAddIns | Where-Object { ($_.ProgId+' '+$_.Description) -match 'PowerPointIntelligence|PowerPoint Intelligence' })
    if($matches.Count -ne 1) { throw "Expected one registered PowerPoint Intelligence add-in; found $($matches.Count)." }
    if(!$matches[0].Connect) { throw 'Add-in is registered but not connected. Check COM Add-ins load error.' }
    $api=$matches[0].Object
    $result.LoadedBuild=$api.GetBuildInfo()
    $ribbonSource=[IO.File]::ReadAllText((Join-Path $repo 'apps/powerpoint-vsto/DemoRibbon.cs'))
    $expectedMarker=[regex]::Match($ribbonSource,'BuildMarker = "([^"]+)"').Groups[1].Value
    if(!$expectedMarker -or $result.LoadedBuild -notmatch ([regex]::Escape($expectedMarker)+'(?:;|$)')) { throw 'PowerPoint loaded an outdated add-in.' }
    $result.RibbonOnLoad=$api.GetRibbonLoaded()
    if(!$result.RibbonOnLoad) { throw 'PowerPoint did not invoke the ribbon OnLoad callback. Ribbon rendering is not verified.' }
    [xml]$xml=$api.GetRibbonXml()
    $result.ToolBindings=$xml.SelectNodes("//*[local-name()='button' and @onAction='RunDemoTool']").Count
    if($result.ToolBindings -ne 90) { throw 'Loaded ribbon does not expose all 90 tools.' }
    if($xml.SelectNodes("//*[local-name()='button' and @id='ppiRenderShapes' and @onAction='InsertRenderedShapes']").Count -ne 1) { throw 'Loaded ribbon is missing the native renderer sample command.' }
    if($xml.SelectNodes("//*[local-name()='menu' and @id='ppiDataSource']/*[local-name()='button' and @onAction='LinkDataSource']").Count -ne 4) { throw 'Loaded ribbon is missing the four Data Source options.' }
    if($xml.SelectNodes("//*[local-name()='button' and @id='ppiTemplateLibrary' and @onAction='ShowTemplateLibrary']").Count -ne 1) { throw 'Loaded ribbon is missing the Library command.' }
    $groups=@($xml.SelectNodes("//*[local-name()='tab' and @id='ppiTab']/*[local-name()='group']") | ForEach-Object {$_.label})
    if(($groups -join ',') -ne 'Structure,Elements,AI,Data,Text,Layout,Color,Quality,Collaboration,More') { throw 'Installed workflow group order is outdated.' }
    foreach($workflow in @('agenda-insert','agenda-contents','agenda-settings','agenda-split','agenda-refresh','annotation-refresh','percent-area','connections','text-left','text-center','text-right','text-justify','map-excel','relink-selected','data-manager','axis-settings','match-scales','series-settings','quality-review','text-top','text-middle','text-bottom','text-horizontal','text-downward','text-upward','color-fill','color-outline','color-text','export-pdf')) {
        if($xml.SelectNodes("//*[local-name()='button' and @tag='$workflow' and @onAction='RunWorkflowCommand' and @getEnabled='GetWorkflowEnabled']").Count -ne 1) { throw "Installed workflow binding missing: $workflow" }
    }
    foreach($elementSource in @('ElementCommands.cs','ProductivityTools.cs')) {
        $commandsSource=[IO.File]::ReadAllText((Join-Path $repo ('apps/powerpoint-vsto/'+$elementSource)))
        $literal=[regex]::Match($commandsSource,'string\[\] Ids=\{([^}]+)\}').Groups[1].Value
        $commandIds=@([regex]::Matches($literal,'"([^"]+)"') | ForEach-Object {$_.Groups[1].Value})
        if(!$commandIds.Count){throw "Cannot identify current command set: $elementSource"}
        $callback=if($elementSource -eq 'ElementCommands.cs'){'RunElementCommand'}else{'RunProductivityTool'}
        foreach($commandId in $commandIds) {
            if($xml.SelectNodes("//*[local-name()='button' and @tag='$commandId' and @onAction='$callback']").Count -ne 1){throw "Installed element/productivity binding missing: $commandId"}
        }
    }
    if($xml.SelectNodes("//*[local-name()='button' and @id='ppiContextProperties' and @onAction='ShowContextProperties']").Count -ne 1){throw 'Loaded ribbon lacks the selection-aware Properties command.'}
    $charts=$xml.SelectSingleNode("//*[local-name()='menu' and @id='ppiCategoryCharts' and @size='large' and @image='ppi_charts-gallery_32']")
    if(!$charts -or $charts.PreviousSibling.id -ne 'ppiTemplateLibrary'){throw 'Installed ribbon lacks the large Charts menu immediately beside Library.'}
    [xml]$contextState=$api.GetContextPropertiesState()
    if($contextState.DocumentElement.LocalName -ne 'properties'){throw 'Installed contextual property inspection API is unavailable.'}
    $groupIcons=@($xml.SelectNodes("//*[local-name()='tab' and @id='ppiTab']/*[local-name()='group']"))
    if($groupIcons.Count -ne 10 -or @($groupIcons | Where-Object {$_.image -notmatch '^ppi_.+_32$'}).Count -ne 0) { throw 'Installed collapsed groups lack their custom icons.' }
    $loadedPath=($result.LoadedBuild -split '; assembly=',2)[1]
    if(!$loadedPath -or !(Test-Path $loadedPath)) { throw 'Cannot inspect loaded assembly path.' }
    $result.AssemblySha256=(Get-FileHash $loadedPath -Algorithm SHA256).Hash
    $receiptPath=Join-Path $repo 'artifacts/installed-demo-receipt.json'
    if(!(Test-Path $receiptPath)) { throw 'Install receipt missing; rebuild/install using Install-Local.ps1.' }
    $receipt=Get-Content $receiptPath -Raw | ConvertFrom-Json
    if(!$receipt.SourceSha256 -or $receipt.SourceSha256 -ne (Get-PpiAddinFingerprint $repo)) { throw 'Installed package does not match the tested add-in source fingerprint. Rebuild/install the current source.' }
    if($receipt.AssemblySha256 -ne $result.AssemblySha256) { throw 'PowerPoint loaded a different DLL than the latest installed package.' }
    $result.Status='PASS'
    Write-Output 'PASS installed build identity, assembly hash, ribbon OnLoad and all 90 command bindings. Visual layout and command behavior require separate checks.'
} catch {
    $result.Error=$_.Exception.GetType().Name+': '+$_.Exception.Message
    throw
} finally {
    try { if($fixture) { $fixture.Saved=-1; $fixture.Close() } }
    catch { Write-Warning 'Could not close this test fixture; no user presentation was targeted.' }
    $result | ConvertTo-Json | Set-Content $report
}
