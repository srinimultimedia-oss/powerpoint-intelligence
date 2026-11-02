param([switch]$InteractiveOffice,[switch]$AgendaOnly,[switch]$ElementsOnly,[switch]$PropertiesOnly,[switch]$LinkedChartDiagnostic,[switch]$PauseOwnAddIn,[switch]$MappedExcelOnly)
$ErrorActionPreference='Stop'
if($MappedExcelOnly -and (!$InteractiveOffice -or $AgendaOnly -or $ElementsOnly -or $PropertiesOnly -or $LinkedChartDiagnostic -or $PauseOwnAddIn)){throw 'MappedExcelOnly requires InteractiveOffice and cannot be combined with other focused modes.'}
if($PauseOwnAddIn -and !$LinkedChartDiagnostic){throw 'PauseOwnAddIn is limited to LinkedChartDiagnostic.'}
if($LinkedChartDiagnostic -and (!$InteractiveOffice -or $AgendaOnly -or $ElementsOnly -or $PropertiesOnly)){throw 'LinkedChartDiagnostic requires InteractiveOffice and cannot be combined with other modes.'}
if($PropertiesOnly -and ($AgendaOnly -or $ElementsOnly)){throw 'PropertiesOnly cannot be combined with other focused modes.'}
if($PropertiesOnly -and !$InteractiveOffice){throw 'PropertiesOnly requires InteractiveOffice.'}
if($AgendaOnly -and $ElementsOnly){throw 'Choose AgendaOnly or ElementsOnly, not both.'}
if($ElementsOnly -and !$InteractiveOffice){throw 'ElementsOnly requires InteractiveOffice; no Office process was started.'}
if($AgendaOnly -and !$InteractiveOffice){throw 'AgendaOnly requires InteractiveOffice; no Office process was started.'}
$repo=Split-Path $PSScriptRoot -Parent
Push-Location $repo
try {
    $vswhere=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
    $vsPath=& $vswhere -latest -products '*' -property installationPath
    $compiler=Join-Path $vsPath 'MSBuild/Current/Bin/Roslyn/csc.exe'
    $pia=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Shared/Visual Studio Tools for Office/PIA/Office15'
    New-Item -ItemType Directory -Force artifacts/demo-tests | Out-Null
    & $compiler /nologo /debug:full /langversion:7.3 /target:exe /main:DemoOfficeTests /out:artifacts/demo-tests/DemoOfficeTests.exe "/link:$pia/office.dll" "/link:$pia/Microsoft.Office.Interop.PowerPoint.dll" /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll /r:Microsoft.CSharp.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Runtime.Serialization.dll /r:System.Xml.Linq.dll apps/powerpoint-vsto/TemplateCatalog.cs apps/powerpoint-vsto/TemplateLibrary.cs apps/powerpoint-vsto/LocalLibraryGallery.cs apps/powerpoint-vsto/LocalLibraryWorkspace.cs src/Library.Core/*.cs apps/powerpoint-vsto/DemoRibbon.cs apps/powerpoint-vsto/RibbonIcons.cs apps/powerpoint-vsto/NativeFormatting.cs apps/powerpoint-vsto/PresentationQuality.cs apps/powerpoint-vsto/NativeSeriesControls.cs apps/powerpoint-vsto/NativeChartControls.cs apps/powerpoint-vsto/MappedLinkOperations.cs apps/powerpoint-vsto/ExcelRangeMapping.cs apps/powerpoint-vsto/ChartAnnotations.cs apps/powerpoint-vsto/ProductivityTools.cs apps/powerpoint-vsto/XmlPartUpdates.cs apps/powerpoint-vsto/ContextPanelModel.cs apps/powerpoint-vsto/ContextFormatting.cs apps/powerpoint-vsto/ContextChartFormatting.cs apps/powerpoint-vsto/ContextSelectionBridge.cs apps/powerpoint-vsto/ContextPropertiesView.cs apps/powerpoint-vsto/ContextPropertiesController.cs apps/powerpoint-vsto/ContextActionRouter.cs apps/powerpoint-vsto/ElementCommands.cs apps/powerpoint-vsto/IndicatorElements.cs apps/powerpoint-vsto/UniversalConnectors.cs apps/powerpoint-vsto/TextShapeElements.cs apps/powerpoint-vsto/ProcessFlowTools.cs apps/powerpoint-vsto/ChapterAgendaModel.cs apps/powerpoint-vsto/ChapterAgenda.cs apps/powerpoint-vsto/PresentationAgenda.cs apps/powerpoint-vsto/ServiceConnections.cs apps/powerpoint-vsto/DataLinkManager.cs apps/powerpoint-vsto/DemoTools.cs apps/powerpoint-vsto/DataSources.cs tests/s0/DataSourceTests.cs apps/powerpoint-vsto/DemoInputs.cs apps/powerpoint-vsto/DemoLinks.cs apps/powerpoint-vsto/FeatureInventory.cs apps/powerpoint-vsto/PowerPointHost.cs apps/powerpoint-vsto/Contracts.cs apps/powerpoint-vsto/MetadataCodec.cs tests/s0/DemoScenario.cs tests/s0/MappedExcelOfficeTests.cs tests/s0/ContextPropertiesOfficeTests.cs tests/s0/ContextFormattingOfficeTests.cs tests/s0/ProductivityToolOfficeTests.cs tests/s0/TextShapeOfficeTests.cs tests/s0/ProcessFlowOfficeTests.cs tests/s0/IndicatorOfficeTests.cs tests/s0/UniversalConnectorOfficeTests.cs tests/s0/ChapterAgendaOfficeTests.cs tests/s0/WorkflowOfficeTests.cs tests/s0/ContextForegroundPolicy.cs tests/s0/ContextForegroundPrompt.cs tests/s0/LinkedChartDiagnostic.cs tests/s0/DemoOfficeTests.cs
    if($LASTEXITCODE) { throw 'Demo feature test compilation failed.' }
    if($InteractiveOffice) {
        $output=Join-Path $repo ('artifacts/demo-features-'+[Guid]::NewGuid().ToString('N'))
        $mode=if($MappedExcelOnly){'--interactive-mapped-excel'}elseif($LinkedChartDiagnostic -and $PauseOwnAddIn){'--diagnose-linked-chart-paused'}elseif($LinkedChartDiagnostic){'--diagnose-linked-chart'}elseif($PropertiesOnly){'--interactive-properties'}elseif($AgendaOnly){'--interactive-agenda'}elseif($ElementsOnly){'--interactive-elements'}else{'--interactive-office'}
        $probeStarted=Get-Date
        & ./artifacts/demo-tests/DemoOfficeTests.exe $mode $output
        $nativeExit=$LASTEXITCODE
        if($LinkedChartDiagnostic){
            try {
                $eventReadErrors=@();$events=Get-WinEvent -FilterHashtable @{LogName='Application';Id=1000;StartTime=$probeStarted.AddMinutes(-30)} -ErrorAction SilentlyContinue -ErrorVariable eventReadErrors
                if(@($eventReadErrors | Where-Object {$_.FullyQualifiedErrorId -notmatch 'NoMatchingEventsFound'}).Count){throw $eventReadErrors[0]}
                $crashes=@(foreach($event in $events){$eventXml=[xml]$event.ToXml();$fields=@{};foreach($field in $eventXml.Event.EventData.Data){$fields[$field.Name]=$field.'#text'};if($fields.AppName -match '^(POWERPNT|EXCEL)\.EXE$'){[pscustomobject]@{Time=$event.TimeCreated;Application=$fields.AppName;Module=$fields.ModuleName;ExceptionCode=$fields.ExceptionCode}}})
                ConvertTo-Json -InputObject $crashes | Set-Content (Join-Path $output 'office-crashes.json')
            }catch{('Crash metadata unavailable: '+$_.Exception.GetType().Name)|Set-Content (Join-Path $output 'office-crashes-unavailable.txt')}
        }
        Write-Output "Per-feature evidence: $output"
        if($nativeExit) { throw 'Demo feature acceptance incomplete: inspect per-feature failure rows.' }
    } else { Write-Output 'Demo feature runner compiled. Office checks skipped without -InteractiveOffice.' }
} finally {Pop-Location}
