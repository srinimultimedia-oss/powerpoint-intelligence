param([ValidateRange(1,14)][int]$Days=3, [string]$OutputDirectory, [ValidateSet('McAfee','Defender','All')][string]$Scanner='McAfee')
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
if(!$OutputDirectory) { $OutputDirectory=Join-Path $repo ('artifacts/scene-threat-'+(Get-Date -Format 'yyyyMMdd-HHmmss')) }
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$errors=New-Object System.Collections.Generic.List[object]
function Record-ReadFailure([string]$Source,$Record) {
    $errors.Add([pscustomobject]@{Source=$Source;Type=$Record.Exception.GetType().FullName;HResult=('0x{0:X8}' -f $Record.Exception.HResult);Message=$Record.Exception.Message})
}
$report=[ordered]@{Time=(Get-Date).ToString('o');Process64=[Environment]::Is64BitProcess;ReadOnly=$true;Scanner=$Scanner;Products=@();McAfeeDetections=@();DefenderStatus=$null;Detections=@();Events=@();BuildReceipts=@();Errors=@();Conclusion='UNVERIFIED: no scanner verdict is inferred from file disappearance.'}
if($Scanner -in @('McAfee','All')) {
    try {
        $report.Products=@(foreach($key in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall')) {
            Get-ItemProperty "$key\*" -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -match 'McAfee' } | Select-Object DisplayName,DisplayVersion,Publisher
        })
    } catch { Record-ReadFailure 'McAfee installed product metadata' $_ }
    $mcafeeRows=New-Object System.Collections.Generic.List[object]
    $mcafeeLog=Join-Path $env:ProgramData 'McAfee/wps/detection.log'
    try {
        foreach($line in @(Select-String -LiteralPath $mcafeeLog -Pattern '(?i)(SceneContractTests|TemplateLibraryTests)\.exe|scene-contracts-x64|template-library-x64' -ErrorAction Stop)) {
            try { $entry=$line.Line | ConvertFrom-Json -ErrorAction Stop }
            catch { Record-ReadFailure ("McAfee matching log line "+$line.LineNumber) ([Management.Automation.ErrorRecord]::new([IO.InvalidDataException]::new('Matching McAfee log entry is not valid JSON.'),'McAfeeInvalidJson',[Management.Automation.ErrorCategory]::InvalidData,$null)); continue }
            if([string]$entry.target_name -notmatch '(?i)(SceneContractTests|TemplateLibraryTests)\.exe|scene-contracts-x64|template-library-x64') { continue }
            $mcafeeRows.Add([pscustomobject]@{Log=$mcafeeLog;Line=$line.LineNumber;TimeUtc=$entry.timestamp;Target=$entry.target_name;SHA256=$entry.target_hash;Detection=$entry.detection_name;Action=$entry.final_result;Sensor=$entry.sensor;FinalDetectionSource=@($entry.all | Where-Object final_detection_source | Select-Object -ExpandProperty final_detection_source)})
        }
    } catch { Record-ReadFailure 'McAfee matching detection log' $_ }
    $report.McAfeeDetections=@($mcafeeRows.ToArray())
}
if($Scanner -in @('Defender','All')) {
try { $report.Products=@(Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct -ErrorAction Stop | Select-Object displayName,productState) }
catch { Record-ReadFailure 'SecurityCenter2 products' $_ }
try {
    if(!(Get-Command Get-MpThreatDetection -ErrorAction SilentlyContinue)) {
        Import-Module (Join-Path $env:windir 'System32/WindowsPowerShell/v1.0/Modules/Defender/Defender.psd1') -ErrorAction Stop
    }
    $report.DefenderStatus=Get-MpComputerStatus -ErrorAction Stop | Select-Object AMProductVersion,AMEngineVersion,AntivirusSignatureVersion,AntivirusSignatureLastUpdated,RealTimeProtectionEnabled
} catch { Record-ReadFailure 'Defender status/module' $_ }
try {
    $matchingDetections=@(Get-MpThreatDetection -ErrorAction Stop | Where-Object { ($_.Resources -join "`n") -match '(?i)(SceneContractTests|TemplateLibraryTests)\.exe|scene-contracts-x64|template-library-x64' })
    $threats=@(Get-MpThreat -ErrorAction Stop)
    $report.Detections=@($matchingDetections | ForEach-Object {
        $detection=$_
        $threat=$threats | Where-Object ThreatID -eq $detection.ThreatID | Select-Object -First 1
        [pscustomobject]@{ThreatID=$detection.ThreatID;ThreatName=$threat.ThreatName;SeverityID=$threat.SeverityID;InitialDetectionTime=$detection.InitialDetectionTime;LastThreatStatusChangeTime=$detection.LastThreatStatusChangeTime;ActionSuccess=$detection.ActionSuccess;ThreatStatusID=$detection.ThreatStatusID;CleaningActionID=$detection.CleaningActionID;Resources=@($detection.Resources | Where-Object { $_ -match '(?i)(SceneContractTests|TemplateLibraryTests)\.exe|scene-contracts-x64|template-library-x64' })}
    })
} catch { Record-ReadFailure 'Defender matching detections' $_ }
$eventRows=New-Object System.Collections.Generic.List[object]
foreach($channel in @(
    @{LogName='Microsoft-Windows-Windows Defender/Operational';Id=@(1116,1117,1118,1119,1121,1122)},
    @{LogName='Microsoft-Windows-CodeIntegrity/Operational';Id=@(3033,3034,3076,3077,3089)}
)) {
    try {
        $filter=@{LogName=$channel.LogName;Id=$channel.Id;StartTime=(Get-Date).AddDays(-$Days)}
        foreach($event in @(Get-WinEvent -FilterHashtable $filter -MaxEvents 1000 -ErrorAction Stop)) {
            $xml=[xml]$event.ToXml()
            if($xml.OuterXml -notmatch '(?i)(SceneContractTests|TemplateLibraryTests)\.exe|scene-contracts-x64|template-library-x64') { continue }
            # Keep only named diagnostic fields. Never export unrelated events,
            # full event messages, command lines, account names or data contents.
            $fields=[ordered]@{}
            foreach($field in $xml.Event.EventData.Data) {
                $name=[string]$field.Name
                if($name -match '^(Threat ?Name|Threat ?ID|Detection ?ID|Detection ?Time|Path|FileName|FileNameBuffer|Process ?Name|ProcessNameBuffer|Action ?Name|Action ?ID|Error ?Code|Error ?Description|Severity ?Name|Category ?Name|Source ?Name|Detection ?Source|Engine ?Version|Product ?Version|Security intelligence ?Version|RequestedSigningLevel|ValidatedSigningLevel|Status)$') {
                    $fields[$name]=[string]$field.InnerText
                }
            }
            $eventRows.Add([pscustomobject]@{Time=$event.TimeCreated;EventID=$event.Id;RecordID=$event.RecordId;Channel=$channel.LogName;Fields=$fields})
        }
    } catch {
        if($_.FullyQualifiedErrorId -notmatch 'NoMatchingEventsFound') { Record-ReadFailure $channel.LogName $_ }
    }
}
$report.Events=@($eventRows.ToArray())
}
$receipts=New-Object System.Collections.Generic.List[object]
$receiptFolders=@(Get-ChildItem (Join-Path $repo 'artifacts') -Directory -Filter 'scene-contracts-*')
foreach($run in @(Get-ChildItem (Join-Path $repo 'artifacts') -Directory -Filter 'validation-*' | Sort-Object Name -Descending | Select-Object -First 5)) {
    $snapshot=Join-Path $run.FullName 'source/artifacts/scene-contracts-x64'
    if(Test-Path -LiteralPath $snapshot) { $receiptFolders+=Get-Item -LiteralPath $snapshot }
}
foreach($folder in $receiptFolders) {
    $receipt=Join-Path $folder.FullName 'compiled.json'; $executable=Join-Path $folder.FullName 'SceneContractTests.exe'
    if(Test-Path -LiteralPath $receipt) {
        $receipts.Add([pscustomobject]@{Path=$receipt;Receipt=(Get-Content -LiteralPath $receipt -Raw | ConvertFrom-Json);ExecutablePresent=(Test-Path -LiteralPath $executable);CurrentSha256=$(if(Test-Path -LiteralPath $executable){(Get-FileHash -LiteralPath $executable -Algorithm SHA256).Hash})})
    }
}
$report.BuildReceipts=@($receipts.ToArray());$report.Errors=@($errors.ToArray())
if($report.McAfeeDetections.Count -gt 0 -or $report.Detections.Count -gt 0 -or $report.Events.Count -gt 0) { $report.Conclusion='Matching scanner/policy records collected. Review the exact threat name, affected path and action before deciding remediation. This is not a false-positive verdict or proof that protection is active.' }
$reportPath=Join-Path $OutputDirectory 'scene-threat-evidence.json'
$report | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $reportPath
Write-Output "Evidence: $reportPath"
Write-Output ("Matching McAfee detections={0}; Defender detections={1}; events={2}; read errors={3}" -f $report.McAfeeDetections.Count,$report.Detections.Count,$report.Events.Count,$report.Errors.Count)
Write-Output $report.Conclusion
if($report.Errors.Count -gt 0) { Write-Warning 'Some diagnostic sources were inaccessible. Run this read-only collector in an approved desktop session; it never elevates itself or changes protection settings.' }
