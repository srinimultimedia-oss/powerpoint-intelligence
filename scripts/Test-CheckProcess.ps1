$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'Check-Process.ps1')
$folder=Join-Path ([IO.Path]::GetTempPath()) ('ppi-process-check-'+[Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $folder | Out-Null
$shell=(Get-Process -Id $PID).Path
$log=Join-Path $folder 'streams.log'
$result=Invoke-PpiCheckProcess $shell @('-NoProfile','-Command','1..4000 | ForEach-Object { [Console]::Out.WriteLine("output-"+$_); [Console]::Error.WriteLine("error-"+$_) }; exit 7') $log 30
$content=Get-Content $log -Raw
if($result.ExitCode -ne 7 -or $result.TimedOut -or $content -notmatch 'output-4000' -or $content -notmatch 'error-4000'){throw 'FAIL simultaneous large streams or nonzero exit propagation'}
Write-Output 'PASS large stdout/stderr streams complete without pipe deadlock and retain exit code'
$result=Invoke-PpiCheckProcess $shell @('-NoProfile','-Command','[Console]::Out.WriteLine("before-timeout"); Start-Sleep -Seconds 30') (Join-Path $folder 'timeout.log') 1
if(!$result.TimedOut -or $result.ExitCode -eq 0 -or $result.Seconds -gt 10){throw 'FAIL bounded owned-process timeout'}
Write-Output 'PASS bounded timeout stops the owned check and reports failure'
$result=Invoke-PpiCheckProcess $shell @('-NoProfile','-Command','[Console]::Out.WriteLine("a value with spaces and quotes"); exit 0') (Join-Path $folder 'arguments.log') 20
if($result.ExitCode -ne 0 -or (Get-Content (Join-Path $folder 'arguments.log') -Raw) -notmatch 'a value with spaces and quotes'){throw 'FAIL command-line argument preservation'}
Write-Output 'PASS argument quoting preserves spaces and embedded quotes'
