# Runs a console check with file-backed streams: no pipe buffer can deadlock the parent.
function ConvertTo-PpiNativeArgument([string]$Value) {
    if($null -eq $Value){$Value=''}
    return '"'+([regex]::Replace(([regex]::Replace($Value,'(\\*)"','$1$1\"')),'(\\+)$','$1$1'))+'"'
}
function Invoke-PpiCheckProcess([string]$Executable,[string[]]$Arguments,[string]$Log,[int]$TimeoutSeconds=300) {
    $stdout=$Log+'.stdout';$stderr=$Log+'.stderr'
    $line=($Arguments | ForEach-Object {ConvertTo-PpiNativeArgument $_}) -join ' '
    $started=[DateTime]::UtcNow
    $process=Start-Process -FilePath $Executable -ArgumentList $line -RedirectStandardOutput $stdout -RedirectStandardError $stderr -WindowStyle Hidden -PassThru
    # Retain the native handle before waiting: Windows PowerShell Start-Process can otherwise lose the exit code after the process exits.
    $nativeHandle=$process.Handle
    $timeout=$false;$code=1
    try {
        if(!$process.WaitForExit($TimeoutSeconds*1000)){$timeout=$true;$process.Kill();$null=$process.WaitForExit(5000)}
        if(!$timeout){$process.WaitForExit();$code=$process.ExitCode;if($null -eq $code){throw 'The owned check exited but its exit code is unavailable. No PASS result is inferred.'}}
    } finally {
        # A timeout stops only this owned check process. Never kill user Office processes.
        "Started UTC: $($started.ToString('o')); owned PID: $($process.Id); timeout seconds: $TimeoutSeconds" | Set-Content -LiteralPath $Log
        foreach($stream in @(@{Name='stdout';Path=$stdout},@{Name='stderr';Path=$stderr})){
            "--- $($stream.Name) ---" | Add-Content -LiteralPath $Log
            if(Test-Path -LiteralPath $stream.Path){Get-Content -LiteralPath $stream.Path -ErrorAction Continue | Add-Content -LiteralPath $Log}
        }
        if($timeout){'TIMEOUT: owned check process stopped. Descendant/Office cleanup is not inferred; remaining suite checks must stop.' | Add-Content -LiteralPath $Log}
        $process.Dispose()
    }
    return [pscustomobject]@{ExitCode=$code;TimedOut=$timeout;Seconds=([DateTime]::UtcNow-$started).TotalSeconds}
}
