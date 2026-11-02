function Test-PpiPayloadMarker {
    param([byte[]]$Bytes, [string]$Marker)
    if (!$Marker -or !$Bytes) { return $false }
    # CLR string heaps can start a UTF-16 literal at either byte alignment.
    # Search exact bytes instead of decoding the entire PE at a fixed alignment.
    $needle = [Text.Encoding]::Unicode.GetBytes($Marker)
    for ($index = 0; $index -le $Bytes.Length - $needle.Length; $index++) {
        if ($Bytes[$index] -ne $needle[0]) { continue }
        $matches = $true
        for ($part = 1; $part -lt $needle.Length; $part++) {
            if ($Bytes[$index + $part] -ne $needle[$part]) { $matches = $false; break }
        }
        if ($matches) { return $true }
    }
    return $false
}
