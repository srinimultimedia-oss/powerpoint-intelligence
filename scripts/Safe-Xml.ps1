function Read-PpiSafeXml([string]$Path) {
    $settings=[Xml.XmlReaderSettings]::new()
    $settings.DtdProcessing=[Xml.DtdProcessing]::Prohibit
    $settings.XmlResolver=$null
    $settings.MaxCharactersInDocument=16777216
    $settings.MaxCharactersFromEntities=1024
    $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
    try {
        if($stream.Length -gt 16777216){throw 'XML input exceeds 16 MiB.'}
        $reader=[Xml.XmlReader]::Create($stream,$settings)
        try {
            $document=[Xml.XmlDocument]::new()
            $document.XmlResolver=$null
            $document.PreserveWhitespace=$true
            $document.Load($reader)
            return ,$document
        } finally {$reader.Dispose()}
    } finally {$stream.Dispose()}
}
