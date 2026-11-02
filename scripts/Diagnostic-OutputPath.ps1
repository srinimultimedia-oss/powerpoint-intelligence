function Assert-PpiDiagnosticDirectPath([string]$Path) {
    # GetAttributes also checks a dangling link; Exists alone can miss one.
    $current=[IO.Path]::GetFullPath($Path)
    while($current) {
        try {
            $attributes=[IO.File]::GetAttributes($current)
            if(($attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw [IO.InvalidDataException]::new('Diagnostic output cannot traverse reparse points.')
            }
        } catch [IO.FileNotFoundException] {
            # A new output path is permitted; inspect its existing parents.
        } catch [IO.DirectoryNotFoundException] {
        }
        $current=[IO.Path]::GetDirectoryName($current)
    }
}

function Get-PpiDiagnosticOutputDirectory([string]$Requested,[string]$RepositoryRoot) {
    if([string]::IsNullOrWhiteSpace($Requested) -or $Requested.Length -gt 32767 -or
        $Requested -match '\p{Cc}' -or $Requested.StartsWith('\\?\') -or $Requested.StartsWith('\\.\')) {
        throw [IO.InvalidDataException]::new('Invalid diagnostic output path.')
    }
    $parts=$Requested.Replace('/','\').Split('\')
    for($i=0;$i -lt $parts.Length;$i++) {
        $part=$parts[$i]
        if(!$part){continue}
        $drive=$i -eq 0 -and $part -match '^[A-Za-z]:$'
        if($part -in @('.','..') -or $part.EndsWith('.') -or $part.EndsWith(' ') -or
            $part -match '[<>"|?*]' -or (!$drive -and $part.Contains(':'))) {
            throw [IO.InvalidDataException]::new('Ambiguous or traversing diagnostic output component.')
        }
        $stem=$part.Split('.')[0].TrimEnd(' ','.')
        if($stem -match '^(CON|PRN|AUX|NUL|COM[1-9¹²³]|LPT[1-9¹²³])$') {
            throw [IO.InvalidDataException]::new('Device names are not diagnostic output directories.')
        }
    }
    if($Requested -match '^[A-Za-z]:(?![\\/])' -or $Requested -match '^[/\\](?![/\\])') {
        throw [IO.InvalidDataException]::new('Drive-relative diagnostic output is unsupported.')
    }
    $repository=[IO.Path]::GetFullPath($RepositoryRoot)
    $candidate=[IO.Path]::GetFullPath($(if([IO.Path]::IsPathRooted($Requested)){$Requested}else{[IO.Path]::Combine($repository,$Requested)}))
    $artifacts=[IO.Path]::GetFullPath([IO.Path]::Combine($repository,'artifacts')).TrimEnd('\','/')
    $temporary=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\','/')
    foreach($root in @($artifacts,$temporary)) {
        $equal=[string]::Equals($candidate.TrimEnd('\','/'),$root,[StringComparison]::OrdinalIgnoreCase)
        if($equal -and $root -ne $artifacts){continue}
        if(!$candidate.StartsWith($root+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase) -and !($equal -and $root -eq $artifacts)){continue}
        # The argument selects a trusted data root, but cannot become that root.
        $relative=$candidate.Substring($root.Length).TrimStart('\','/')
        $safe=$root
        foreach($component in $relative.Split([char[]]'\/',[StringSplitOptions]::RemoveEmptyEntries)) {
            $name=[IO.Path]::GetFileName($component)
            if($name -cne $component){throw [IO.InvalidDataException]::new('Invalid diagnostic filename component.')}
            $safe=[IO.Path]::Combine($safe,$name)
        }
        if(![string]::Equals($safe.TrimEnd('\','/'),$candidate.TrimEnd('\','/'),[StringComparison]::OrdinalIgnoreCase)) {
            throw [IO.InvalidDataException]::new('Diagnostic path reconstruction changed its location.')
        }
        Assert-PpiDiagnosticDirectPath $safe
        return $safe
    }
    throw [IO.InvalidDataException]::new('Diagnostic reports must be under repository artifacts or a temporary subdirectory.')
}

function Get-PpiDiagnosticReportPath([string]$Directory) {
    $path=[IO.Path]::Combine($Directory,'scene-threat-evidence.json')
    Assert-PpiDiagnosticDirectPath $path
    return $path
}
