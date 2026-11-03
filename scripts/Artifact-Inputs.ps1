# Recovery inputs are data under this checkout's artifacts, never arbitrary paths.
function Resolve-PpiArtifactInput([string]$Path,[switch]$AllowMissing) {
    if([string]::IsNullOrWhiteSpace($Path) -or $Path -match '(^|[\\/])\.{1,2}([\\/]|$)'){throw 'Invalid artifact input path.'}
    $root=[IO.Path]::GetFullPath((Join-Path (Split-Path $PSScriptRoot -Parent) 'artifacts')).TrimEnd('\','/')
    $full=[IO.Path]::GetFullPath($Path)
    if(!$full.StartsWith($root+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Recovery input must be inside this checkout artifacts directory.'}
    if($full.Substring($root.Length+1).Contains(':')){throw 'Alternate streams are not artifact inputs.'}
    $cursor=$full
    while($cursor){
        if([IO.File]::Exists($cursor) -or [IO.Directory]::Exists($cursor)){
            if(([IO.File]::GetAttributes($cursor) -band [IO.FileAttributes]::ReparsePoint) -ne 0){throw 'Reparse paths are not recovery inputs.'}
        }
        $cursor=[IO.Path]::GetDirectoryName($cursor)
    }
    if(!$AllowMissing -and ![IO.File]::Exists($full)){throw 'Artifact input file is missing.'}
    return $full
}
function Read-PpiArtifactJsonText([string]$Path,[long]$Limit=1048576) {
    $safe=Resolve-PpiArtifactInput $Path
    if([IO.Path]::GetExtension($safe) -ne '.json'){throw 'Recovery input must be a JSON file.'}
    $stream=[IO.File]::Open($safe,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
    try {
        if($stream.Length -gt $Limit){throw 'Artifact input exceeds review bounds.'}
        $reader=[IO.StreamReader]::new($stream,[Text.Encoding]::UTF8,$true)
        try{return $reader.ReadToEnd()}finally{$reader.Dispose()}
    } finally {$stream.Dispose()}
}
