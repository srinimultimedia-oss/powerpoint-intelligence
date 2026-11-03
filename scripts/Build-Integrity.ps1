function Get-PpiContentHash([string]$Path) {
    # Fingerprinting is a data read, not code retrieval or execution. Hash one
    # locked stream so the name cannot select different bytes during the read.
    $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
    $sha=[Security.Cryptography.SHA256]::Create()
    try {[BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-','')}
    finally {$sha.Dispose();$stream.Dispose()}
}

function Assert-PpiMicrosoftBuildTool([string]$Path,[ValidateSet('csc.exe','MSBuild.exe','vswhere.exe')][string]$Name) {
    if(![string]::Equals([IO.Path]::GetFileName($Path),$Name,[StringComparison]::OrdinalIgnoreCase)){throw 'Unexpected build tool filename.'}
    $signature=Get-AuthenticodeSignature -LiteralPath $Path -ErrorAction Stop
    if($signature.Status -ne 'Valid' -or !$signature.SignerCertificate -or
        $signature.SignerCertificate.Subject -notmatch '(^|,\s*)O=Microsoft Corporation(,|$)') {
        throw 'Build tool requires a valid Microsoft Authenticode signature; no trust policy was changed.'
    }
    $original=[Diagnostics.FileVersionInfo]::GetVersionInfo($Path).OriginalFilename
    if(![string]::Equals($original,$Name,[StringComparison]::OrdinalIgnoreCase)){throw 'Signed build tool identity differs from its expected role.'}
}

function Get-PpiVerifiedArtifact([string]$Path,[string]$RebuiltPath) {
    # RebuiltPath must come from the current build recipe, never a cached receipt.
    # Keep the original locked until its caller finishes loading/launching it.
    $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
    $sha=[Security.Cryptography.SHA256]::Create()
    try {
        if($stream.Length -gt 128MB){throw 'Executable artifact exceeds the verification bound.'}
        $hash=[BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-','')
        if($hash -ne (Get-PpiContentHash $RebuiltPath)){throw 'Cached code differs from the independently rebuilt current source. Execution refused; cache and receipt retained.'}
        $stream.Position=0
        $bytes=New-Object byte[] ([int]$stream.Length)
        $offset=0
        while($offset -lt $bytes.Length){$read=$stream.Read($bytes,$offset,$bytes.Length-$offset);if(!$read){throw 'Artifact ended during verification.'};$offset+=$read}
        [pscustomobject]@{Stream=$stream;Bytes=$bytes;SHA256=$hash;Path=$Path}
    } catch {$stream.Dispose();throw}
    finally {$sha.Dispose()}
}

function Assert-PpiSigningCertificate([Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,[string]$Thumbprint) {
    if($Thumbprint -notmatch '\A[0-9A-Fa-f]{40}\z'){throw 'Certificate thumbprint must be exactly 40 hexadecimal characters.'}
    if(!$Certificate -or ![string]::Equals($Certificate.Thumbprint,$Thumbprint,[StringComparison]::OrdinalIgnoreCase) -or
        !$Certificate.HasPrivateKey -or $Certificate.NotBefore -gt (Get-Date) -or $Certificate.NotAfter -le (Get-Date)) {
        throw 'The selected signing certificate must match the approved thumbprint, be current and have its private key.'
    }
    $codeSigning=$false
    foreach($extension in $Certificate.Extensions){if($extension -is [Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension]){foreach($usage in $extension.EnhancedKeyUsages){if($usage.Value -eq '1.3.6.1.5.5.7.3.3'){$codeSigning=$true}}}}
    if(!$codeSigning){throw 'The selected certificate must explicitly permit code signing.'}
    $chain=New-Object Security.Cryptography.X509Certificates.X509Chain
    try {
        $null=$chain.ChainPolicy.ApplicationPolicy.Add([Security.Cryptography.Oid]::new('1.3.6.1.5.5.7.3.3'))
        # Normal chain policy; do not import roots, disable revocation or accept
        # untrusted/self-signed chains just to make publishing succeed.
        if(!$chain.Build($Certificate)){throw 'Signing certificate trust/revocation validation failed under normal Windows policy.'}
    } finally {$chain.Dispose()}
}

function Get-PpiApprovedSigningCertificate([string]$Thumbprint) {
    if($Thumbprint -notmatch '\A[0-9A-Fa-f]{40}\z'){throw 'Certificate thumbprint must be exactly 40 hexadecimal characters.'}
    $store=New-Object Security.Cryptography.X509Certificates.X509Store('My','CurrentUser')
    try {
        $store.Open([Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
        $matches=$store.Certificates.Find([Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint,$Thumbprint,$false)
        if($matches.Count -ne 1){throw 'The approved thumbprint must identify exactly one CurrentUser/My certificate.'}
        $certificate=$matches[0]
        Assert-PpiSigningCertificate $certificate $Thumbprint
        return $certificate
    } finally {$store.Close()}
}

function Get-PpiReproducedArtifacts([Collections.IDictionary]$Files,[scriptblock]$Build) {
    # Verification outputs must not live beside attacker-writable cache entries.
    # Use this build user's temp area; never launch the verification copies.
    $directory=[IO.Path]::Combine([IO.Path]::GetTempPath(),'PowerPointIntelligence.SourceProof.'+[Guid]::NewGuid().ToString('N'))
    $null=New-Item -ItemType Directory -Force -Path $directory
    $rebuilt=[ordered]@{}
    foreach($key in $Files.Keys){$rebuilt[$key]=[IO.Path]::Combine($directory,[IO.Path]::GetFileName($Files[$key]))}
    # Staging files are verification-only: never loaded, installed or launched.
    $leases=New-Object Collections.Generic.List[object]
    try {
        & $Build $rebuilt
        foreach($key in $Files.Keys){$leases.Add((Get-PpiVerifiedArtifact $Files[$key] $rebuilt[$key]))}
        return ,$leases
    } catch {foreach($lease in $leases){$lease.Stream.Dispose()};throw}
    finally {
        # Delete only the generated, fixed-basename files, never a recursive tree.
        foreach($path in $rebuilt.Values){try{if([IO.File]::Exists($path)){[IO.File]::Delete($path)}}catch [IO.IOException]{Write-Warning 'A verification file is locked; it was retained without execution.'}catch [UnauthorizedAccessException]{Write-Warning 'Verification cleanup was denied; no security policy was changed.'}}
        try{[IO.Directory]::Delete($directory)}catch [IO.IOException]{Write-Warning 'Verification workspace retained because it contains unexpected files; no recursive cleanup attempted.'}
    }
}

function Assert-PpiSceneDependency([string]$Path) {
    $signature=Get-AuthenticodeSignature -LiteralPath $Path -ErrorAction Stop
    $identity=[Reflection.AssemblyName]::GetAssemblyName($Path)
    $token=[BitConverter]::ToString($identity.GetPublicKeyToken()).Replace('-','').ToLowerInvariant()
    if($signature.Status -ne 'Valid' -or !$signature.SignerCertificate -or
        $signature.SignerCertificate.Subject -notmatch '(^|,\s*)O=Microsoft Corporation(,|$)' -or
        $identity.Name -ne 'Newtonsoft.Json' -or $identity.Version -ne [Version]'13.0.0.0' -or $token -ne '30ad4fe6b2a6aeed') {
        throw 'The scene build dependency needs independent approved publisher/identity verification; a locally generated receipt is insufficient. The configured Microsoft-signed Newtonsoft.Json 13 reference was not verified.'
    }
}
