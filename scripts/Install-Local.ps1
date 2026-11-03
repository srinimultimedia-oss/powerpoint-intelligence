param([ValidateSet('x86','x64')][string]$Platform = 'x64', [string]$CertificateThumbprint, [switch]$SkipBuild)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'Source-Fingerprint.ps1')
$sourceFingerprint=Get-PpiAddinFingerprint $repo
$log = Join-Path $repo ('artifacts/install-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.log')
New-Item -ItemType Directory -Force (Join-Path $repo 'artifacts') | Out-Null
Start-Transcript -Path $log | Out-Null
Push-Location $repo
try {
    if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
        throw 'Save your work and close PowerPoint normally before installation. No user process was closed.'
    }
    $manifest = Join-Path $repo "apps/powerpoint-vsto/bin/$Platform/Release/PowerPointIntelligence.vsto"
    if (!$SkipBuild) {
    Write-Host 'STEP signing certificate and build'
    # A rebuilt ClickOnce payload needs a newer deployment version so the
    # installed add-in receives the changed binaries rather than a cached build.
    $deploymentVersion = [Version]'1.0.0.0'
    if (Test-Path $manifest) {
        $previousManifest = [xml](Get-Content -LiteralPath $manifest -Raw)
        $identity = $previousManifest.SelectSingleNode("/*[local-name()='assembly']/*[local-name()='assemblyIdentity']")
        if ($identity) {
            $previousVersion = [Version]$identity.GetAttribute('version')
            if ($previousVersion.Revision -ge 65535) { throw 'Deployment revision exhausted; update the project version explicitly.' }
            $deploymentVersion = [Version]::new($previousVersion.Major,$previousVersion.Minor,$previousVersion.Build,$previousVersion.Revision + 1)
        }
    }
    Write-Host "Building deployment version $deploymentVersion"
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
    $vsPath = & $vswhere -latest -products '*' -property installationPath
    if (!$vsPath) { throw 'Visual Studio with Office development tools is required.' }
    if (!$CertificateThumbprint) {
        $cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert |
            Where-Object { $_.Subject -eq 'CN=PowerPoint Intelligence Local Development' -and $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date).AddDays(1) } |
            Sort-Object NotAfter -Descending | Select-Object -First 1
        if (!$cert) {
            $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject 'CN=PowerPoint Intelligence Local Development' `
                -CertStoreLocation Cert:\CurrentUser\My -KeyExportPolicy NonExportable -HashAlgorithm SHA256 -NotAfter (Get-Date).AddDays(90)
        }
        $CertificateThumbprint = $cert.Thumbprint
    }
    $certPath = 'Cert:\CurrentUser\My\' + $CertificateThumbprint
    $cert = Get-Item -LiteralPath $certPath
    if (!$cert.HasPrivateKey -or $cert.NotAfter -le (Get-Date)) { throw 'Signing certificate must have a private key and be unexpired.' }
    # Snapshot rollback can preserve older source timestamps. Force compilation;
    # never reuse a later payload just because incremental inputs look older.
    & (Join-Path $vsPath 'MSBuild/Current/Bin/MSBuild.exe') apps/powerpoint-vsto/PowerPointIntelligence.csproj `
        /t:Build /p:PpiForcePayloadCompile=true /p:Configuration=Release "/p:Platform=$Platform" /p:SignManifests=true /p:UseSharedCompilation=false `
        "/p:ManifestCertificateThumbprint=$CertificateThumbprint" "/p:ApplicationVersion=$deploymentVersion" /v:minimal /nologo
    if ($LASTEXITCODE) { throw 'Signed VSTO build failed; see install log.' }
    if($sourceFingerprint -ne (Get-PpiAddinFingerprint $repo)) { throw 'Add-in source changed during build. Installation stopped; rerun after source edits finish.' }
    }
    if (!(Test-Path $manifest)) { throw 'Build did not produce the expected VSTO deployment manifest.' }
    & (Join-Path $PSScriptRoot 'Test-DemoPackage.ps1') -Platform $Platform
    $deploymentXml = [xml](Get-Content -LiteralPath $manifest -Raw)
    if ($deploymentXml.SelectNodes("//*[local-name()='Signature']").Count -eq 0) { throw 'Deployment manifest is unsigned.' }
    $installer = Join-Path $env:CommonProgramFiles 'microsoft shared/VSTO/10.0/VSTOInstaller.exe'
    if (!(Test-Path $installer)) { throw 'VSTO runtime installer is missing.' }
    # Interactive installer presents its normal publisher/trust prompt. No root
    # certificate, Trusted Publisher entry or Windows security policy is changed here.
    Write-Host 'STEP VSTO installation: complete the visible installer prompt (check Alt+Tab if needed).'
    $installation = Start-Process -FilePath $installer -ArgumentList @('/Install', ('"' + $manifest + '"')) -PassThru -WindowStyle Normal
    Write-Host "Waiting for VSTO installer process $($installation.Id)."
    $installation.WaitForExit()
    $installation.Refresh()
    if ($installation.ExitCode -ne 0) { throw "VSTO installer failed: $($installation.ExitCode)." }
    [pscustomobject]@{Manifest=$manifest;SourceSha256=$sourceFingerprint;AssemblySha256=(Get-FileHash (Join-Path (Split-Path $manifest) 'PowerPointIntelligence.dll') -Algorithm SHA256).Hash;Version=$deploymentXml.SelectSingleNode("/*[local-name()='assembly']/*[local-name()='assemblyIdentity']").GetAttribute('version');InstallerExitCode=$installation.ExitCode;Time=(Get-Date).ToString('o')} |
        ConvertTo-Json | Set-Content (Join-Path $repo 'artifacts/installed-demo-receipt.json')
    Write-Host 'VSTO installer succeeded. Open PowerPoint and check the NexWave ribbon tab.'
    Write-Host 'Runtime loading and ribbon appearance still require verification in PowerPoint.'
} finally {
    Pop-Location
    Stop-Transcript | Out-Null
    Write-Host "Installation log: $log"
}
