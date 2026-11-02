$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
Push-Location $repo
try {
    foreach($package in @('chart-domain','chart-engine')) {
        & npm.cmd ci --ignore-scripts --no-audit --no-fund --prefix (Join-Path 'packages' $package)
        if($LASTEXITCODE){throw "Locked dependency restoration failed: $package"}
    }
    'PASS locked JavaScript development dependencies restored; package lifecycle scripts were disabled. Office installation and held runner review are separate.'
} finally {Pop-Location}