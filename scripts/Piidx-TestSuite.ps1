. (Join-Path $PSScriptRoot 'Build-Integrity.ps1')
function Import-PpiPiidxSuite {
    param([string]$Root,[string]$RunnerRoot,[switch]$RecoverMissingRunner,[string]$ScannerEvidence)
    $ErrorActionPreference='Stop'
    $sourceFiles=@(Get-ChildItem (Join-Path $Root 'src/Library.Core') -File -Filter '*.cs';Get-ChildItem (Join-Path $Root 'tests/library/piidx') -File -Filter '*.cs') | Sort-Object FullName
    $sources=@($sourceFiles | ForEach-Object {[pscustomobject]@{File=$_.FullName.Substring($Root.Length+1);SHA256=(Get-PpiContentHash $_.FullName)}})
    if(!$RunnerRoot){$RunnerRoot=$Root}
    $folder=Join-Path $RunnerRoot 'artifacts/piidx';$null=New-Item -ItemType Directory -Path $folder -Force
    $dll=Join-Path $folder 'PiidxTestSuite.dll';$receipt=Join-Path $folder 'test-suite-build.json'
    $prior=$null;if(Test-Path -LiteralPath $receipt){$prior=Get-Content -LiteralPath $receipt -Raw | ConvertFrom-Json}
    $recovery=$null
    if($prior -and !(Test-Path -LiteralPath $dll)){
        if(!$RecoverMissingRunner){throw 'PIIDX test assembly is missing despite its receipt. Inspect security/quarantine evidence before rebuilding; no automatic restoration is attempted.'}
        . (Join-Path $PSScriptRoot 'Piidx-Recovery.ps1')
        $recovery=Start-PpiPiidxSourceRecovery $receipt $dll $ScannerEvidence $sources
    } elseif($RecoverMissingRunner){throw 'Explicit PIIDX recovery requires an existing receipt and a missing DLL.'}
    if($prior -and !$recovery -and (Get-PpiContentHash $dll) -ne $prior.AssemblySHA256){throw 'PIIDX test assembly differs from its build receipt; preserve evidence and review before rebuilding.'}
    $vswhere=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
    Assert-PpiMicrosoftBuildTool $vswhere 'vswhere.exe'
    $vs=& $vswhere -latest -products '*' -property installationPath
    $compiler=Join-Path $vs 'MSBuild/Current/Bin/Roslyn/csc.exe'
    Assert-PpiMicrosoftBuildTool $compiler 'csc.exe'
    $inputText='piidx-reproduced-v1|'+(Get-PpiContentHash $compiler)+'|'+(($sources | ForEach-Object {$_.File+'='+$_.SHA256}) -join "`n")
    $hasher=[Security.Cryptography.SHA256]::Create();try{$inputHash=[BitConverter]::ToString($hasher.ComputeHash([Text.Encoding]::UTF8.GetBytes($inputText))).Replace('-','')}finally{$hasher.Dispose()}
    $build={param($targets)
        $compilerOutput=& $compiler /nologo /target:library /deterministic+ /langversion:7.3 /codepage:65001 /platform:AnyCPU "/pathmap:$Root=/_/PowerPointIntelligence" "/out:$($targets.Assembly)" /reference:System.Xml.Linq.dll /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll @($sourceFiles.FullName)
        if($LASTEXITCODE){throw ('PIIDX source compilation failed. '+($compilerOutput -join [Environment]::NewLine))}
    }.GetNewClosure()
    $loaded='PiidxSchemaTests' -as [type]
    if($loaded -and (!$prior -or $prior.InputSHA256 -ne $inputHash -or $script:PpiVerifiedPiidxInput -ne $inputHash)){throw 'PIIDX loaded assembly has no matching current-source verification. Start a fresh PowerShell process.'}
    if(!$loaded -and ($recovery -or !$prior -or $prior.InputSHA256 -ne $inputHash)) {
        if(!$prior -and (Test-Path -LiteralPath $dll)){throw 'An unreceipted PIIDX assembly exists. Preserve/review it; this script will not overwrite it.'}
        & $build @{Assembly=$dll}
        if($recovery){Assert-PpiRunnerScannerClear $recovery.Report ([ordered]@{Assembly=$dll}) ([ordered]@{Assembly=(Get-PpiContentHash $dll)});Assert-PpiPiidxSources $Root $sources}
        [pscustomobject]@{InputSHA256=$inputHash;AssemblySHA256=(Get-PpiContentHash $dll);State='Compilation only; not execution evidence';Sources=$sources} | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $receipt
    }
    $leases=Get-PpiReproducedArtifacts ([ordered]@{Assembly=$dll}) $build
    try {
        Assert-PpiPiidxSources $Root $sources
        if(!$loaded){$null=[Reflection.Assembly]::Load([byte[]]$leases[0].Bytes);$script:PpiVerifiedPiidxInput=$inputHash}
        [pscustomobject]@{Assembly=$dll;AssemblySHA256=$leases[0].SHA256;InputSHA256=$inputHash;Sources=$sources}
    } finally {foreach($lease in $leases){$lease.Stream.Dispose()}}
}
function Assert-PpiPiidxSources {
    param([string]$Root,[object[]]$Sources)
    foreach($source in $Sources){if((Get-FileHash -LiteralPath (Join-Path $Root $source.File) -Algorithm SHA256).Hash -ne $source.SHA256){throw 'PIIDX source changed while tests ran. Results must not be accepted.'}}
}
