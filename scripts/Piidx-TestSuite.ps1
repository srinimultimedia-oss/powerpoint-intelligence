function Import-PpiPiidxSuite {
    param([string]$Root,[string]$RunnerRoot)
    $ErrorActionPreference='Stop'
    $files=@(Get-ChildItem (Join-Path $Root 'src/Library.Core') -File -Filter '*.cs';Get-ChildItem (Join-Path $Root 'tests/library/piidx') -File -Filter '*.cs') | Sort-Object FullName
    $sources=@($files | ForEach-Object {[pscustomobject]@{File=$_.FullName.Substring($Root.Length+1);SHA256=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash}})
    $inputText=($sources | ForEach-Object {$_.File+'='+$_.SHA256}) -join "`n"
    $hasher=[Security.Cryptography.SHA256]::Create();try{$inputHash=[BitConverter]::ToString($hasher.ComputeHash([Text.Encoding]::UTF8.GetBytes($inputText))).Replace('-','')}finally{$hasher.Dispose()}
    if(!$RunnerRoot){$RunnerRoot=$Root}
    $folder=Join-Path $RunnerRoot 'artifacts/piidx';$null=New-Item -ItemType Directory -Path $folder -Force
    $dll=Join-Path $folder 'PiidxTestSuite.dll';$receipt=Join-Path $folder 'test-suite-build.json'
    $prior=$null;if(Test-Path -LiteralPath $receipt){$prior=Get-Content -LiteralPath $receipt -Raw | ConvertFrom-Json}
    if($prior -and !(Test-Path -LiteralPath $dll)){throw 'PIIDX test assembly is missing despite its receipt. Inspect security/quarantine evidence before rebuilding; no automatic restoration is attempted.'}
    if($prior -and (Get-FileHash -LiteralPath $dll -Algorithm SHA256).Hash -ne $prior.AssemblySHA256){throw 'PIIDX test assembly differs from its build receipt; preserve evidence and review before rebuilding.'}
    if('PiidxSchemaTests' -as [type]) {
        if(!$prior -or $prior.InputSHA256 -ne $inputHash){throw 'PIIDX sources changed after this PowerShell loaded the suite. Start a fresh PowerShell session; loaded assemblies are not replaced.'}
    } else {
        if(!$prior -or $prior.InputSHA256 -ne $inputHash) {
            if(!$prior -and (Test-Path -LiteralPath $dll)){throw 'An unreceipted PIIDX assembly exists. Preserve/review it; this script will not overwrite it.'}
            $vswhere=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
            $vs=& $vswhere -latest -products '*' -property installationPath
            $compiler=Join-Path $vs 'MSBuild/Current/Bin/Roslyn/csc.exe'
            $compilerOutput = & $compiler /nologo /target:library /deterministic+ /langversion:7.3 /codepage:65001 /platform:AnyCPU "/out:$dll" /reference:System.Xml.Linq.dll /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll @($files.FullName)
            if($LASTEXITCODE){throw ('PIIDX .NET Framework test-suite compilation failed. '+($compilerOutput -join [Environment]::NewLine))}
            [pscustomobject]@{InputSHA256=$inputHash;AssemblySHA256=(Get-FileHash -LiteralPath $dll -Algorithm SHA256).Hash;State='Compilation only; not execution evidence';Sources=$sources} | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $receipt
        }
        $null=[Reflection.Assembly]::LoadFrom($dll)
    }
    [pscustomobject]@{Assembly=$dll;AssemblySHA256=(Get-FileHash -LiteralPath $dll -Algorithm SHA256).Hash;InputSHA256=$inputHash;Sources=$sources}
}
function Assert-PpiPiidxSources {
    param([string]$Root,[object[]]$Sources)
    foreach($source in $Sources){if((Get-FileHash -LiteralPath (Join-Path $Root $source.File) -Algorithm SHA256).Hash -ne $source.SHA256){throw 'PIIDX source changed while tests ran. Results must not be accepted.'}}
}