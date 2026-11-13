# Build instructions

Commands reverse-document source; they are not execution claims for this export. Run from repository root in an approved desktop environment.

## Compile only

```powershell
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
$vs = & $vswhere -latest -products '*' -property installationPath
& (Join-Path $vs 'MSBuild/Current/Bin/MSBuild.exe') apps/powerpoint-vsto/PowerPointIntelligence.csproj /t:Compile /p:Configuration=Release /p:Platform=x64 /v:minimal /nologo
```

Allowed platforms AnyCPU,x86,x64; Debug emits symbols, Release optimizes. Compile alone does not create signed deployment/validate loading. Selected domain/renderer/Library.Core files compile directly into add-in; the separate domain project is not the whole production graph. `apps/powerpoint-vsto/PowerPointIntelligence.csproj:39`.

## Signed self-contained deployment

```powershell
.\easySetup\Build-EasySetup.ps1 -CertificateThumbprint '<approved-40-hex-thumbprint>'
.\easySetup\Test-EasySetup.ps1
```

Builder validates Microsoft tool identity and approved current certificate/private key/EKU/chain; increments four-part version; builds signed x86/x64 payloads; validates payload marker/hash/template inclusion; packages and signs pending EXE; requires Valid Authenticode plus unchanged source; then atomically replaces public EXE. Failures preserve the previous publication. `easySetup/Build-EasySetup.ps1:19` `easySetup/Build-EasySetup.ps1:84`.

PpiForcePayloadCompile deletes only known intermediate assembly before CoreCompile to avoid restored timestamp reuse. PpiEasySetupBuild prevents recursive package builds; PpiSkipEasySetup suppresses the signed Release hook deliberately. Signed Release Build otherwise calls builder automatically. `apps/powerpoint-vsto/PowerPointIntelligence.csproj:73`.

## Watcher

Watch-EasySetup polls default 5 seconds (allowed2–300), waits for stable fingerprint, rebuilds once for a changed source and keeps watching. Output after publication is normal idle polling. Ctrl+C stops. A failed source version keeps the old EXE and waits for another edit, rather than endlessly retrying. `easySetup/Watch-EasySetup.ps1:1`.

## Check side effects

Test-All creates artifacts/source snapshots and .NET/TS outputs; -InteractiveOffice adds real Office tests. Test-EasySetup -SourceOnly still compiles/runs inert fixtures under artifacts. They were not invoked unchanged in this documentation-only task; their source is analyzed and current safe validation is documented separately.
