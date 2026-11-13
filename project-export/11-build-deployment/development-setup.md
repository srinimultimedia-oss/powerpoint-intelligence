# Development setup

## Verified toolchain

Windows desktop PowerPoint, .NET Framework 4.8 targeting pack, C# 7.3, Visual Studio Office/VSTO tools, VSTO Runtime, Office/PowerPoint PIAs and Windows PowerShell are the native stack. The project is a legacy MSBuild VSTO library, not Office.js/.NET8 or a standalone native application. `apps/powerpoint-vsto/PowerPointIntelligence.csproj:9`.

Office version15 PIAs use embedded interop types while runtime guard requires Office 16.0 API. x86/x64 refer to PowerPoint process, not OS bitness. A configured profile is not product certification. `apps/powerpoint-vsto/OfficeCompatibility.cs:69`.

Node/npm/TypeScript/Ajv support deterministic chart/domain tests. CI uses Node 20 for chart-core and Node 24 for S1, not one central pin. Python/lxml support offline repair/security tests/maintenance. These are development tools, not bootstrapper prerequisites.

## Preparation

1. Install approved Visual Studio Office workload/.NET4.8 targeting pack, desktop Office and VSTO Runtime via normal estate tooling.
2. Restore JavaScript with scripts/Restore-LockedDependencies.ps1 (npm ci --ignore-scripts for domain/engine).
3. Review fresh-cache Newtonsoft provenance before Test-S1: its missing-DLL branch downloads pinned NuGet without a pinned hash (SEC-01).
4. Use normal Windows PowerShell for signing/packaging and an approved certificate in Windows storage. Never export a key into the checkout.
5. Save/close presentations before install/interactive tests.

No installation/restoration/build was executed for this documentation export; existing scripts write outside project-export. [Build instructions](build-instructions.md) are reproducible operator procedures; [test results](../12-testing/test-results.md) distinguish actual current safe checks.
