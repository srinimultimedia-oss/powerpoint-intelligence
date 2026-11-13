# Deployment guide and acceptance

## Intent versus certification

Runtime guard targets common Office 16.0 API. Profile auto/office2016/office2019/microsoft365 records intent; edition is not auto-detected. Architecture any/x86/x64 checks actual PowerPoint process; configured full build minimum fails closed if build unavailable. Full Office channel/build/DPI/locale certification is UNKNOWN. `apps/powerpoint-vsto/OfficeCompatibility.cs:85` `docs/DECISIONS.md:29`.

## Release sequence

1. Review baseline, dependency provenance, security findings and target estate decisions.
2. Run Test-All in approved desktop checkout; retain exact source/test evidence.
3. Build signed easySetup with approved certificate; Test-EasySetup in normal Windows PowerShell; retain build.json/README.
4. Save/close presentations; run EXE, complete normal publisher prompt, retain extracted versions.
5. Open PowerPoint; confirm NexWave/build identity and run native/link/library/edit/save/reopen tests plus visual/DPI checks.
6. Distribute only allowed licensed templates and approved publisher chain.

On-prem API/Azure gateway are not implemented; no server topology/ports/secret environment are invented. Office/VSTO/.NET prerequisites, target publisher trust, identity/models/retention policies remain enterprise dependencies. `services/ai-gateway/README.md:1`.

## Historical evidence

The current build receipt records a prior publication. docs/implementation/deployment-validation-20261103.md records earlier24/24 local checks, version1.0.0.18 desktop signing and activation limitations. These are historical records, not checks rerun now. Current export verification is in [test results](../12-testing/test-results.md).

## Unsupported assumptions

No macOS/web/mobile add-in, unattended server Office automation, automatic downloaded prerequisites, remote admin/updater service, license enforcement or HA/DR implementation exists. Source edits update EXE only when signed Release hook/watcher runs successfully; a source save alone is not publication.
