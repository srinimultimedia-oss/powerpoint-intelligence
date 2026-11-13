# OPS-07 — Validation orchestration and CI

**Evidence category:** source verified; execution status is stated separately. Baseline includes the actual inspected checkout, not an assumption that every working byte equals HEAD.

## Design

| Dimension | Implementation-grounded description |
|---|---|
| 1. Purpose | Compile/test deterministic contracts, stable caches and optional desktop Office scenarios, retaining exact failure/log boundaries. |
| 2. Implementation status | Test runners/CI recipes implemented; this leaf performs source analysis only. Current safe executions belong to parent test-results. |
| 3. Internal components/classes/interfaces | Test-All, Check-Process, per-suite wrappers, focused DemoOfficeTests switches, PIIDX wrappers, four GitHub workflows and root workflow duplicate. |
| 4. Important functions/methods | Invoke-PpiCheckProcess, ConvertTo-PpiNativeArgument, Test-All Invoke-Check; per-suite param entrypoints and PIIDX Assert-PpiPiidxSources. |
| 5. Inputs and outputs | Current source/fixtures/dependencies/platform/interactive flags; outputs source snapshot fingerprint, results.json/md/logs, synthetic decks/scenes/catalogs and process metadata. |
| 6. Dependencies | PowerShell/Roslyn/MSBuild/PIAs/.NET4.8/node/npm/Python/lxml; stable runner source proof and optional desktop COM/publisher context. |
| 7. Internal workflow | Snapshot source with before/after/captured fingerprint equality; run local suites; threearchitecture policy/contracts/compile; optional installed/template/native/demo/Office iterations; stop after timeout/activation denial; collect owned crash metadata; emit results even on failure. |
| 8. External integrations | Owned child processes and optionally Office; CI uses windows-latest for metadata/domain/compatibility and ubuntu-latest for chart-core, Node 20/24 by workflow. |
| 9. Exceptions and error handling | File-backed stdout/stderr avoid pipe deadlock; bounded timeout kills only owned process and stops remaining suite; no exit code inference. Native steps stop at first failed assertion; later steps are not passes. |
| 10. Configuration | Local300s default, Office3600s, OfficeRuns3; compile-only/metadata-only/focusedmodes. Test-All does not automatically execute every standalone wrapper/regression. |
| 11. Security | Snapshots omit signing/env secret files; no user Office process is killed; runner reuse checks provenance. Test sources can create artifacts/temp, so not invoked unchanged under export-only rule. |
| 12. Tests | Test-CheckProcess exercises largestreams/timeout/quoting; Test-Ribbon invokes cache/process/marker contracts; specialized metadata/domain/scene/native/data/agenda/property tests; CI only selected subset. |
| 13. Known limitations | No measured codecoverage/branchcoverage; offline compile does not prove ribbon or native editability. Root chart-core.yml outside .github is not a separate active GitHub workflow. |
| 14. Technical debt | Partial CI matrix, inconsistent Node/compiler signature verification across older wrappers, standalone tests outside aggregate; native environment remains external gate. |
| 15. Extension points | Add consistent test manifest/coverage and approved desktop matrix without turning runtime prerequisites into skipped passes; keep immutable source/evidence and safe timeout semantics. |

## 16. Complete owned source map

| Source | Analysis | Responsibility/declarations |
|---|---|---|
| `scripts/Check-Process.ps1:1` | Complete text read; no runtime execution | Declarations: ConvertTo-PpiNativeArgument, Invoke-PpiCheckProcess |
| `scripts/New-PiidxSample.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `scripts/Test-All.ps1:1` | Complete text read; no runtime execution | Declarations: Invoke-Check |
| `scripts/Test-CheckProcess.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `scripts/Test-DemoFeatures.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `scripts/Test-InstalledRibbon.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `scripts/Test-NativeRenderer.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `scripts/Test-OfficeCompatibility.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `scripts/Test-Piidx.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `scripts/Test-PiidxCatalog.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `scripts/Test-PiidxCorpus.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `scripts/Test-PiidxFoundation.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `scripts/Test-PiidxIsolation.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `scripts/Test-PiidxPreviews.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `scripts/Test-PiidxSourceReader.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `scripts/Test-PiidxSqliteCapability.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `scripts/Test-Ribbon.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `scripts/Test-S0.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `scripts/Test-S1.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `scripts/Test-SceneContracts.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `scripts/Test-SecurityRemediation.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `scripts/Test-TemplateLibrary.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `.github/workflows/chart-core.yml:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `.github/workflows/office-compatibility.yml:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `.github/workflows/s0-metadata.yml:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `.github/workflows/s1-domain.yml:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `chart-core.yml:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |

## 17. Handover and maintenance checks

Review the exact source/hash/line counts in `09-security/reviewed-files.csv` and this directory’s `source-module-mapping.csv`. Do not substitute historical PASS logs for current execution. Preserve unrelated module behavior and tests before modifying a boundary; new methods must have typed inputs, explicit error semantics and a corresponding negative test. See the dedicated AI/API/security/deployment documents for this module’s external contract.
