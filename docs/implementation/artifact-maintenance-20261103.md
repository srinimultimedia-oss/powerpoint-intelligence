# NexWave artifact maintenance and PIIDX recovery — 3 November 2026

## Scope

Cleanup is restricted to `C:\Srini\powerpoint-intelligence\powerpoint-intelligence\artifacts`. Source/runtime folders outside artifacts, the published easySetup EXE, installed add-in path and per-user PowerPoint data were not cleanup targets. The installed receipt points to `apps/powerpoint-vsto/bin/x64/Release/PowerPointIntelligence.vsto`, outside cleanup scope.

Initial inventory: 4,559,539,762 bytes / 67,581 files (before maintenance logs and recovered outputs). Large consumers were historical validation source trees and repeated installer staging payloads.

Policy in `scripts/Maintain-Artifacts.py`: retain at least the two newest validation/installer runs, stable runner/cache receipts, dependencies, scanner evidence, merge backups and native acceptance folders outside selected groups. Signing/key/environment files, previous-setup rollback packages and historical binary source inputs are protected. Select only older validation `source` subtrees, older easySetup stages and old synthetic PIIDX/gallery/test fixture directories. Top-level validation results/logs remain. No tracked artifact files were found in the repository inventory.

Before removal, all candidate paths and hashes are audited. Historical source and readable/native evidence are stored in a content-addressed ZIP, deduplicated by SHA-256; all archived objects are read back and hashed before deletion starts. Generated executables/dependencies/databases in disposable subtrees are discarded, never archived/restored as runnable backups. Any source input binary is retained in place. The manifest maps original artifact-relative paths to archived SHA-256 objects. Changed/locked files are retained. No reparse paths are followed or removed; all cleanup targets are checked inside artifacts. Directory cleanup removes only empty directories, nonrecursively.

Run without `--apply` to inspect a plan; `--apply` performs the audited archive/delete sequence. A fresh `--audit-name maintenance-...` is required for each run. The source archive contains local licensed assets/history and stays private in artifacts; do not distribute it as a product package. To inspect archived evidence, read `manifest.json`, locate its `objects/<sha256>` entry and verify its hash before copying that data/source to a chosen safe location. No binary/cache recovery is implied by the archive.

Policy verification: `tests/security/test_artifact_maintenance.py` passes isolated fixture checks for boundaries, archive hashes, retained evidence/latest runs/dependencies/runner paths/keys/binary inputs, and untouched files outside artifacts. Evidence: `artifacts/maintenance-20261103/cleanup-policy-tests.log`.

## PIIDX repair

The cause of the original missing DLL is unconfirmed. Fresh 14-day scanner evidence was collected at `artifacts/maintenance-20261103/scanner/scene-threat-evidence.json`: 14 historical McAfee records, zero read errors, no PIIDX path/hash match. This is not a false-positive verdict or proof of active antivirus protection.

The owner explicitly requested recovery. `scripts/Piidx-Recovery.ps1` adds an opt-in source-recovery path: require an absent output and a bounded historical receipt with valid source/hash provenance, current source enumeration, complete scanner evidence collected within one day, and no matching historic hash/path detection. Unsupported/escaping source names and existing files fail closed. It preserves the exact prior receipt and its SHA-256 in a new audit; no old DLL is adopted/restored.

`Import-PpiPiidxSuite` still refuses ordinary missing-cache runs. With explicit recovery, it authenticates the compiler, rebuilds current source at the existing `artifacts/piidx/PiidxTestSuite.dll`, checks the new hash against scanner records and source stability, independently reproduces it, compares bytes and loads the verified snapshot. Default missing/hash/unreceipted controls remain enabled. No trust, execution-policy or antivirus setting was changed.

Historical DLL hash recorded in the prior receipt: `4272474C852AD1D77FB5844774473318F734741DB772F72BC8DF903E4723E445`.
Recovered current-source DLL hash: `62D512358EE54F2782857D1A7D2FA803C0FEE4AA46A477496E13D576FADDB086`.
Audit: `artifacts/piidx/test-suite-build.json.source-recovery-0cb469faa98c409bb76011efb24a2d19.json`.

Recovery command used (only while the DLL was absent):

```powershell
.\scripts\Test-Piidx.ps1 -RecoverMissingRunner -ScannerEvidence artifacts/maintenance-20261103/scanner/scene-threat-evidence.json
```

Subsequent tests use ordinary `.\scripts\Test-Piidx.ps1`, with the same stable DLL path. Do not repeat the recovery switch for an existing DLL. Never delete a receipt or disable scanner checks to force recovery.

All seven PIIDX suites passed: foundation, sources, catalog, previews, edgecases, workspace and corpus. Corpus: 62 packages / 497 slides, 13,049 objects, 2,078 tags. No Office automation was executed. Evidence: `artifacts/maintenance-20261103/piidx-recovery.log` and `artifacts/piidx/suite-2f00523a36fd4b9a819c8a9c42e6d499/results.json`. Ten inert recovery-policy assertions also pass (matched detections, stale/incomplete evidence, bad paths and existing-output refusal). Date handling accepts typed PS7 dates and invariant ISO strings without locale-dependent reinterpretation.

Final cleanup totals and full-suite evidence are appended after execution below.

## Final local validation

Full noninteractive Test-All completed successfully: 24/24 PASS, including security/recovery policy, PIIDX engine/isolation, all three scene-contract architectures and all three VSTO architecture compilations. Evidence: artifacts/validation-20261103-100523-60468512/results.json; source snapshot DE7A664B3777F378C5B86AF805C9CB69A9CA54A5C3A4FDB6DC034E37BFEE66EA. The ordinary PIIDX suite (without the recovery switch) passed again in this source snapshot. Interactive PowerPoint activation/loading/native save-reopen and installer signing were not run for this maintenance task; these results do not establish their acceptance. No application feature source was changed by this maintenance task.



## Completed cleanup

57,561 files removed / 3,454,576,029 bytes. Verified source/evidence archive: 49,926,801 bytes / 1,961 unique SHA-256 objects; SHA-256 1465D2642E210981DAD1FC9AF5665C39466CD542A6DC321C5FF69E1EFD0DBE23. Net reclaimed after archive and plan: 3,385,916,712 bytes (3.39 decimal GB). No skipped files. Audit: artifacts/maintenance-20261103-clean/cleanup-result.json and cleanup-plan.json. Protected runner/dependency/receipt/scanner paths were checked present afterward. Post-cleanup inventory is recorded separately because the full validation created a fresh source snapshot concurrently. PIIDX DLL hash stayed 62D512358EE54F2782857D1A7D2FA803C0FEE4AA46A477496E13D576FADDB086 on ordinary reuse.

