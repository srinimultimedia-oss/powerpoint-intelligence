# OPS-09 — Artifact maintenance and repository dependency audit

**Evidence category:** source verified; execution status is stated separately. Baseline includes the actual inspected checkout, not an assumption that every working byte equals HEAD.

## Design

| Dimension | Implementation-grounded description |
|---|---|
| 1. Purpose | Inventory local files/dependency references and conservatively reclaim obsolete generated artifacts with recoverable source/evidence history. |
| 2. Implementation status | Implemented offline utilities; documentation task does not execute their cleanup/audit writes outside export. |
| 3. Internal components/classes/interfaces | Maintain-Artifacts digest/linked/confined/files_under/main; Clean-Workspace preview/apply; audit-library-dependencies readscan/export; maintenance policy fixtures. |
| 4. Important functions/methods | Maintain --apply/--keep/--audit-name; Clean-Workspace -Apply/-KeepLatestSnapshots; audit --root/--git. |
| 5. Inputs and outputs | Artifact groups, installed manifest receipt, repository/Git metadata; outputs plan/archive/hash/result or classified inventory/dependency-hit tables. |
| 6. Dependencies | Python/PowerShell filesystem/ZIP/hash, GitCLI for dependency audit. No Office or network. |
| 7. Internal workflow | Default maintenance plan selects old groups but keeps two latest; protects signing/rollback/source binary inputs and active installed manifest; content-address source/evidence; read/hash every archive object; recheck file hash before unlink; remove only emptydirs. Audit records everyfile/hash/classification and text dependency hits. |
| 8. External integrations | Local filesystem and Git metadata; existing audit-library script writes docs/migration, so it is not safe to invoke unchanged for export-only task. |
| 9. Exceptions and error handling | Confined/reparse checks reject outside root; changed/locked files retained or reported; audit source errors/missing tracked records fail. Existing audit plan refuses overwrite. |
| 10. Configuration | Maintain default keep 2 and bounded maintenance-* name; roots derive from script repository. Older Clean-Workspace may also target apps/powerpoint-vsto/obj, unlike artifact-only Maintain. |
| 11. Security | No old executables restored from archive; generated binarycache data discarded with audit hashes. Existing proprietary/local asset history may be sensitive/licensed; not portable product source. |
| 12. Tests | test_artifact_maintenance verifies archive hashes, two latest snapshots, stable runners/deps/sourcebinary/key preservation and outside/root boundary refusal. |
| 13. Known limitations | Historical archive recovery is data/source inspection, not executable restoration. Read-only audit label includes outputs; metadata/package signature inventory is not semantic source/runtime review. |
| 14. Technical debt | Two maintenance tools have distinct scopes/retention rules; normalize operational naming/docs before unattended use. Existing audit recursively includes ignored/admin files, requiring sensitive-data discipline. |
| 15. Extension points | Explicit dry-run confirmation/approved automation retention and source policy tests; preserve stable caches and verified evidence before deleting expanded history. |

## 16. Complete owned source map

| Source | Analysis | Responsibility/declarations |
|---|---|---|
| `scripts/Clean-Workspace.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `scripts/Maintain-Artifacts.py:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `scripts/audit-library-dependencies.py:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `tests/security/test_artifact_maintenance.py:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |

## 17. Handover and maintenance checks

Review the exact source/hash/line counts in `09-security/reviewed-files.csv` and this directory’s `source-module-mapping.csv`. Do not substitute historical PASS logs for current execution. Preserve unrelated module behavior and tests before modifying a boundary; new methods must have typed inputs, explicit error semantics and a corresponding negative test. See the dedicated AI/API/security/deployment documents for this module’s external contract.
