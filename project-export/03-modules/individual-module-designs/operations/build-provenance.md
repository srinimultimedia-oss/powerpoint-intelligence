# OPS-05 — Build provenance, recovery and dependency restoration

**Evidence category:** source verified; execution status is stated separately. Baseline includes the actual inspected checkout, not an assumption that every working byte equals HEAD.

## Design

| Dimension | Implementation-grounded description |
|---|---|
| 1. Purpose | Protect stable test runners against silent overwrite/recovery and prove loaded bytes match current trusted source/toolchain. |
| 2. Implementation status | Implemented local guard/review/reproduction paths; does not establish a comprehensive scanner verdict. |
| 3. Internal components/classes/interfaces | Build-Integrity functions; Scene-RunnerCache state/review; Piidx-TestSuite loader; Piidx-Recovery; Artifact-Inputs; Safe-Xml; Source-Fingerprint; locked JS restore. |
| 4. Important functions/methods | Assert-PpiMicrosoftBuildTool, Get-PpiVerifiedArtifact/ReproducedArtifacts, Assert-PpiSceneDependency, Get-PpiSceneCacheState, Start-PpiReviewedRunnerBuild/LegacySourceReview/PiidxSourceRecovery, Assert-PpiRunnerScannerClear. |
| 5. Inputs and outputs | Stable files/receipts, current source/reference/compiler hashes, reviewed scanner evidence. Outputs verified immutable byte leases/current cache receipt or explicit review audit; never a forged historical receipt. |
| 6. Dependencies | PowerShell/.NET IO/crypto/certificate identity, Roslyn deterministic builds, Newtonsoft approved identity and owned artifact/temp workspaces. |
| 7. Internal workflow | Refuse missing/unreceipted/hash-mismatched unsafe states; explicitly review old provenance/scanner records if allowed; compile current source; independently reproduce outputs outside cache; compare locked bytes; memory-load verified PIIDX snapshot or launch locked scene/template files; validate sources remain unchanged. |
| 8. External integrations | Windows signature/certificate providers and native compiler; npm ci --ignore-scripts for locked JS tooling. These are development operations, not runtime data service calls. |
| 9. Exceptions and error handling | Any provenance/source/detection/identity mismatch throws and preserves evidence. Verification cleanup is fixed-file/nonrecursive, retaining unexpected locked files with warning. |
| 10. Configuration | Stable artifacts/piidx, scene-contracts-<platform>, template-library-<platform>; PIIDX scanner evidence <=1 day old; normal/read-only complete fields; bounds1MBreceipt/16MBscanner/128MBverified code. |
| 11. Security | Read sharing denies mutation during verification/load; compiler role identity; detected old/new path/hash refused; no automatic quarantine restore. Recovery JSON constrained under active checkout artifacts. Normal shared cache reads retain existing verification contract. |
| 12. Tests | Test-BuildIntegrity (forged receipt/reproduced mismatch/immutable bytes/signatureinput); Test-SceneRunnerCache explicit/legacy missing recovery; Test-PiidxRecovery; Test-Checkmarx113558; Test-SourceFingerprint inert-script noexecution. |
| 13. Known limitations | Requires trusted checkout/toolchain/user. Scanner matching only sees supplied report; current collector lacks PIIDX scope. Default culture sort affects cross-host source fingerprint. Initial Test-S1 download path has a separate integrity gap. |
| 14. Technical debt | Collector scope mismatch and dependency restoration consistency need resolution; source/signature checks are not remote compromise attestation. |
| 15. Extension points | Expand scanner fixture coverage and deterministic recipe identities; standardize approved dependency policy; maintain stable paths and exact historical evidence rather than blanket exclusions. |

## 16. Complete owned source map

| Source | Analysis | Responsibility/declarations |
|---|---|---|
| `scripts/Build-Integrity.ps1:1` | Complete text read; no runtime execution | Declarations: Get-PpiContentHash, Assert-PpiMicrosoftBuildTool, Get-PpiVerifiedArtifact, Assert-PpiSigningCertificate, Get-PpiApprovedSigningCertificate |
| `scripts/Scene-RunnerCache.ps1:1` | Complete text read; no runtime execution | Declarations: Get-PpiSceneBuildKey, Get-PpiSceneCacheState, Start-PpiReviewedRunnerBuild, Read-PpiReviewedScannerEvidence, Assert-PpiRunnerScannerClear |
| `scripts/Piidx-Recovery.ps1:1` | Complete text read; no runtime execution | Declarations: Start-PpiPiidxSourceRecovery |
| `scripts/Artifact-Inputs.ps1:1` | Complete text read; no runtime execution | Declarations: Resolve-PpiArtifactInput, Read-PpiArtifactJsonText |
| `scripts/Source-Fingerprint.ps1:1` | Complete text read; no runtime execution | Declarations: Get-PpiFingerprintFileHash, Get-PpiSourceFingerprint, Get-PpiAddinFingerprint |
| `scripts/Safe-Xml.ps1:1` | Complete text read; no runtime execution | Declarations: Read-PpiSafeXml |
| `scripts/Piidx-TestSuite.ps1:1` | Complete text read; no runtime execution | Declarations: Import-PpiPiidxSuite, Assert-PpiPiidxSources |
| `scripts/Restore-LockedDependencies.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `scripts/Test-BuildIntegrity.ps1:1` | Complete text read; no runtime execution | Declarations: Check, Reject |
| `scripts/Test-Checkmarx113558.ps1:1` | Complete text read; no runtime execution | Declarations: Check, Reject |
| `scripts/Test-PiidxRecovery.ps1:1` | Complete text read; no runtime execution | Declarations: Evidence, Check, Reject |
| `scripts/Test-SourceFingerprint.ps1:1` | Complete text read; no runtime execution | Declarations: Check, Digest |
| `scripts/Test-SceneRunnerCache.ps1:1` | Complete text read; no runtime execution | Declarations: Check, Reject |

## 17. Handover and maintenance checks

Review the exact source/hash/line counts in `09-security/reviewed-files.csv` and this directory’s `source-module-mapping.csv`. Do not substitute historical PASS logs for current execution. Preserve unrelated module behavior and tests before modifying a boundary; new methods must have typed inputs, explicit error semantics and a corresponding negative test. See the dedicated AI/API/security/deployment documents for this module’s external contract.
