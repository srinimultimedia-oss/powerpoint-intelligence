# OPS-03 — Signed packaging build and update watcher

**Evidence category:** source verified; execution status is stated separately. Baseline includes the actual inspected checkout, not an assumption that every working byte equals HEAD.

## Design

| Dimension | Implementation-grounded description |
|---|---|
| 1. Purpose | Build/verify/publish one self-contained x86+x64 VSTO installer while retaining previous successful publication on failure. |
| 2. Implementation status | Implemented source/build recipe; no new signing/installation executed during export. |
| 3. Internal components/classes/interfaces | Build-EasySetup.ps1, Common.ps1, Watch-EasySetup.ps1, VSTO MSBuild targets, Install-Local and package-marker/hash validators. |
| 4. Important functions/methods | Get-EasySetupFingerprint, Get-EasySetupEmbeddedPayloadHash, Test-PpiPayloadMarker; build/publish/watch entry scripts and Assert-IdentityHash helper. |
| 5. Inputs and outputs | Current sources/POTX assets, approved certificate thumbprint, optional deployment version. Outputs signed payloads, EXE, build.json, README and historical build workspace. |
| 6. Dependencies | Visual Studio Office/MSBuild/Roslyn, VSTO targets/runtime, .NET4.8, Microsoft tool signature checks, certificate chain, Safe-Xml and local assets. |
| 7. Internal workflow | Mutex; source fingerprint; approved tool/certificate preflight; version increase; staged config template path; signed x86/x64 payloads; marker/manifest/template coverage; ZIP and inventory; bundle hash embedded in compiled bootstrap; Authenticode sign; require Valid and unchanged source; atomically replace public EXE/README/receipt. |
| 8. External integrations | Windows certificate store and Microsoft tool executables; no prerequisite/key download. Watcher polls local source and invokes same builder. |
| 9. Exceptions and error handling | Any tool/signature/source/manifest failure throws before publish. Mutex prevents concurrent build/validation. Watcher retains prior EXE and waits for next edit after failure. |
| 10. Configuration | x86/x64 always bundled; PpiEasySetupBuild prevents recursion; PpiForcePayloadCompile invalidates only intermediate assembly; PpiSkipEasySetup controls automatic hook. Watch poll2–300 seconds/default5. |
| 11. Security | Approved40hex certificate/private key/EKU/time/normal chain; Valid Authenticode required; no trust-store/policy/exclusion changes. Local Install-Local can create a nonexportable dev cert, which is not automatic enterprise trust. |
| 12. Tests | SetupPolicyTests, resource-reader atomic replacement/repeated validation/missing-resource tests, marker alignment tests, payload-recompile timestamp regression and Test-DemoPackage hashes/signature presence. |
| 13. Known limitations | Normal Windows PowerShell intended. ICU/NLS sort can alter fingerprint across hosts. Assembly.Load resource helper must receive already verified immutable bytes. Packaging does not prove installed ribbon/native acceptance. |
| 14. Technical debt | Related scripts retain implicit XML casts, dev certificate distribution remains estate-specific, no timestamp/renewal service, watcher needs active session. |
| 15. Extension points | Normalize fingerprint ordering with shared vectors; introduce approved timestamp/publisher lifecycle and transactional publication receipt policy without weakening freshness/signature checks. |

## 16. Complete owned source map

| Source | Analysis | Responsibility/declarations |
|---|---|---|
| `easySetup/Build-EasySetup.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `easySetup/Common.ps1:1` | Complete text read; no runtime execution | Declarations: Get-EasySetupEmbeddedPayloadHash, Get-EasySetupFingerprint |
| `easySetup/Watch-EasySetup.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `easySetup/Test-EasySetup.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `easySetup/Test-ResourceReader.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `easySetup/SetupPolicyTests.cs:1` | Complete text read; no runtime execution | Declarations: Check, Reject, Main |
| `scripts/Install-Local.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `scripts/Package-Content.ps1:1` | Complete text read; no runtime execution | Declarations: Test-PpiPayloadMarker |
| `scripts/Test-DemoPackage.ps1:1` | Complete text read; no runtime execution | Declarations: Assert-IdentityHash |
| `scripts/Test-PackageMarker.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `scripts/Test-PayloadRecompile.ps1:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |
| `apps/powerpoint-vsto/PowerPointIntelligence.csproj:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |

## 17. Handover and maintenance checks

Review the exact source/hash/line counts in `09-security/reviewed-files.csv` and this directory’s `source-module-mapping.csv`. Do not substitute historical PASS logs for current execution. Preserve unrelated module behavior and tests before modifying a boundary; new methods must have typed inputs, explicit error semantics and a corresponding negative test. See the dedicated AI/API/security/deployment documents for this module’s external contract.
