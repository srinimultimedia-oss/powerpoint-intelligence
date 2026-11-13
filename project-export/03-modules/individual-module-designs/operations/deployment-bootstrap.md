# OPS-04 — Per-user deployment bootstrapper

**Evidence category:** source verified; execution status is stated separately. Baseline includes the actual inspected checkout, not an assumption that every working byte equals HEAD.

## Design

| Dimension | Implementation-grounded description |
|---|---|
| 1. Purpose | Verify/extract a packaged bundle and invoke normal VSTO install/upgrade/rollback with visible user consent. |
| 2. Implementation status | Implemented managed WinForms EXE entrypoint; desktop publication evidence is historical. |
| 3. Internal components/classes/interfaces | IntelligenceSetup static class; SafePath, Digest, OfficeArchitecture, RunVsto and STA Main. |
| 4. Important functions/methods | SafePath rebuilds restricted archive path; OfficeArchitecture reads registry and PE; RunVsto waits for process exit; Main executes phases. |
| 5. Inputs and outputs | Embedded IntelligencePayload.zip plus compiled BundleHash. Outputs per-user version cache/log/install receipt and VSTO registration via official installer. |
| 6. Dependencies | Windows registry/PE/file APIs, .NET4.8, ZipArchive/SHA256, WinForms, installed VSTOInstaller and desktop Office 16. |
| 7. Internal workflow | Require closed PowerPoint and prerequisites; identify one x86/x64 host; user confirmation; bundle/inventory/per-file hash checks; bounded extraction to bundle-hash version; validate prior receipt; uninstall changed old manifest; install new; write receipt; attempt prior install on failure. |
| 8. External integrations | Reads HKLM App Paths/Office InstallRoot; launches VSTOInstaller /Install and /Uninstall; no admin service/download. |
| 9. Exceptions and error handling | Phase+exceptiontype+HRESULT logged; message shown; returns1 failure/2 cancelled/0 success. Modified existing version file or invalid prior receipt is refused; rollback failure separately logged. |
| 10. Configuration | LocalAppData/PowerPointIntelligence/easySetup/versions/<bundlehash>; installed-manifest.txt; bootstrap bound<=5000 entries/512MiB expanded/2MB inventory. |
| 11. Security | SafePath rejects absolute/backslash/colon escape; exact archive/inventory/file hashes; no overwrite of modified versions; normal publisher prompt preserved. |
| 12. Tests | SetupPolicyTests eight assertions cover path escapes and known digest. Resource tests concern package reader helper; no current native install test. |
| 13. Known limitations | No macOS/web/mobile/MSI/server deployment, automatic prerequisite installation, forced host shutdown or global trust management. PE major16 is not all-build compatibility proof. |
| 14. Technical debt | No central update channel/enterprise publisher lifecycle; installation cache integrity depends on ordinary workstation policy; rollback depends on old version files. |
| 15. Extension points | Add verified distribution metadata/operational target-matrix testing while retaining per-user ownership, signed manifests and refusal to close user work. |

## 16. Complete owned source map

| Source | Analysis | Responsibility/declarations |
|---|---|---|
| `easySetup/Bootstrapper.cs:1` | Complete text read; no runtime execution | Declarations: Digest, SafePath, OfficeArchitecture, RunVsto |

## 17. Handover and maintenance checks

Review the exact source/hash/line counts in `09-security/reviewed-files.csv` and this directory’s `source-module-mapping.csv`. Do not substitute historical PASS logs for current execution. Preserve unrelated module behavior and tests before modifying a boundary; new methods must have typed inputs, explicit error semantics and a corresponding negative test. See the dedicated AI/API/security/deployment documents for this module’s external contract.
