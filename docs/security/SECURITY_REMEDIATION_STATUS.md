# Feature-preserving security remediation — 2 November 2026

The owner approved fixes that do not impact existing features after Run 1. This implementation is limited to the confirmed corpus-output and offline-repair weaknesses plus the related publication gap. It does not implement unapproved enterprise path/trust policies, suppress findings, alter dependencies or claim a new Checkmarx scan.

## Original finding traceability

| Original record | Original severity / CWE | Change | Verification status |
|---|---|---|---|
| CX-R007; Similarity **889475886**; Stored_Path_Traversal Result 4/5 | High / CWE-22 | Data-derived corpus names pass `tests/s3/SceneArtifactPath.cs` before `SceneContractTests.cs:76` writes them. Parent/root/drive/UNC/ADS/device/control/trailing-dot-or-space names and duplicate/case-colliding names are rejected. Caller-selected output roots and valid existing fixture names are retained. | Implemented and locally verified; scanner revalidation pending. |
| CX-R080; Similarity **-2077706329**; Path_Traversal Result 72/107 | High / CWE-22 | Same underlying sink and control as CX-R007; both original records are retained. | Implemented and locally verified; scanner revalidation pending. |
| CX-R116; Similarity **-900544832** | Medium / CWE-606 | `Repair-TemplateAlignment.py` delegates bounded I/O/publication to `RepairTemplateSafety.py`; the alignment algorithm retains its geometry formulas, with budget checks before container/label matching and text loops. ZIP member/size metadata and actual bytes, incremental XML node/depth counts, aggregate work and elapsed time are bounded. | Implemented and locally verified; scanner revalidation pending. |
| CX-R117; Similarity **-1745516950** | Medium / CWE-606 | The repair validation pass uses the same resource context and explicit entity/network-disabled parser. Text/effect/untouched-part/source comparisons use explicit exceptions rather than removable asserts. | Implemented and locally verified; scanner revalidation pending. |
| Additional review **AR-03**, not an original Checkmarx ID | Proposed Low / defensive correctness | Existing deck and sidecar files are refused. Work is staged in owned same-directory temporary files; only fully verified files are published without replacement. A failed second publication removes the deck just published by this operation; temporary files are cleaned on exceptions/cancellation. The two files are not claimed to be one filesystem-atomic transaction. | Implemented and locally verified, including optimized Python and induced failures. |

All original report IDs/severities remain unchanged. Four reported records address **two** root causes; they are not four distinct production vulnerabilities and are not marked resolved in Checkmarx without a rescan/reviewer confirmation. The original 118 result count remains a historical scan count, not a new scan result.

## Compatibility and resource budgets

No `apps/powerpoint-vsto` or production `src` implementation was changed. Native ribbon/rendering, Office profiles, linked-data mapping, palettes, original template insertion, task pane and persisted element schemas remain unchanged. The existing repair geometry/alignment formulas and CLI `--input` / `--output` are retained. Keep `RepairTemplateSafety.py` beside `Repair-TemplateAlignment.py`; the utility uses the existing Python/lxml dependency. Python 3.9+ is supported by the added test workflow.

Default per-package repair limits: 512 MiB source file, 10,000 ZIP members, 64 MiB central directory, 64 MiB/member, 512 MiB total inflated data, 40 MiB/slide XML, 1,000,000 XML nodes, depth 128, 25,000,000 aggregate work units, 4 GiB processed bytes across passes and 120 seconds elapsed. A trusted programmatic caller may provide an explicit positive `Limits` object; document input never controls these limits. Malformed/ambiguous members, split packages and over-budget work fail safely instead of silently truncating. Extreme valid files outside these bounds are now explicitly rejected; no claim of compatibility with every formerly unbounded input is made.

Calibration on all 62 existing source packages found maxima of 418 members, 99,020,946 inflated bytes and 32,154,334 bytes for one member. The final defaults accept the entire existing library. XML node limits apply while parsing; central-directory records are counted before constructing ZipFile entry objects; the quadratic matching loop charges its work before evaluating candidate labels. Elapsed checks occur between bounded operations, not as an OS-level preemptive timeout.

Path checks intentionally restrict the data-derived child filename, not all user-picked source directories. Physical ancestor/junction/race hardening (AR-02) remains conditional and is not claimed solved by lexical containment. The output directory remains a caller-authorized test workspace.

## Validation evidence

- `scripts/Test-SecurityRemediation.ps1`: **28** Windows C# filename-policy assertions against the actual helper, **22** Python security/preservation tests. Includes a small compressed-bomb fixture, input mutation refusal, explicit output/sidecar no-clobber, induced validation/publication failure, alignment geometry and preservation under optimized Python. Initial tests demonstrated unsafe-input acceptance and sidecar overwrite on the baseline; the final tests pass.
- Existing stable scene runners were rebuilt at their same paths using the established explicit reviewed-source procedure. Their prior executable/library/dependency hashes matched the receipts. Existing historical scanner evidence `artifacts/scanner-blocker-review/scene-threat-evidence.json` was used; this is **not a new antivirus scan or vendor clearance**. No detection/legacy gate was removed and no alternate executable path was used to bypass one.
- `artifacts/security-scene-AnyCPU.log`, `security-scene-x86.log`, `security-scene-x64.log`: each reports **148** .NET contracts and **73** Node → .NET → Node roundtrips, plus a real-runner malicious-corpus test that leaves its outside-output sentinel unchanged. Actual x86/x64 process architecture is printed in the logs. These are native Windows CLR contract tests, not PowerPoint COM tests.
- `artifacts/security-repair-library-final-20261102/results.json` and `artifacts/security-repair-library-final.log`: **62 packages / 497 slides** accepted under the final safety implementation; original source hashes, text, native timing/transitions and all untouched package parts preserved. This is offline package/XML validation, **not native rendered fidelity or Office reopen evidence**.
- `easySetup/Test-EasySetup.ps1 -SourceOnly`: **8** setup-policy assertions and source checks pass. Installer signature, installation and upgrade acceptance are not implied.
- Full local-suite result and final source hash are recorded below after completion. No new Checkmarx scan or native Office UI/COM test was executed in this agent session.

### Final local acceptance

`artifacts/validation-20261102-102816-d85c5823/results.json`: **24/24 local checks PASS**, including the integrated 28 + 22 security assertions, renderer/domain/calculation/library checks, AnyCPU/x86/x64 scene contracts and all three production compilation targets. The validated source snapshot SHA-256 is **F31ED025A44CED0B8FE8EB122D3E73B886186200D04B07930C6D8BBC10545382**. The preceding validation run used an earlier intermediate snapshot and is not substituted for this final evidence.

Comparison against the Run 1 baseline confirms **97 existing production add-in/rendering/domain/library and dependency files are byte-identical**. No package versions or lock files changed. This supports feature preservation; it does not turn offline XML tests into native Office rendering or establish a new scanner result.

The installer rebuild was attempted with the owner's existing thumbprint and stopped before packaging because the AppContainer certificate store cannot see it: `artifacts/security-easySetup-build.log`. The prior published EXE was not replaced or represented as current. Rebuild in normal desktop PowerShell using the existing approved signer, then validate freshness/signature. No private key was exported or trust-store fallback added.

## Deployment and artifact freshness

`Source-Fingerprint.ps1` now includes Python source files. `easySetup/Common.ps1` includes the whole source fingerprint in addition to add-in and packaging/template inputs, so tooling/test/security changes cannot leave an old installer falsely reported current. Signed Release builds and the optional edit watcher still generate the installer/README. Restart an already-running watcher once to load its updated fingerprint functions.

The previous easySetup EXE is not represented as freshly rebuilt by this session. Build/sign from normal desktop PowerShell with the approved certificate, then run `easySetup/Test-EasySetup.ps1`; do not bypass its freshness/signature gate. The runtime add-in payload has no UI/feature changes, but source/package provenance must still reflect this revision.

Python bytecode is ignored and excluded from source snapshots; this run's generated bytecode was removed from the index after verifying it did not exist in HEAD. No unrelated staging or commits were changed.

## Deliberately still open

- **110 potential false positives** require reviewer/threat-model confirmation, not blanket file restrictions or scanner suppression.
- **Three conditional provenance findings** (CX-R001/002/003) need the build/cache/CI actor and approved toolchain policy. Existing checks remain enabled.
- The **one High SCS summary** still lacks its original ID/rule/location and needs a redacted raw result. The download/untrusted-functionality taxonomy discrepancy also remains unresolved.
- Additional **AR-01 dependency integrity** requires approved independent package identity/hash/signature material; the current version was not upgraded or given an invented local checksum. **AR-02 physical filesystem identity** needs an approved hostile-directory model and Windows integration coverage.
- Full native Office acceptance, a freshly signed easySetup installer and a Checkmarx rescan remain separate evidence gates. No remaining finding is silently waived, and this partial remediation is not full security acceptance.

## Latest traversal-only follow-up — 2 November 2026, report 11:36 Asia/Dubai

The original counts above describe the earlier baseline and must not be substituted for the latest scan. The latest report contains 85 Path_Traversal and two Stored_Path_Traversal records (87 High/CWE-22), seven other detailed records and one summary-only SCS result. Reconstructed workspace paths, scalar CLI inputs and independent legacy-root approval are implemented and locally verified. Full evidence and all current original IDs: `CHECKMARX_PATH_TRAVERSAL_20261102_113655.md`, its JSON companion, and `PATH_TRAVERSAL_VALIDATION_20261102_113655.md`.

Final local suite: 23 PASS / 1 held PIIDX missing-assembly failure, source `02FDCB26D95E6A574E59C67E8A8C9585D1BFF7AD2F69C584E93D6F5C536F1C99`. No Checkmarx suppression/exclusion, dependency upgrade, trust change or antivirus-policy change was made. Checkmarx rescan, reviewed PIIDX recovery and desktop Office acceptance remain open; full compliance is not claimed.

## Traversal follow-up — report 12:11, 2 November 2026 Asia/Dubai

The full 12:06 Checkmarx scan explicitly resolves all 85 Path_Traversal and two Stored_Path_Traversal C# records; original ID occurrences match the previous register. One PowerShell traversal record (-160022997) remains in the summary despite being grouped as CWE-829 in details. The earlier 11:36 scope classification missed that record; its corrected traversal total was 88, not 87 across all languages. This collector-output issue is now constrained to independently owned artifacts/temp paths and tested. See `CHECKMARX_PATH_TRAVERSAL_20261102_121110.md` / JSON. Source `7871C33660B6ADBD55973B103F1858528CBCACCA7E673C690BE52F813EB502A2`: 31 diagnostic assertions and full local 23 PASS / 1 held missing-PIIDX failure. Confirming collector rescan remains open; six other detailed findings and the SCS summary are not claimed resolved. No security controls were disabled.

## Integrity/untrusted-code follow-up — report 13:42, 2 November 2026 Asia/Dubai

The supplied scan reports zero active traversal and six detailed High PowerShell records (five integrity-check flows and one untrusted-functionality flow), plus one unidentified SCS summary. Added independent deterministic source proof and locked/byte-snapshot execution for cached PIIDX/scene/template code; exact certificate selection, explicit code-signing chain policy and authenticated Microsoft tool roles; locked source hashing. Two source-fingerprint flows are read-only hashes rather than downloads/execution and require scanner/reviewer validation as potential false positives, not suppression. All six original IDs and remaining gates are retained in `CHECKMARX_INTEGRITY_20261102_134253.md` / JSON.

Final source `6D4B966CB5F7A07544FE087842E325B14E06319FD208ED721CEC6B32D8737915`, aggregate `artifacts/validation-20261102-145526-7c15431e`: 23 PASS / 1 held missing-PIIDX failure. Twenty integrity assertions and setup source checks pass; signed private-key publishing/full PIIDX runtime require desktop/recovery acceptance. No missing/quarantined code restored, dependencies upgraded, or trust/antivirus/Checkmarx controls disabled. The owner supplied the two script paths but no distinct 2ms rule/ID; the SCS summary remains unverified. Confirming scanner results remain required.

## Fingerprint-flow review — report 09:13, 3 November 2026 Asia/Dubai

The scan explicitly resolves all six prior PowerShell ID occurrences and reports one new High SAST record (1072025071, CWE-494) plus the unidentified 2ms summary. The remaining SAST source is enumeration in Get-PpiAddinFingerprint; its recorded sink is Get-PpiContentHash, whose implementation opens a read-only stream and returns SHA256. No download/load/execution of the selected file appears in this chain. Nine targeted regression assertions demonstrate digest correctness, mutation detection, exclusion behavior and no execution of inert source content. Existing security tests remain passing. No production rewrite, query exclusion, suppression or finding-state change was made. The SAST record remains To Verify, assessed as a potential false positive requiring Checkmarx reviewer/query validation; the SCS rule still lacks details. Evidence and exact original record: `CHECKMARX_FINGERPRINT_20261103_091304.md` / JSON. Full compliance is not claimed.
