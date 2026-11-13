# OPS-06 — Activation and scanner diagnostics

**Evidence category:** source verified; execution status is stated separately. Baseline includes the actual inspected checkout, not an assumption that every working byte equals HEAD.

## Design

| Dimension | Implementation-grounded description |
|---|---|
| 1. Purpose | Collect bounded native activation/token/registration and selected scanner facts without pretending that file disappearance proves antivirus causation. |
| 2. Implementation status | Read-only diagnostic logic implemented; report writing has filesystem side effects and was not executed by this leaf. |
| 3. Internal components/classes/interfaces | ActivationToken P/Invoke helper; Diagnose-PowerPointActivation; Collect-SceneThreatEvidence; Diagnostic-OutputPath; Report-ReadFailure. |
| 4. Important functions/methods | GetTokenInformation/GetTypeFromProgID/CreateInstance/Application.Version; Get-PpiDiagnosticOutputDirectory/ReportPath/AssertDirectPath; filtered McAfee/Defender event extraction. |
| 5. Inputs and outputs | Days1–14, scanner selector, output path; outputs activation.json or scene-threat-evidence.json with Errors and explicit unverified verdict. |
| 6. Dependencies | Interactive Windows identity/token/registry/COM, McAfee detection log, optional Defender module/CIM/EventLog and path policy. |
| 7. Internal workflow | Validate scanner output authority; inspect product/log/event/receipt evidence; retain only matching fields; record read errors; write fixed report. Activation reports resolution/instance/version stages and releases RCW without Quit. |
| 8. External integrations | Windows token/ClassesRoot registry; PowerPoint COM activation; scanner logs/Defender commands. No protection settings edited/elevation requested. |
| 9. Exceptions and error handling | Per-source read failures are recorded; no-match is not false-positive verdict/protection proof. COM stage exceptions/HRESULT unwrap; existing user app never Quit. |
| 10. Configuration | Scanner McAfee default; Days3 default/max14. Collector output under repository artifacts or OS temp subdirectory; strict component/device/ADS/reparse rejection. |
| 11. Security | Rejects path traversal before directory creation; selected event fields only, not unrelated messages/commandlines/accountdata. Matching regex currently scene/template only, not PIIDX. |
| 12. Tests | Test-DiagnosticOutputPath validates traversal/device/ADS/output roots and actual entry rejection; junction fixture conditional on host permission. Activation requires desktop environmental comparison. |
| 13. Known limitations | Denied scanner/event access limits conclusion; products installed does not prove active protection. Token/registry diagnostics do not repair COM permissions. |
| 14. Technical debt | PIIDX path/hash collection absent; output schemas not formally versioned; full read-only description should distinguish local report writes from no policy mutation. |
| 15. Extension points | Parameterized bounded runner match list/positive fixtures, explicit schema version and causal desktop-token comparison; never infer clearance from empty results. |

## 16. Complete owned source map

| Source | Analysis | Responsibility/declarations |
|---|---|---|
| `scripts/Collect-SceneThreatEvidence.ps1:1` | Complete text read; no runtime execution | Declarations: Record-ReadFailure |
| `scripts/Diagnostic-OutputPath.ps1:1` | Complete text read; no runtime execution | Declarations: Assert-PpiDiagnosticDirectPath, Get-PpiDiagnosticOutputDirectory, Get-PpiDiagnosticReportPath |
| `scripts/Diagnose-PowerPointActivation.ps1:1` | Complete text read; no runtime execution | Declarations: Read |
| `scripts/Test-DiagnosticOutputPath.ps1:1` | Complete text read; no runtime execution | Declarations: Check, Reject |

## 17. Handover and maintenance checks

Review the exact source/hash/line counts in `09-security/reviewed-files.csv` and this directory’s `source-module-mapping.csv`. Do not substitute historical PASS logs for current execution. Preserve unrelated module behavior and tests before modifying a boundary; new methods must have typed inputs, explicit error semantics and a corresponding negative test. See the dedicated AI/API/security/deployment documents for this module’s external contract.

## Scanner time-range nuance

Days bounds the Windows event query. The current McAfee branch searches matching detection.log lines without a date cutoff; a freshly generated report can intentionally include older matched detections. Report-generation freshness is not detection freshness or complete runner coverage.
