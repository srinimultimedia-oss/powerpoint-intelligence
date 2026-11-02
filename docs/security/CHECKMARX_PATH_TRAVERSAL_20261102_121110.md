# Checkmarx traversal follow-up — report 2026-11-02 12:11

## Confirmed scan evidence

Report: `Cx_ScanReport_2026-11-02_121110.md`; scan ID `6cd9bdd1-dcd1-4511-9b4f-f14ca67486e2`, full ASA Premium scan at **12:06 Asia/Dubai, 2 November 2026**; report created 12:11. 342 files / 314 SAST files scanned. No source commit SHA is supplied, so the report is not represented as proof of an exact local commit.

The report explicitly lists **87 resolved C# traversal records: 85 Path_Traversal and two Stored_Path_Traversal**. Every original similarity-ID occurrence from the prior 11:36 assessment was compared with the resolved tables; all 87 match, including repeated IDs. This is scanner evidence of those records resolving, not merely local-test inference.

The report still has **8 High results**: seven detailed PowerShell records plus one summary-only SCS result. Its summary identifies five integrity-check findings, **one path-traversal finding**, and one untrusted-functionality finding. All seven details are incorrectly grouped together under `Inclusion of Functionality from Untrusted Control Sphere`, CWE-829, with no descriptions. Both original detail taxonomy and summary classification are preserved in the JSON companion.

**Correction to the earlier assessment:** the prior 11:36 review counted the 87 explicit C# traversal records but incorrectly treated all seven PowerShell detail records as non-traversal. That earlier summary also identifies the collector as traversal. Thus that report contained **88 traversal-classified records: 87 C# plus one PowerShell**. The collector record was missed in that scope classification; it is addressed here. No result is suppressed or silently reclassified as resolved.

## Remaining traversal record and source fix

- Original similarity ID: **-160022997**, recurrent High.
- Original detailed query/CWE: Inclusion of Functionality from Untrusted Control Sphere / **829**.
- Summary classification: Improper Limitation of a Pathname to a Restricted Directory (**Path Traversal**, CWE-22 terminology).
- Original source: `scripts/Collect-SceneThreatEvidence.ps1:1`, `$OutputDirectory`.
- Original sink: same script, line 89, `Set-Content -LiteralPath $reportPath`.
- Root cause: caller-selected output could create directories and overwrite the fixed report filename anywhere the user could write. Canonicalization alone would not establish an authorized output boundary.

Implemented controls in `scripts/Diagnostic-OutputPath.ps1` and the collector:

1. Select an independent data root: repository `artifacts` or an OS temporary **subdirectory**. Relative output paths resolve against the repository; absolute destinations under those roots remain supported. The temp root itself and arbitrary outside directories are rejected. Normal `-OutputDirectory artifacts/scanner-...` commands remain compatible.
2. Reject parent/current traversal, drive-relative/root-relative ambiguity, devices/extended paths, ADS, reserved names, controls/wildcards and trailing-dot/space aliases before creating output or reading scanner history.
3. Canonicalize, compare complete directory boundaries, and reconstruct the requested child from `Path.GetFileName` components under the independent root; reject any location change.
4. Reject existing reparse ancestors and the fixed report destination. `GetAttributes` is attempted directly rather than relying on Exists, which can miss dangling links. Recheck the fixed `scene-threat-evidence.json` destination immediately before writing. The unsafe reassignment from raw `$OutputDirectory` at the old sink was removed.
5. Preserve JSON structure, normal report replacement, scanner collection/conclusion logic, read-only security posture and permission-error reporting. No protection settings, trust rules or scanner exclusions were changed. These checks rely on existing directory ACLs; they do not claim handle-based protection against namespace races or all hard-link attacks.

## Executed verification

Final source SHA-256: `7871C33660B6ADBD55973B103F1858528CBCACCA7E673C690BE52F813EB502A2`.

- **31 new diagnostic-path assertions pass**, including a call to the actual collector with malicious output. No scanner history is read and no target output is created on rejected entry. Valid nested, spaced and absolute artifact paths and temp descendants remain accepted. `scripts/Test-DiagnosticOutputPath.ps1` is included in `Test-SecurityRemediation.ps1`.
- Actual valid collector execution produces readable JSON at `artifacts/diagnostic-output-tests/functional/scene-threat-evidence.json`. Collection completed with 14 historical McAfee detections, zero Defender detections/events/read errors; this is not a false-positive verdict or proof that protection is active. Evidence: `artifacts/diagnostic-output-functional.log`.
- Fixture junction creation was denied by this AppContainer and explicitly skipped. Separately, the actual guard successfully rejected the existing `C:\Users\All Users` Windows reparse point without creating links or changing policy. Evidence: `artifacts/path-review-121110-reparse.log`.
- Full local suite: **23 PASS / 1 FAIL**, evidence `artifacts/validation-20261102-122044-01d9bb98/results.json` and `source-snapshot.json`. Current source fingerprint matches the captured source. The only failure is the historical missing PIIDX cached DLL; its recovery gate remains enabled. All architecture builds/compatibility, C#/scene/template/ribbon/contracts and security checks pass.
- Existing 30 C# path-authority assertions, 28 corpus-name assertions and 22 offline repair security tests remain passing. Evidence: final aggregate `security-remediation.log` and `artifacts/path-review-121110-security.log`.

PowerShell tests ran in the available PowerShell session. An ordinary Windows PowerShell 5.1 test attempt returned `AuthorizationManager check failed` before execution (`artifacts/path-review-121110-windows-powershell.log`); no execution-policy bypass was attempted. Windows PowerShell 5.1 runtime acceptance and native Office testing are not claimed. Application source/ribbon behavior was not changed in this follow-up. No dependency upgrade, Checkmarx suppression/custom sanitizer declaration, antivirus exclusion or Windows policy change was made.

## Acceptance status

**Prior 87 C# traversal records: scanner-confirmed resolved in the supplied report.**

**PowerShell record -160022997: fixed in source and locally verified; confirming Checkmarx rescan pending.** A rescan must show the summary traversal count and this original flow cleared, including review of any new path records. The six other detailed findings and summary-only SCS result remain outside this traversal-only task; full-report compliance is not claimed. The missing PIIDX runner and deployment freshness remain separate gates. No quarantined/missing runner was restored or executed to avoid them.
