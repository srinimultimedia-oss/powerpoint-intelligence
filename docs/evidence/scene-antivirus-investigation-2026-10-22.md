# Scene x64 runner: missing executable and reported antivirus alert

22 October 2026. Status: investigation open; no restore, exclusion or verified antivirus fix.

## Evidence

- `artifacts/scene-contracts-x64/SceneContractTests.exe` is absent. Its last compilation receipt records SHA256 `EA5C34B16761FD30AC9ADF0F8CB4DE362B1AA81A93DD9D170D06EA0A087FA7A3` at 16:03:26. The immutable full-suite snapshot records a separate build hash `72BE5A0B44966754BC2D7911861419415A40489490314A420B265F3800AE7F1E`. Do not confuse builds when matching the scanner's affected path or hash.
- The last x64 attempt compiled successfully, failed to start with ResourceUnavailable / 0xfffffffe, and subsequently had no executable present. File disappearance alone does not prove quarantine. The owner now reports an antivirus threat alert; its exact scanner, detection name, affected resource and remediation action have not yet been supplied or retrieved.
- The compiler's Authenticode signature validates as Microsoft Corporation. The local Newtonsoft dependency's Authenticode signature also validates. The compiled local test artifacts are unsigned. Signature/provenance details and hashes are in `artifacts/activation-diagnostics/compiler-dependency-provenance.json`. These checks do not establish a malware-free verdict or identify a signature trigger.
- Source review of SceneContractTests.cs and its compiled contract/parser sources found no process launch, network retrieval, registry modification, self-deletion or dynamic assembly loading. The runner reads fixture files, validates closed DTOs, hashes scenes and writes round-trip JSON. Newtonsoft TypeNameHandling is None. Static review cannot establish whether a scanner detection is a false positive.
- The AnyCPU scene control passed 148 contract tests and 73 TypeScript/.NET/TypeScript round trips. Its launch.json records exit 0 and the executable remaining present. This is a control result, not substitution for the failed x64 check.
- The explicitly x64 native renderer runner passed all 41 core tests and seven ribbon preparation tests. `artifacts/native-renderer-x64/execution.json` records target x64, source/executable hashes and exit 0. This disproves a universal inability to execute x64 .NET test programs in this environment; it does not explain the scene runner detection. Its Office attempt still fails at activation, before drawing assertions.

## Diagnostic access and changes

This agent has a restricted AppContainer token. Antivirus product/detection CIM queries and protected event log reads are denied. Importing the Windows Defender module in PowerShell 7 also fails because its Windows PowerShell remoting compatibility path is unavailable. A direct Windows PowerShell 5 read-only collector attempt was rejected by AuthorizationManager before the script ran. No execution policy, security log permission, antimalware policy or service setting was changed.

`scripts/Collect-SceneThreatEvidence.ps1` is a read-only local collector for an approved normal desktop session. It gathers:

- registered scanner names and Defender signature/engine version;
- matching detection IDs, threat names, resource paths, action/status codes and times;
- relevant Defender 1116/1117/1118/1119/1121/1122 and Code Integrity event fields;
- matching build receipts from the current directories and five recent immutable validation snapshots;
- explicit read failures. Zero matching records with read failures is not a clean scan.

It exports no unrelated events, full event messages, command lines, raw dataset contents or account names, performs no scans/restores, and sends nothing to a vendor. The first agent-context collection is `artifacts/scene-threat-20261022-161846/scene-threat-evidence.json`; its five read failures mean it contains no scanner verdict.

`scripts/Test-SceneContracts.ps1` now retains launch.json and execution.log, including pre/post presence, exit code and exception chain. A missing executable produces a diagnostic warning rather than an implied antivirus diagnosis. The executable source and compiler flags were not changed to evade detection. The revised harness was verified through the successful AnyCPU control; the x64 suspect was not regenerated again while awaiting its detection record.

Microsoft documents detection and remediation events separately: [Defender event IDs](https://learn.microsoft.com/en-us/defender-endpoint/troubleshoot-microsoft-defender-antivirus), [Get-MpThreatDetection](https://learn.microsoft.com/en-us/powershell/module/defender/get-mpthreatdetection). These records, not the missing-file symptom, must identify the reason/action.

## Next evidence gate

From normal approved Windows PowerShell in the repository:

```powershell
.\scripts\Collect-SceneThreatEvidence.ps1
```

Supply the generated scene-threat-evidence.json. If history access is unavailable, the owner can inspect the detection's exact threat name and affected items in [Windows Security Protection History](https://support.microsoft.com/en-us/windows/security/windows-security/protection-history-in-the-windows-security-app), or the equivalent scanner UI. The computer-use skill explicitly prohibits automating antimalware apps, so this UI step remains with the owner.

Correlate the detection resource with a specific recorded build and time. Investigate the named detection's source/dependency/behavior evidence. Remediate an actual issue if identified; if a false positive is supported, use the scanner vendor's review process with explicit approval before uploading any artifact. A new executable hash or lack of a later alert by itself does not prove remediation. Do not restore, rename, exclude, allowlist, suppress, replace the x64 test with AnyCPU, or skip it to manufacture 20/20. After a supported resolution, rebuild from reviewed source in the same target, rerun the x64 scene and renderer tests, confirm scanner status, then run the complete suite. All renderer acceptance criteria remain open until actual Office/ribbon/edit/save/reopen evidence is available.

## Update at 21:08 Dubai time: McAfee cause confirmed, protected retest passes

This update supersedes the earlier unknown-scanner and pending-detection statements above. The owner identified McAfee Total Antivirus; installed product metadata reports McAfee 16.3.1.108. Its readable `C:\ProgramData\McAfee\wps\detection.log` contains six matching records. The collector now defaults to McAfee and exports only matching diagnostic fields; use `-Scanner Defender` or `-Scanner All` for other investigations.

The root x64 receipt hash `EA5C34B16761FD30AC9ADF0F8CB4DE362B1AA81A93DD9D170D06EA0A087FA7A3` exactly matches McAfee's 12:03:34 UTC / **16:03:34 Dubai** record: detection `ti!EA5C34B16761`, sensor `section execute`, final detection source `rp-s`, action `infection quarantined`. The immutable full-suite hash `72BE5A0B44966754BC2D7911861419415A40489490314A420B265F3800AE7F1E` likewise matches a 16:04:27 Dubai quarantine record, detection `ti!72BE5A0B4496`. Thus **McAfee quarantine caused these executables to disappear during launch**, rather than a failed compiler or a missing source file. Curated evidence: artifacts/activation-diagnostics/mcafee-confirmed-quarantine.json.

The log does not disclose the internal reason for `rp-s`, and no authoritative vendor definition for that internal code was established. Do not infer malware family, low reputation or a proven false positive from the generic detection name. Source review and local signatures are supporting evidence, not a vendor verdict.

The owner confirmed that the first successful retry followed pausing protection. That initial 20/20 result at artifacts/validation-20261022-200545-68a50fa3 is therefore functional evidence under paused protection, not security remediation evidence. The owner then confirmed McAfee active again. The **same executable hash** E02CDBF2FEFB13AD09CB45D78750391FDC49A38AE46541A29D94A4E6BD96B17E was executed without rebuilding: 148 tests passed, exit 0, executable remained present. Receipt: artifacts/activation-diagnostics/scene-x64-protected-retry.json.

A subsequent fresh-build full run, **artifacts/validation-20261022-204800-08191651**, passed **all 20 local checks** with owner-confirmed active protection. It also attempted all four Office integration checks; all failed at COM activation 0x80070005. Source snapshot SHA256: EA3A7E46C070A98E698EBBD39BEC3CD3E3FCAEB39800424ECC3990789ECC2239. The fresh x64 runner hash F141AB9796AD09C6E0B63CD3B1DE321C2394BB4C7C222F954DFA076FCAB47990 passed 148 tests plus 73 cross-language round trips and remained present. Renderer core: 41/41 plus seven ribbon preparation tests pass. The McAfee log still contains the same six matching detections, with no later matching quarantine through the final read at 21:08 Dubai; read errors were zero (artifacts/scene-threat-20261022-210832/scene-threat-evidence.json).

No agent action disabled protection, added an exclusion, restored quarantine, renamed the runner, changed its source/compiler flags to alter detection, skipped x64 execution or uploaded an artifact. The missing file has been rebuilt and successfully retested in its original location. **Current x64 execution is verified; a lasting vendor-classification correction is not established.** A vendor review draft is provided separately, not submitted. Office activation, installation of the new renderer build and its native-group round trip remain open gates.
