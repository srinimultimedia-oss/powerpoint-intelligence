# Remaining-blocker remediation — 29 October 2026

Status: source fixes implemented and local regressions run; desktop Office acceptance and held-runner review are pending. No antivirus exclusions or Windows security policies were changed.

## Evidence and changes

| Failure | Evidence | Change / verification boundary |
|---|---|---|
| JSON/schema, engine, TypeScript, cross-language and renderer prerequisites | Desktop log reports missing `ajv` and `tsc`; both packages had no node_modules dependencies | Restored exact npm lockfiles with lifecycle scripts disabled. Added `scripts/Restore-LockedDependencies.ps1` for future clean/moved workspaces. All corresponding local checks pass. |
| Scene/template runners | New-folder caches contain receipts but no runner executables/libraries. The workspace merge script excludes .exe/.dll files and deliberately carries receipts. | Missing files are explained by migration omissions; this is not evidence of a new quarantine. Source/toolchain receipts also differ for some runners. Added the missing explicit `-RecoverMissingRunner` scene option, requiring `-RebuildReviewedRunner` plus scanner evidence and prior hashed receipts. No held binary was restored, built or executed. Known detections still require clearance; exclusions are not clearance. |
| Mapped Excel refresh and chart annotations | Both desktop stack traces fail casting COM enumerator objects to PowerPoint.Series (0x80004002) | Replaced COM foreach enumeration with one-based typed SeriesCollection.Item access. The audit also removed the same pattern from value-axis validation and the series-type editor. Actual chart operations require desktop rerun. |
| Process-grid metadata ambiguity | Desktop reports two canonical text parts after PickUp/Apply style transfer | Replaced PickUp/Apply with explicit native font, alignment, margins, theme/RGB, solid/transparent fill and line formatting transfer. Source attachments are not copied. Gradient/texture/pattern fills are refused explicitly; originals remain preserved. Existing Office grid tests verify canonical ownership and source metadata preservation after installation. |
| Text formatting assertion | Desktop combined assertion failed without identifying which native property differed | Probe duplication/deletion now restores the prior native shape/text selection. Added non-content diagnostics for actual fill, outline and alignment. Selection disruption is a candidate cause, not a verified root cause until native rerun. |
| Connector no-op marks deck dirty | Desktop no-op Saved-state assertion fails; canonical endpoint reconciliation previously rewrote every native float observation | Retain prior coordinates for the same binding within 0.1 point; real movement and native rebinding still update. Added three pure regression assertions. Native Saved-state outcome remains unverified. |
| Agenda registry deletion | Desktop fails at old CustomXMLPart.Delete with a stream-about-to-be-deleted error | Registry update now verifies its original attachment, detaches first, creates/verifies the final replacement and restores the original payload on failure using the tested replacement transaction. This removes add-before-delete ordering; causal confirmation and native success require rerun. |
| Chart workbook close / commit | Desktop reports disconnected workbook and a failed native commit check | Reacquire shape.Chart for post-close commit checks rather than relying on the pre-close chart RCW. Disconnected-close success still requires independently verified native values. No failure or timeout is suppressed. Native outcome remains unverified. |

Payload marker: `PPI-DEMO-RIBBON-20261029-v17-piidx2-blockers`. Reinstall before assessing source changes through the installed add-in.

## Executed checks

`artifacts/validation-20261029-152708-7fbfdda4/results.json`: **19 PASS / 4 FAIL**. All four failures are held scene/template runner prerequisites. This was a non-interactive run; no native Office assertion is inferred from it. Its snapshot predates the subsequent scene recovery-option wiring, dependency-restoration helper and final audit of equivalent series-enumeration sites. The final chart-control edits were separately revalidated by Test-Ribbon.ps1 and all three production compilations; see ribbon-final.log and build-*.log.

- PIIDX seven groups: 143 assertions pass.
- JSON schema: 25 pass; chart engine: 178 pass; typecheck and cross-language pass.
- Renderer core: 41 pass; ribbon preparation: 9 pass.
- Ribbon/element contracts pass, including 165 connector assertions and 33 XML/locale assertions.
- Production AnyCPU, x86, x64 compile passes. Separate logs: `artifacts/blocker-fixes-20261029/build-*.log`.
- Subsequent targeted runner policy: **23 assertions pass**; scene recovery script parses. Fixtures are inert. Log: `artifacts/blocker-fixes-20261029/cache-policy.log`.

The owner's preceding desktop run already passed installed-ribbon and all three S0 Office checks. Those results apply to the preceding payload; they do not validate these edits. This agent's previously confirmed AppContainer COM activation boundary prevents independent execution of native acceptance here.

## Required desktop evidence

With McAfee active, run from the new repository:

```powershell
.\scripts\Collect-SceneThreatEvidence.ps1 -OutputDirectory artifacts\scanner-blocker-review
.\scripts\Install-Local.ps1
.\scripts\Test-DemoFeatures.ps1 -InteractiveOffice
```

Complete the publisher prompt and keep the owned test PowerPoint window in the foreground when requested. Report the printed artifacts path. The scanner report must be reviewed before invoking any held-runner recovery; no extra exclusions are recommended from missing-file evidence alone. After approved recovery, rerun the complete interactive suite. Do not mark full acceptance complete while these gates remain open.