# Desktop follow-up — 29 October 2026

Input: desktop payload 1.0.0.4, marker piidx2-blockers; evidence `artifacts/demo-features-68f6884a1a7044aa9f3ddf1eabf93b6f` and current scanner report.

The preceding text-measurement edit introduced a dynamic-binding regression: Shape.Application caused ActiveWindow/Selection access to remain dynamic and the returned integer selection type could not be compared with PpSelectionType. Application, DocumentWindow and Selection are now explicitly typed. This is a source fix; native success must be rerun.

Chart annotations/axis/series controls now progress through their native workflow assertions. The saved workflow deck has two internal slide-jump relationships to IDs 256/index 1 and 257/index 3. The save/reopen test no longer requires the third SubAddress display-name field to stay blank; it still requires an internal hyperlink, exact slide identity and exact index. Further native verification is required.

Additional changes:
- Connector no-op planning skips restoration snapshots and color/theme/adjustment reads when no change is planned. Its failing Saved-state test now records read/planning stages in connector-saved-state.txt, without forcing Saved=true or suppressing failure.
- Agenda registry verification uses the native namespace-selected collection with indexed access, verifies identity/payload and reports mismatch counts/flags without logging XML. This avoids scanning unrelated or orphaned parts. The prior native failure was preserved; this is not yet a confirmed native fix.
- Embedded workbook content is assigned in a single Value2 matrix with explicit text formatting for headers/categories, instead of queuing per-cell chart updates. Four pure assertions verify dimensions, headers, numeric/category typing, signed/zero/missing values and source immutability. Only explicitly rejected close calls are retried, within a 15-second budget with a message-pumping pause. Disconnections still require independent commit verification; otherwise the test fails.

## Scanner and runner review

The desktop report has 14 historical quarantine records, latest 23 October, and zero read errors. Prior AnyCPU, x86 and latest template-runner executable hashes do not match these records. Their missing outputs were explained by the verified merge exclusions. Explicit recovery audits preserved old receipts and rebuilt at the existing stable paths; no exclusion or antivirus policy was changed.

- AnyCPU scene: 148 assertions and 73 cross-language round trips pass.
- x86 scene: 148 assertions and 73 cross-language round trips pass.
- Template local checks: 36 antialias-policy assertions, 520 layouts, 20 palettes and 10,400 combinations pass; these are not Office fidelity results.
- x64 scene remains held: the legacy receipt lacks BuildKey and per-file hashes. No legacy adoption or binary restoration was performed. Explicit source-review recovery has been proposed to the owner, with historical hash refusal before execution.

Current source marker: `PPI-DEMO-RIBBON-20261029-v17-piidx3-typed-selection`. Reinstall before native rerun. Current ribbon contracts and demo test compilation pass. Production architecture build logs are in artifacts/blocker-followup-20261029. Native agenda, connector Saved-state, workbook close/commit and the repaired text/grid paths remain acceptance gates. These issues are not marked complete.
Latest complete non-interactive run: artifacts/validation-20261029-154954-82af3b9c/results.json — 22 PASS, one held x64 scene check reported FAIL. Its snapshot predates the final optional connector stage diagnostics; the final diagnostics/test compilation, ribbon suite and all three production compilations were then checked separately. No new native Office pass is claimed.

29 October recovery update: owner-approved x64 legacy source review is complete; the runner passes and reuses the same path. Latest local suite is **23/23 PASS** (artifacts/validation-20261029-160704-ee2c43de). See docs/implementation/x64-reviewed-recovery-20261029.md. This supersedes the earlier x64 hold; it does not establish current native Office acceptance.
