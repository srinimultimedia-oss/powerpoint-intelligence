# Comprehensive validation — 22 October 2026

201 unit-test executions pass: 27 metadata/vector/retry checks, 35 C# domain
checks, 25 JSON contract checks, six calculation checks and 36 compatibility
policy checks under each of AnyCPU, x86 and x64. C# JSON output cross-validates.
Add-in Compile targets pass for AnyCPU, x86 and x64. git diff --check passes.

Interactive Office activation from this Codex session fails before fixture setup
with 0x80070005. This does not reproduce the desktop chart failures or prove any
runtime fix. The latest owner run stopped at Excel workbook.Close with explicit
0x80010001 rejection; a bounded, narrowly scoped retry is implemented and tested.
Original-metadata-after-duplication and cross-document transfer remain unverified.

A full Release x64 Build fails because ClickOnce manifest signing is not selected.
No signed package, installed add-in load, ribbon/task pane, full PowerPoint
process restart, or Office 2016/2019/365 product matrix is certified by this run.

scripts/Test-All.ps1 -InteractiveOffice is now available for a desktop session.
It executes all suites, compiles all three architectures, attempts the Office
sequence three times and writes per-check logs plus results.json/results.md in
artifacts/validation-<unique run>. It stops retrying activation on access denial.
Office assertions after any failed assertion remain untested, not passed.
The runner does not change security settings, install certificates, or close
presentations owned by the user. No additional S1 feature work was undertaken.

Machine-readable evidence: artifacts/validation-20261022-114844-6ff62325/results.json
