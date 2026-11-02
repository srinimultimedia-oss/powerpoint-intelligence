# Approved x64 source recovery — 29 October 2026

Owner explicitly approved a fresh source-reviewed recovery at the existing stable path, with prior receipt preservation and refusal of any hash found in the McAfee evidence. This is a new build from reviewed source, not adoption or restoration of a legacy/quarantined binary.

The normal cache path still refuses legacy receipts. The additional `-ReviewLegacyReceipt` path requires x64, `-RebuildReviewedRunner`, `-RecoverMissingRunner`, absent output files, a bounded legacy receipt with both prior hashes, a current source/compiler/reference build key and a complete read-only scanner report with no read errors. Prior path/hash detections block recovery. A separate plain-JSON audit retains the exact original receipt and its SHA256 before compilation.

After compilation, every new output is checked against recorded detection paths/hashes before a new receipt is written or execution is allowed. Reuse requires the retained audit and unchanged scanner report; normal cache hash/build-key checks remain active. No antivirus exclusion or Windows execution/security policy changed.

## Executed evidence

- `artifacts/blocker-followup-20261029/legacy-review-policy.log`: 33 inert policy assertions pass, including refusal of detected prior/new hashes, Windows-prefixed detection paths, incomplete reports and existing outputs. No fixture executable was built or executed.
- `artifacts/blocker-followup-20261029/scene-x64-reviewed.log`: genuine Windows x64 runner executes successfully; 148 scene contracts and 73 cross-language round trips pass.
- `artifacts/blocker-followup-20261029/scene-x64-reuse.log`: the same runner is reused successfully without rebuilding and repeats the passing tests.

Stable executable: `artifacts/scene-contracts-x64/SceneContractTests.exe`.
SHA256: `6C70DECAE12564CCC9E4990EA8F70F3D4A91A74EC8B810706E0177399F546052`.
Current source/toolchain build key: `726EF2B9084840DC89C96224FDAA6196DBA9763FFCF2CB90F77725818D798553`.
Scanner evidence SHA256: `C59AB16A91EF3C4B21C64A1735E874B7110A5E61FF95E4BBACCA1BCDBC681AC2`.
Preserved prior receipt audit: `artifacts/scene-contracts-x64/compiled.json.legacy-source-review-57c12081a2504124b9b7239427da290d.json`.

The new hash is absent from the 14 supplied historical detections. That observation and successful execution are not a vendor false-positive verdict or proof of antivirus protection status. Current native PowerPoint fixes still require desktop installation and interactive regression evidence; clearing this runner does not close those acceptance gates.
Full local regression after recovery: **23/23 PASS**, source snapshot D3F7ED711211D23ADEF3CF2C5E113C7BC588AC74CE64E7923F3838049C296E26. Evidence: artifacts/validation-20261029-160704-ee2c43de/results.json. Interactive Office checks were not requested in this run; native acceptance remains open.
