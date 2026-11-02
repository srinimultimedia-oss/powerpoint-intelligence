# Path-traversal validation — 2 November 2026

Implementation and report mapping: [111-record register](CHECKMARX_PATH_TRAVERSAL_20261102.md) and its JSON companion.

Final source fingerprint: `9DB40FCBAE5EEF6D6D30A6D569BE24D3114739FFA1B3A65E48138DA5922E6CEB`.

## Executed checks

- `scripts/Test-All.ps1`: **23 PASS / 1 FAIL**. Evidence: `artifacts/validation-20261102-112554-3fe490fc/results.json` and `source-snapshot.json`. Source fingerprint matches the current working source. The failure is the missing PIIDX cached test assembly, not a passing/waived PIIDX result.
- Security remediation: **24 actual path-authority assertions**, 28 existing corpus-name assertions and 22 offline repair security tests pass. Evidence: `artifacts/path-policy-security.log` and the final aggregate `security-remediation.log`.
- Scene contracts: **148 assertions and 73 cross-language round trips on each of AnyCPU, x86 and x64**. The real runner rejects malicious corpus traversal without outside-output mutation. Existing executable names were retained; reviewed rebuilds preserved provenance and scanner checks. Evidence: `artifacts/path-policy-scene-AnyCPU.log`, `path-policy-scene-x86.log`, `path-policy-scene-x64.log`.
- Template library: 36 fidelity-policy assertions, 520 layout previews and 10,400 layout/palette combinations pass offline. Evidence: `artifacts/path-policy-templates.log`. These do not establish native Office fidelity.
- Actual gallery and Office CLI executables reject `..\outside` with exit 2 before file access / COM activation. Evidence: `artifacts/path-policy-cli.log`. Existing mode flags are retained.
- Ribbon contracts, source/data model tests, compatibility on all three architectures, JSON schema, chart engine/typecheck, native renderer core, cross-language checks and all three production compile targets pass in the final aggregate run.

## Remaining blockers and acceptance

`artifacts/piidx/PiidxTestSuite.dll` is absent although `artifacts/piidx/test-suite-build.json` retains the historical receipt (29 October). Its existing recovery guard correctly refuses rebuilding automatically. Fresh evidence `artifacts/scanner-path-review/scene-threat-evidence.json` has 14 historical McAfee records, zero Defender records/read errors, and **no PIIDX detection**. The cause of disappearance is not verified; do not label this a confirmed antivirus quarantine or disable protection. A reviewed recovery must retain the prior receipt, inspect the affected path/hash and preserve the cache policy before the PIIDX tests can run again. A diagnostic library compilation was performed to check compilation only; it was not executed as an alternate PIIDX test runner.

Native Office UI, document editing, linked workbook refresh and save/reopen were not executed here. Desktop regression is required because of the known AppContainer activation restriction. An installed old package is not evidence for this source.

A new Checkmarx scan is required before claiming report compliance or changing original records to resolved. Compare all 111 original traversal records, examine residual/new paths and keep the eight non-traversal detailed records and summary-only supply-chain issue separate. No scanner exclusions/suppressions or Windows security policy changes were added. The older signed installer is stale; rebuild and sign after source approval using the existing deployment workflow.
