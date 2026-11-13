# Current validation and historical evidence

| Current check | Result | Cases | Passes | Start UTC | End UTC | Evidence |
|---|---|---|---|---|---|---|
| chart-core | PASS | 6 | 6 | 2026-11-13T05:51:18.733255+00:00 | 2026-11-13T05:51:18.900159+00:00 | 12-testing/current-validation/chart-core.log |
| chart-domain | PASS | 25 | 25 | 2026-11-13T05:51:18.900159+00:00 | 2026-11-13T05:51:19.170510+00:00 | 12-testing/current-validation/chart-domain.log |
| chart-engine | PASS | 178 | 178 | 2026-11-13T05:51:19.170510+00:00 | 2026-11-13T05:51:19.932274+00:00 | 12-testing/current-validation/chart-engine.log |
| typescript-noemit | PASS | typecheck | n/a | 2026-11-13T05:51:19.932274+00:00 | 2026-11-13T05:51:20.419689+00:00 | 12-testing/current-validation/typescript-noemit.log |

Additional **current** checks:17 S0 metadata assertions passed from unchanged source, compiled/run entirely in export-local temporary storage; see [metadata log](current-validation/metadata-tests.log) and `_analysis/run-metadata.ps1` in the full local export for the exact command. Fifty PowerShell files parsed without syntax errors; see [AST log](current-validation/powershell-ast.log) and [per-file results](current-validation/powershell-ast.json). Parsing executes none of those scripts. Metadata and AST evidence files' filesystem timestamps identify this export run; Node JSON records precise start/end UTC. Node runtime v24.14.0.

Total currently executed named assertions: **229 =6 calculations +25 domain/schema +178 engine +17 metadata +3 read-only policy controls**. Type-checking and50-file syntax parsing are separate successful checks, not226 more tests. No line/branch coverage, mutation testing or independent production packaging build was run.

**NOT EXECUTED now:** full Test-All, Office automation, installed ribbon, linked Excel native refresh, original template fidelity, renderer native save/reopen, x86/x64 host process integration, installer signing/trust/install/rollback, external service calls, scanner rescan, PIIDX corpus execution and security scripts that write outside export. None is a pass inferred from absence of a failure.

**Historical only:** docs/implementation/deployment-validation-20261103.md records24/24 local checks and desktop package1.0.0.18 validation. Existing artifacts contain older results, including failures during development. Such source records establish history; they are not13-Nov current execution evidence. Earlier Office successful focused assertions do not cover every chart family, every template, DPI/locale or Office estate. Historical Checkmarx remediation is not a fresh vulnerability scan.

Current read-only policy controls reused the pre-existing ordinary JSON and junction fixture from an earlier artifacts validation snapshot: valid JSON was read, artifact input rejected the junction, and fingerprint hashing rejected the junction before reading. Three assertions PASS, including the positive control; no fixture was created/modified. See [junction policy log](current-validation/junction-policy.log). This closes this narrowly scoped junction check for the inspected helper, not every path-policy caller or Windows filesystem race.
