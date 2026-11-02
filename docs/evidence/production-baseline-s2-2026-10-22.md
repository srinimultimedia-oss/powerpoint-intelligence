# Production baseline and S2 increment — 22 October 2026

Active objective: the latest supplied architecture/functional baseline contains
98 requirement IDs and S0–S10 execution order. The prior generic technical
blueprint and broad demonstration do not replace this baseline. Windows desktop,
.NET Framework 4.8 in-process VSTO, on-prem services, Azure OpenAI only, canonical
models and native editability remain fixed. Section 18 permits S1/S2 deterministic
work alongside the still-open COM feasibility work. No production release gate
is marked accepted here. Exact signed 100% semantics and other OPEN decisions
remain unresolved; no database, SSO or Azure production policy was invented.

Source archived in docs/requirements/architecture-functional-goal-2026-10-22.md.
All 98 IDs are in production-requirement-traceability.json. Five chart IDs have
partial calculation evidence; none is marked fully accepted.

S2: independent TypeScript clustered column/bar, mixed-sign stacked column/bar,
positive-only 100% stacking, line paths with explicit gap/baseline policy,
zero-baseline area polygons and single-series nonnegative pie angles. Runtime
validation enforces IDs, dimensions, known fields, finite numbers and points.
Raw values remain distinct from normalized/imputed render values. Nulls, zero
totals, overflow and unsupported policies are explicit. No PowerPoint, browser
chart library, database or model inference is used by this calculation package.

Evidence: TypeScript 5.9.3 strict/noUncheckedIndexedAccess/exactOptionalPropertyTypes
checks pass. Twenty tests pass, including independent numeric references,
immutability, stable semantic IDs and 500 seeded mixed-sign property fixtures.
Executed under Node 24.14.0. No preview/commit SLO is claimed. This is an internal
calculation result, not the final S3 scene execution contract or a migration of
the S1 v2 persisted clustered-column schema. Native renderer, text metrics,
labels, interaction and durable model wiring remain further slices.

Test/stability work retained because it supports baseline sections 12/15:
- Complete 90-command desktop test scenarios now supply fixture prompt/file inputs
  and use command-specific assertions plus native save/reopen. They compile but
  have not executed here; PowerPoint activation is denied in this session.
- 18 cancel paths and eight invalid numeric/input cases pass before COM mutation.
  NaN font sizes and NaN Excel CAGR inputs are now rejected.
- Installed ribbon log from validation-20261022-141745-f5af4921 proves a matching
  installed DLL, actual ribbon OnLoad and 90 bindings for that installed revision.
  It does not prove the subsequently edited source is installed or visually correct.
- The same desktop run recorded native Excel/PowerPoint crashes. S0 instability
  is not fixed by passing other runs. Automatic selection/window callbacks no
  longer create/dispose panes or perform broad metadata recovery; explicit user
  actions rebind panes. This is hardening, not a verified crash root-cause fix.
- Office suites refuse to mutate fixtures when user presentations are open.
- Validation captures source/dependency copies and a checked source fingerprint
  before running, preventing concurrent edits from producing mixed-source runs.
  Signing/installation stops if source changes during build. Loaded-source and
  loaded-assembly hashes must match the install receipt.

One intermediate snapshot omitted TypeScript's node_modules/typescript/bin and
failed typecheck due to missing compiler entry point. Snapshot copying now
preserves dependency bin directories while excluding application outputs and
key/certificate/env files. Final suite result is recorded separately below.

Final local suite: validation-20261022-144237-06f299b9, source fingerprint
30BF775C919674991C6B6702DF2A26CBB14E431EC71EA9CB1F9B5B6D9D8414BF.
All sixteen noninteractive checks passed, including S2 numerical tests/typecheck,
existing domain/schema/compatibility checks, compiled 90-command ribbon contracts,
90-case feature runner compilation and add-in Compile for AnyCPU/x86/x64.
No interactive Office check ran in that suite. This is not a signed-install,
full-feature, S0 crash-resolution or 98-requirement completion claim.

## Deterministic SVG inspection increment

Added `packages/chart-engine/src/svg-preview.ts` for all six calculation families.
It escapes accessible title content, emits fixed geometry only, preserves missing
line/area breaks, and renders a single positive pie value as a full circle.
This is internal S2 inspection output, not the persisted S3 scene schema or a
PowerPoint renderer. Axes, labels, legends, editing handles and theme parity are
not implemented by this preview.

Validation: 24 calculation/preview tests passed (including the existing 500 seeded
stack fixtures), strict TypeScript passed, and all 16 noninteractive checks passed
in `artifacts/validation-20261022-144744-9bcb9f20`. Snapshot source SHA256:
`E04A5C3E8FE0F99D1DFABD9D952B6CEE6AB3B2856B406DEFDFCF1D70C60515EA`.
An example was generated at `artifacts/s2-preview.svg`; visual inspection has not
been performed. Interactive Office tests were not run. S0 runtime instability and
production acceptance remain open.
