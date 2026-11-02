# Native renderer feasibility evidence — 22 October 2026

The owner approved grouped editable PowerPoint shapes for custom chart families.
AD-013 records that decision and resolves the representation choice in Q-009.
The full 98-requirement objective remains active; this increment does not narrow
the planned feature set or promote the demonstration build to production.

Implemented the first native shape insertion profile: canonical v2 clustered
columns projected as native rectangles in a group. The core checks canonical
values, ordered semantic identities, revision, layout and colors against the
scene before mutating. It skips zero-area bars while retaining bindings, preserves
unrelated objects, prevents reentrancy, checks owner STA access, and attempts
owned-object rollback with an explicit incomplete/uncertain result on failure.

Added self-contained, versioned CustomerData metadata containing canonical JSON,
verified scene JSON and semantic-token/native-ID bindings. A structural XSD and
semantic codec reject unsupported versions, malformed/ambiguous mappings, DTDs
and oversized input. Native IDs are hints; tokens support read-time mapping.
The COM adapter compiles and detects duplicate identity and native geometry/style
drift without overwriting the affected objects. These COM behaviors remain
unverified at runtime.

## Executed checks

Final source snapshot:
`artifacts/validation-20261022-154807-c45d7697`.
SHA256: `C5A8387A4DDA60F74D56630E7DC8AA0B2FF823611991F44DB63259299F32926E`.

- 41 renderer core tests passed, including before/after fault injection at nine
  mutation stages, unrelated-object preservation, incomplete rollback reporting,
  non-owner threads, reentrancy, numeric/model disagreement, XML round trips and
  XSD validation. These use a fake surface and do not prove COM rollback.
- Existing 178 Node tests and strict TypeScript passed.
- Shared scene contract tests passed in AnyCPU and x86 (148 tests and 73
  cross-language round trips per executed architecture).
- Ribbon contracts, demo compilation, S0 unit, S1 domain, schema/calculations,
  compatibility policy and add-in compilation in all three architectures passed.
- **19 of 20 noninteractive checks passed.** The x64 scene contract runner again
  failed to launch with `0xfffffffe`. Its execution remains unverified, with no
  established root cause. Compilation evidence is not counted as an execution pass.

Separate interactive attempt: `artifacts/native-renderer-check.log`.
The actual renderer compiled and its 41 core tests passed, then PowerPoint
activation returned `UnauthorizedAccessException`, HRESULT `0x80070005`.
No presentation was created and no native Office assertion passed. The owned-
presentation runner implements grouping, save/reopen, mapping, native-edit drift,
duplicate identity rejection and persistence-failure rollback assertions, but
they remain pending desktop execution.

## Acceptance limits and next work

The renderer is a spike outside the ribbon and installed add-in. It currently
accepts the canonical clustered-column profile only. Other families/geometry,
text metrics, data labels, axes, legends, updates, undo, manual override capture,
copy identity repair and per-family capability acceptance remain unfinished.
The 2,000-node/minimum-size guards and 0.05-point native read tolerance are proposed
feasibility parameters, not approved production limits. S3 and CHT-001 are not
accepted by these tests. S0 native stability and current ribbon deployment also
remain unresolved.

Reproduction, contract boundaries and Microsoft API references are in
`docs/implementation/native-renderer-spike.md`. Run
`scripts/Test-NativeRenderer.ps1 -InteractiveOffice` only from an interactive
desktop with user presentations saved and closed; it never quits PowerPoint.
