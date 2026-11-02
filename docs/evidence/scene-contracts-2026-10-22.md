# Shared scene contract increment — 22 October 2026

Implemented the baseline RenderResult boundary with `semanticScene`, model
revision, chart-scoped child identities, explicit point geometry, diagnostics and
a SHA-256 render hash. The closed geometry union supports rectangles, ellipses,
polylines, polygons, sectors and text. JSON Schema, strict TypeScript validation
and .NET Framework 4.8 DTO/validation now agree on the tested corpus. Unknown
geometry/actions, extra fields, duplicate JSON keys, numeric coercion, malformed
Unicode, out-of-range coordinates and stale hashes reject before DTO use.

The local Node worker projects the current canonical v2 clustered-column model
to the scene contract without migrating or editing the canonical source. It
preserves revision, layout, IDs and series colors. The projection reports
AXES_NOT_RENDERED and LEGEND_NOT_RENDERED explicitly. It performs no Office calls.

## Reproduced evidence

Final run: `artifacts/validation-20261022-152402-d79c4d2b`.
Source SHA256: `6AE329FB79B961F9C9356F681B1F6FE7C7AD724CF1C18662A220CCF962CDEC6F`.

- 178 Node calculation, preview, schema, semantic and strict JSON tests passed;
  these retain the 500 seeded signed-stack property cases.
- Strict TypeScript checking passed.
- 140 shared positive/negative scene fixtures passed in both AnyCPU and x86.
- 148 .NET tests passed per executed architecture, with 73 TypeScript → .NET →
  TypeScript scene round trips and an independently modified C# producer scene.
- The Node Buffer/SHA-256 oracle and .NET encoder agree on the executed hashes,
  including small/subnormal doubles, signed zero and Unicode. This is fixture
  evidence, not a proof over all IEEE-754 values.
- All existing source checks and AnyCPU/x86/x64 add-in compilation passed.
- Overall suite: **18 PASS, 1 FAIL**. The failure is x64 scene runner launch;
  no x64 scene assertions executed and no x64 execution pass is claimed.

Environment: Windows NT 10.0.26200.0; 64-bit PowerShell 7.6.5; Node 24.14.0;
.NET Framework Release 533509, compiled against the v4.8 targeting pack.

## x64 execution limitation

The x64 compiler returned success, and `compiled.json` recorded an executable
SHA256 of `2883C3FE0C5573EE0F28EB68D40F73597C489A143EDE37509D86B383184CB8E8`
before launch. Launch then failed with `0xfffffffe`; the runner was subsequently
missing. Code Integrity and Defender event-log reads returned access denied.
The root cause is not established, and this report does not attribute it to a
specific Windows policy or antivirus product. The failure remains visible in
the suite instead of being converted to a pass or replaced by compile evidence.

## Remaining scope

S3 remains incomplete: text measurement, clipping/overflow/collision policies,
native rendering, reconciliation and save/reopen evidence are missing. The
contract is not wired into the installed add-in and confers no execution or
approval authority. Target revision checks, permissions, transactions/rollback
and audit must precede production mutation. Renderer/font decisions remain open.
S0's intermittent native Office failure and the latest ribbon installation are
not resolved by these tests. No interactive Office tests ran in the final suite.
The five chart requirements retain partial status; none of the 98 requirements
is newly accepted by this increment.

Reproduction and exact wire/hash rules:
`docs/implementation/scene-contract-draft.md`. Standalone entry points:
`scripts/Test-SceneContracts.ps1` and
`packages/chart-engine/tools/render-chart-v2.mjs`.
