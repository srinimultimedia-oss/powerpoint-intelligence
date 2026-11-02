# S2 basic chart calculations

Pure TypeScript numerical/geometry increment for CHT-001, CHT-002, CHT-010,
CHT-012 and CHT-013. No Office, DOM, AI, storage or network access. Runtime
validation rejects unknown fields, duplicate IDs, dimension errors, coercion and
nonfinite values. Inputs are normalized data, not a persisted ChartDocument.

Commands (Node 24; tested with 24.14.0):

```
npm ci --ignore-scripts
npm run typecheck
npm test
```

TypeScript 5.9.3 is pinned in the lockfile. Node executes the erasable TypeScript
syntax directly for tests; a separate strict `tsc --noEmit` check is mandatory.

## API and declared semantics

`Dataset` contains identified categories and identified series with finite
`number | null` values. `Plot` is a positive-sized rectangle in explicit points.
IDs use the existing S1 semantic-ID format. Child render IDs are local to the
scene; a future host must namespace them under the canonical chart GUID.

- `clustered(dataset, plot, orientation, gapRatio)` supports columns and bars,
  stable category/series order, signed values around zero and preserved null slots.
- `stacked(...)` accumulates positive and negative segments independently.
- `positive100(...)` is explicitly positive-only. It retains raw values alongside
  normalized render values and reports zero/empty categories. Signed inputs fail
  with `SIGNED_100_POLICY_REQUIRED`; OD-12 is not silently resolved.
- `line(dataset, plot, {gapPolicy, includeZero})` requires explicit gap/baseline
  choices. Gap policies are break, connect, or zero with tagged imputation.
- `area(dataset, plot, {gapPolicy, baseline:0})` closes each contiguous path to
  zero. This increment supports straight segments and a zero baseline only.
- `pie(dataset, plot, startAngle)` requires one nonnegative series. Zero totals
  produce an empty result and diagnostic. The final sector closes at 360 degrees.

Numeric model: IEEE-754 doubles; reference tests use relative/absolute tolerance
1e-9. Nonfinite intermediate sums/ranges fail rather than returning invalid
geometry. Degenerate all-zero domains use [-1,1]; constant nonzero line domains
without zero receive 10% or one-unit padding. This is an explicit prototype
linear-axis policy, not implemented nice ticks, breaks, logarithmic/time axes,
text layout, rendering, legend positioning or production financial rounding.

The output `sceneVersion: 1` is a calculation result with local semantic IDs and
point geometry, not the final S3 execution schema. No S1 v2 persisted contract is
changed, and no legacy metadata is implicitly migrated. Other chart families,
canonical host integration remains later work. A separate versioned scene contract
now has matching schema/C# DTOs and cross-language tests.

## Evidence limits

178 tests include exact reference fixtures and 500 deterministic seeded
mixed-sign/property cases. They verify numeric behavior, not PowerPoint output,
font metrics, visual layout, file persistence, undo or latency SLOs. Source tests
are separate from the existing demonstration commands and do not make those
commands production-ready. The full 98-requirement goal remains incomplete.
# Internal SVG preview

`src/svg-preview.ts` renders the six calculation families as deterministic SVG strings in points. It accepts validated datasets, escapes the accessible title, and emits only fixed SVG geometry elements. Missing line/area values break paths. A full-circle pie uses a circle rather than a degenerate arc. The preview uses a fixed palette and zero-inclusive linear scales; it does not yet include axes, labels, legends, themes, or selectable editing handles. It is an internal S2 inspection aid, not the S3 scene contract, a native PowerPoint renderer, or acceptance evidence for Office runtime behavior.

`scene-contract.ts` and `scene-builder.ts` are typed projections with SHA-256
integrity checks, a closed JSON Schema and matching .NET Framework 4.8 validation.
140 shared fixtures cover invalid syntax/semantics and exact cross-language hashes.
`tools/render-chart-v2.mjs` is a Node worker entry point from the existing canonical
v2 profile; it reports the unimplemented axes/legend projection explicitly.
Native rendering, reconciliation and full persistence evidence remain pending.
See `docs/implementation/scene-contract-draft.md` for exact limits and hash rules.

