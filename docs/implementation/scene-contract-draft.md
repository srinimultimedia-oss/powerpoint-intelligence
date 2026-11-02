# Scene contract v1 — calculation boundary

TypeScript `scene-contract.ts` and `scene-builder.ts`, the published
`packages/chart-engine/schema/render-result.v1.schema.json`, and the matching
`src/Chart.Renderers/SceneContracts.cs` / `SceneCodec.cs` define a validated scene
projection for the S2 calculation families. Nodes are a
closed union of rectangle, ellipse, polyline, polygon, sector and text geometry,
with explicit point units, paints, chart-scoped identities and model revision.
The wire field is `semanticScene`, matching the baseline RenderResult interface;
the earlier unpublished TypeScript draft's `nodes` field is not accepted.
Unknown fields, nonfinite coordinates, cross-chart IDs, duplicate IDs and invalid
diagnostic references reject. The normalized-data builder covers the six current
calculation families; it does not migrate persisted v2 ChartDocuments.

The SHA-256 render hash includes all payload fields except the hash itself.
Its PPI-C14N-1 encoding uses tagged null/boolean/number/string/array/object
values, big-endian 32-bit lengths, big-endian IEEE-754 doubles (negative zero
normalizes to zero), UTF-8 without lone surrogates, ordered arrays and ordinally
sorted object keys. A length-prefixed `PPI-SCENE-1` domain string precedes the root.
The hash is an integrity check, not authorization, approval or a signature.
WebCrypto is supplied by the browser/Node runtime; an unavailable SHA-256 service
fails. Sealing snapshots input before awaiting asynchronous hashing.

## Exact validation and wire limits

The structural schema intentionally does not try to express cross-field identity
or hash checks. TypeScript and C# additionally require unique IDs under the exact
lowercase chart GUID, existing-node diagnostic references, and a verified hash.
Revision is a nonnegative Int32; this is a carried model revision, not a substitute
for the host's expected-revision check. Coordinates are finite points bounded to
±10,000,000; plot origins/dimensions are bounded to ±1,000,000 with positive sizes.
Zero-sized rectangles/ellipses are allowed. Polylines require two points, polygons
three, with at most 10,000 points. Sectors use clockwise degrees, a start angle in
[-360,360] and a sweep in (0,360]; full circles are explicit 360-degree sectors.
Paints contain only six-digit colors or null, stroke widths [0,1000] and opacity
[0,1]. Text requires a font choice, point size [1,1000], known alignment and valid
Unicode (4,096 code points maximum; 128 for a font name).

Wire decoding is capped at 8,388,608 UTF-16 code units and 32 value levels. Scene
arrays have at most 200,000 nodes / 10,000 diagnostics. JSON parsing rejects
duplicate keys, trailing commas, comments, single quotes, unquoted properties,
numeric extensions, nonfinite numbers, invalid escapes and lone surrogates. No
numeric-string coercion or polymorphic type-name deserialization is permitted.
All encode/decode paths revalidate. C# ComputeHash accepts an unsealed in-memory
draft; Encode still rejects it until a valid hash has been supplied.

PPI-C14N-1 uses these one-byte tags: null 0; false 1; true 2; number 3; string 4;
array 5; object 6. Strings are a four-byte UTF-8 byte count and UTF-8 bytes. Arrays
are a four-byte element count followed by values; objects are a four-byte member
count followed by sorted key-string/value pairs. The domain prefix is the raw
length-prefixed string `PPI-SCENE-1` (without a value tag). The committed primitive
golden has hash `9a8e3f13aa05f89a369e95bc4b0815f0407ab2851fb2e58beb868efcf4f42fea`.

## Canonical projection and reproduction

The Node-only local worker boundary `packages/chart-engine/tools/render-chart-v2.mjs`
strictly parses and validates the existing canonical v2 clustered-column profile,
retains its revision/layout/series colors, and emits a scene. It normalizes the
projection GUID to lowercase without changing the source. It explicitly reports
`AXES_NOT_RENDERED` and, when requested by the model, `LEGEND_NOT_RENDERED` rather
than claiming the current geometry projection renders those features. File I/O and
AJV remain outside the pure calculation package's source modules.

From the repository on the Windows development machine:

```powershell
npm ci --ignore-scripts --prefix packages/chart-domain
npm ci --ignore-scripts --prefix packages/chart-engine
npm run typecheck --prefix packages/chart-engine
npm test --prefix packages/chart-engine
.\scripts\Test-S1.ps1
.\scripts\Test-SceneContracts.ps1 -Platform AnyCPU
.\scripts\Test-SceneContracts.ps1 -Platform x86
.\scripts\Test-SceneContracts.ps1 -Platform x64
node packages/chart-engine/tools/render-chart-v2.mjs tests/s1/fixtures/clustered-column.v2.json artifacts/clustered-preview.scene.v1.json
```

The scene test runner compiles against the .NET Framework 4.8 targeting pack. It
uses the S1 pinned Newtonsoft dependency, an independent Node Buffer/SHA-256
oracle, a committed primitive golden, 140 shared fixtures, and a C# producer.
Current Node tests: 178; the C# runner defines 148 tests per architecture and 73
bidirectional golden/generated round trips plus the independently changed C#
scene. AnyCPU and x86 executions passed. The x64 runner compiled but failed to
launch and was subsequently missing; security event logs were inaccessible, so
the cause is unproven. Counts are not production acceptance claims.

## Remaining gates

Text metrics, layout/collision policy, clipping, full native rendering, object
reconciliation, expected-revision enforcement and approval/transaction execution
remain unfinished. A separate, unexposed clustered-column shape-renderer spike
now compiles and has core orchestration/metadata tests; its Office execution is
unverified. These APIs are not wired to the installed add-in or production
document store. The schema and hash confer no COM authority. AD-013 resolves the
representation choice; per-family renderer acceptance and Q-011 fonts remain
open. S0's intermittent native Office
failure remains open. This increment does not accept S3 or the full requirements.
