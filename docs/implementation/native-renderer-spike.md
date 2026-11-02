# S3 grouped-shape renderer feasibility increment

AD-013 records the owner's approval of editable grouped native shapes for custom
chart families, with retained canonical data. The representation choice in
Q-009/OD-06 is resolved. Runtime editability and persistence remain evidence gates.

This increment implements the first insertion profile: **canonical v2 clustered
columns projected as editable PowerPoint rectangles in a native group**. It is a
feasibility profile, not a replacement for the current S0 native chart, a complete
renderer, or an accepted chart feature. The ribbon's explicit local sample command
now calls this adapter; installed command behavior remains unverified. Other
scene geometry types, other canonical chart families, text metrics,
labels, axes, legends, overrides, updates, undo and copy identity repair remain
unfinished. Broader approved shape composition remains the intended outcome.

## Boundaries and invariants

`NativeRenderingCore.cs` validates canonical JSON and scene JSON before mutation.
It checks the same chart/revision/layout, the complete ordered set of non-null
category/series IDs, expected signed column geometry (relative/absolute tolerance
1e-9), and canonical colors. Thus a correctly hashed scene for different values
or geometry cannot silently be paired with the canonical dataset. Axes/legend
omissions must be explicitly reported by the projection. The canonical model is
retained, not reconstructed from drawing coordinates.

The profile has a proposed safety guard of 2,000 semantic nodes and a minimum
2-point plot dimension. This does not settle the production sizing decision.
Zero-height bars have no native rectangle and retain an explicit null shape-ID
binding; missing cells have no semantic bar. Two invisible boundary anchors keep
empty/single-bar charts groupable. Unrelated shapes are never included in the
native range. Shape **names**, not IDs passed as collection indices, build the
PowerPoint group.

The core uses an explicit surface interface, prevents reentrancy, and rechecks
owner STA access. Each insertion tracks its owned objects. Failure attempts to
remove only those objects and reports `RollbackComplete=false` if cleanup fails
or a COM creation outcome is uncertain. This is not a claim that every Office
failure can be rolled back. Permission, approval, optimistic concurrency, audit,
transaction history and native Undo policy still require the command coordinator
before production mutation.

`PowerPointSceneSurface.cs` uses native PowerPoint rectangles and groups. It stores
uppercase SHA-256 node tokens rather than case-sensitive semantic IDs in tags,
because Microsoft's [COM Shape.Tags reference](https://learn.microsoft.com/en-us/previous-versions/office/office-12/ff761092%28v%3Doffice.12%29)
documents uppercase storage of tag values. Current build behavior is still a
runtime verification item.
The exact semantic IDs live in metadata. Shape IDs are hints: verified reads
reconcile them through node tokens and detect missing/duplicate tokens, duplicate
chart identity, moved/resized bars and style drift. Ambiguous or manually modified
objects are preserved and rejected rather than overwritten. Group/ungroup and
copy identity repair are not implemented by this spike.

The grouped-object APIs follow Microsoft's documented [ShapeRange.Group](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shaperange.group),
[Shape.GroupItems](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.groupitems)
and [CustomerData.Add](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.customerdata.add)
boundaries. Documentation supports platform capability; it does not prove this
adapter's runtime behavior. The proposed native-read geometry tolerance is 0.05
points for the Single-precision COM API and still needs visual/Office evidence.

## Self-contained versioned persistence

The group owns one CustomerData XML part in
`urn:powerpoint-intelligence:rendered-chart:v1`. It contains the validated canonical
v2 JSON, validated scene v1 JSON and ordered semantic-ID/token/nullable-native-ID
bindings. No external database, source workbook or service is needed to decode
it. `rendered-chart.v1.xsd` supplies the structural contract; the codec also
validates the embedded JSON, numeric projection, identities and mapping semantics.
DTD/entities, extra XML nodes/attributes, unsupported versions, oversized metadata,
duplicate bindings and incorrect tokens reject. This is separate from the S0
XML profile; existing presentations are not migrated.

## Reproduction and current evidence

From the repository, after the S1 dependency has been restored:

```powershell
.\scripts\Test-NativeRenderer.ps1
.\scripts\Test-NativeRenderer.ps1 -Platform x64
.\scripts\Test-NativeRenderer.ps1 -InteractiveOffice
```

The first command compiles the actual COM adapter and executes 41 core tests
plus seven embedded ribbon-sample preparation tests. The explicit platform option
also verifies the reported process architecture and saves an execution log and
source/executable hashes under artifacts/native-renderer-x64.
These include before/after failure injection at nine mutation stages, incomplete
rollback, non-owner threads, reentrancy, canonical/scene disagreement, metadata
round trips, XSD validation, malformed mappings and bounded XML parsing. Fake
surface tests verify orchestration only, not COM behavior.

The interactive runner refuses to operate while user presentations are open. It
owns only its new test presentation and never quits PowerPoint. Its assertions
cover native grouping, zero-bar bindings, unrelated-object preservation,
save/reopen, token mapping, a native edit saved and reopened a second time,
preserved canonical metadata with explicit divergence detection, duplicate identity rejection and
native rollback after an injected persistence failure. These assertions are
implemented but **not yet passed**: the current session returns access denied
0x80070005 when activating PowerPoint, before any presentation is created.

The add-in is not reinstalled or promoted to S3 acceptance by these checks.
