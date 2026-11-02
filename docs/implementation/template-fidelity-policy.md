# Template fidelity: bounded raster policy

Owner approval: 28 October 2026, AD-018. Applies only to raster comparisons in original-template validation. Product slide generation, source assets and formatting are unchanged.

## Acceptance rule

A raster pair passes if decoded pixels are identical, or **all** the following hold:

- Dimensions are identical. No resizing, blurring, regional masks or image registration.
- Changed pixels are at most `min(8, floor(width × height × 0.00005))`. The area limit is **0.005%**, with a hard cap of eight pixels.
- Each RGB channel differs by no more than **10/255**.
- Alpha is exactly unchanged at every pixel.
- Every changed pixel lies on an existing reference edge: at least one neighboring reference pixel within one pixel differs from that reference center by 16 or more in an RGB channel. Flat-area edits do not qualify.

The rule is fixed and not auto-expanded to fit a failing image. The comparator is bounded to two million pixels. A single high-contrast changed pixel, ninth changed pixel, pixel shift, alpha change or dimensions mismatch fails.

The unchanged source's repeated export must meet the same raster policy. Native geometry and formatting must match exactly; raster tolerance cannot override those checks. CSV evidence records exact-pixel equality separately from tolerance acceptance, plus changed-pixel count, channel maximum, alpha changes, non-edge changes and allowed count. A tolerant pass is never reported as byte-exact equality.

## Native checks

The Office runner compares insertion and reopened slides against their source snapshot. Bounds, rotation, object/group order and text-content hashes remain exact. Additional native fingerprints include flips, AutoShape kind/adjustments, persisted DrawingML path/curve commands, control points and guides, visible fills (including gradient type/stops), visible line styles/colors, picture crop/brightness/contrast, text frame margins/anchors/orientation/autosize, and per-character font/paragraph properties. These are explicit native property checks, not a claim that every PowerPoint feature or proprietary extension has been exhaustively inspected. Existing shape/build counts, transitions, source-file hashes, provenance, native text editing and persistence assertions remain required.

Per-character formatting snapshots are limited to 50,000 characters; shape traversal to 2,000 shapes/depth8, and persisted XML parts to 32 million characters. Unsupported COM properties fail rather than silently omitting a check. Formatting diagnostic files contain property values, not slide text; text content is hashed by the geometry recorder.

## Executed evidence

Eleven synthetic comparator assertions pass: exact equality, eight subtle edge changes, rejection of ninth-pixel and delta11 changes, alpha changes, flat-area edits, size mismatch, global recoloring, one-pixel translation, high-contrast edits and the fractional-area limit on a small image.

Offline replay of all **48** saved source/insert/repeated-source/reopened image pairs from `validation-20261028-103116-c3ea917e` passes the approved raster policy. Metrics: `artifacts/fidelity-policy-replay/raster-replay.csv`. Previously failing original6/original9 insertion/repeated source differ by two pixels, maximum delta6; reopening differs by eight pixels, maximum delta10. Replay is not a new Office execution or a pass for the newly added native formatting/path checks.

The test runner is rebuilt at the same `artifacts/template-library-x64/TemplateLibraryTests.exe` path using the existing explicit reviewed-rebuild procedure. Fresh read-only McAfee evidence (`artifacts/scanner-fidelity-policy/scene-threat-evidence.json`) had no read errors and no detection matching that stable path or then-current hash. Historical detections remain retained. The build produced a plain-text audit of the old receipt; no security guard, antivirus preference/exclusion or receipt provenance rule was weakened. The unrelated x64 scene runner remains held.

## Desktop verification

The comparator and native assertions compile into the existing stable runner. No add-in reinstall is required specifically for this test-only policy. If the earlier elements2 source is not yet installed, its installation remains a separate pending step.

Save/close personal presentations, keep antivirus active, and run in normal Windows PowerShell:

```powershell
.\scripts\Test-TemplateLibrary.ps1 -InteractiveOffice
```

Inspect `artifacts/template-library-x64/original-fidelity.csv`, native geometry/format snapshots and launch status. Native execution remains pending because this agent's confirmed AppContainer cannot activate PowerPoint COM. Approval of a tolerance does not itself mark the library accepted.

## Final local validation

Full suite: C:\Users\srini\Documents\Codex\2026-10-21\new-chat\work\powerpoint-intelligence\artifacts\validation-20261028-105726-604d658a/results.json. Source SHA256: C85E3A272263B57DD6EC73EB7BDB3533D72118696585E46E98B5A4C481A5483D. **20 PASS / 1 FAIL**; only the pre-existing x64 scene provenance hold remains. Template-library local checks, new comparator assertions, ribbon contracts and all architecture compile checks pass. The actual stable template runner compiled/executed local catalog checks successfully after explicit reviewed source rebuild. The new native Office fingerprint checks are compiled, not executed.

The prior saved-image replay is reproducible through the existing DemoContractTests.exe --fidelity-replay mode, with its input evidence folder and output folder as arguments. No alternative executable is used for native template tests. Current source matches the final suite snapshot and git diff --check passes.


## EditingType E_FAIL repair — 28 October 2026

Owner desktop run installed deployment1.0.0.30/elements2 successfully, then failed on the first Cubes original during ShapeNode.EditingType with HRESULT0x80004005. This was a new test fingerprint failure; no successful native library result is inferred from installation.

The [Microsoft EditingType reference](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shapenode.editingtype) describes this getter for vertices/control points. It does not establish why this imported source object rejected it. The confirmed defect is the fingerprint's assumption that every imported freeform node exposes that COM property; Office corruption or a particular node kind is not claimed as its cause.

The fingerprint now avoids ShapeNode.EditingType, SegmentType and Points getters. Instead, it makes an owned temporary copy with [Presentation.SaveCopyAs](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.savecopyas), resolves the selected slide through its package relationship, and compares exact persisted DrawingML geometry. This retains group transforms, flips, preset/custom geometry, guide formulas, path commands, curve control points and coordinates. Canonical comparison ignores namespace-prefix/attribute-order spelling, not values or command order. The original source is never saved over; the temporary copy is deleted in finally. No extra PowerPoint presentation is opened. Editor-only node editing-mode state is not inferred from the persisted path.

All existing bounds/text/native formatting and approved raster checks remain. Geometry read errors fail rather than produce a shared unknown sentinel or skip a path. Package parsing prohibits DTDs/external entity resolution, external slide relationships and escaping package paths. Relationship-based lookup is tested with a slide whose file number differs from its presentation position.

The existing11 raster assertions plus15 persisted geometry/package assertions pass (26total). New cases detect changed curve points/commands/guides/group transforms/flips; verify prefix/attribute canonicalization; and reject unknown containers, DTDs, invalid slide indices and external relationships. The parser also read all17 slides from the prior native round-trip evidence deck (artifacts/fidelity-real-package-replay.log). That is saved-package parsing, not a new SaveCopyAs or native Office execution. The new SaveCopyAs path remains a desktop acceptance gate.

The same stable TemplateLibraryTests.exe was rebuilt using a fresh read-only scanner report and the existing reviewed-rebuild process. No add-in reinstall is required for this test-only fix. Re-run Test-TemplateLibrary.ps1 -InteractiveOffice with personal presentations closed and antivirus active.

Final repair suite: artifacts/validation-20261028-111930-9b997d2b/results.json. Source SHA256: 80CDFAA8B1C132FAC175F04DB6BBA82CA6BB3D327D90107EE18A470D13F0D9F7. **20 PASS / 1 FAIL**, only the unchanged x64 scene provenance hold. Template-library local validation and all three architecture compile checks pass. The stable template runner executed its local catalog checks and26 contract assertions; new native SaveCopyAs/template validation remains unexecuted here. Source matches the suite snapshot; git diff --check passes.

## Placeholder inheritance correction — 28 October 2026

The next owner desktop run completed all16 sources and the full reopen/edit cycle. Raster comparisons and native bounds/text geometry passed everywhere. Only original0 (Cubes slide1) and original8 (Pyramids slide6) reported a formatting fingerprint mismatch, on both insertion and reopen.

Exact saved-snapshot comparison found one difference per failing slide: source persisted geometry contained an explicit title transform; destination geometry omitted that local transform and inherited the same transform from its cloned master. All other recorded native-property fields match across the32 insertion/reopen comparisons (`artifacts/placeholder-property-replay.csv`). The title master on both failing slides stores off=(407988,258578), ext=(11412000,387798) EMUs, exactly matching the source's explicit transform. Pixels also match exactly on those two slides. This identifies an inheritance-handling defect in the checker, not evidence of lost visible formatting.

[Microsoft's PresentationML primer](https://download.microsoft.com/download/e/1/4/e14fb96f-83b8-4a2a-84db-7fa8acbe061a/Office%20Open%20XML%20Part%203%20-%20Primer.pdf) describes slide/layout/master inheritance and local overrides. The checker now resolves an absent placeholder transform through its slide-layout relationship, matching the layout placeholder by index, then the master placeholder by type if needed. Explicit slide transforms win over layout/master transforms. It compares the resolved values exactly; no geometry tolerance, default box, skipped placeholder or source-expected value is substituted. Missing/ambiguous associations fail with a diagnostic. The current resolver handles an absent whole transform; it does not claim to implement every partial-property inheritance rule in PresentationML.

Ten additional inheritance assertions pass: explicit-versus-inherited equivalence; layout/slide precedence; rejection of changed inherited coordinates; wrong/ambiguous mapping; unresolved geometry; and actual package relationship resolution. The contract total is36 (11raster +15persisted geometry/package +10inheritance). Existing curve/path/guide/flip and native formatting checks remain exact. Fingerprint schemas advance to drawingml-geometry-v2 and native-template-format-v3.

The saved final evidence presentation contains a deliberate later native text edit, so raw final-package geometry is not treated as a replay of the earlier insertion snapshots. No fresh native acceptance is claimed from offline analysis. A desktop rerun remains required. This is test-only; deployment1.0.0.30/elements2 need not be reinstalled. The stable TemplateLibraryTests.exe was rebuilt through the reviewed process, with scanner evidence retained and no security policy changes.

Final inheritance repair suite: artifacts/validation-20261028-114608-374d46b9/results.json; source SHA256 221739FB04E84F42CEE80D8D7169BDF03F6EBCF177D8F2886B6B781C4AAF09C6. **20 PASS / 1 FAIL** (existing x64 scene provenance hold). All36 fidelity assertions, the local template catalog and architecture compiles pass. Source matches the suite snapshot; whitespace validation passes. Updated native template execution remains pending.

## Native acceptance confirmed — 28 October 2026, 11:51 +04:00

The owner desktop run passed all16 original Cubes/Pyramids templates on Office16.0 x64. Saved launch evidence confirms InteractiveOffice=true, exit0 and PASS, with the current runner hash matching its compiled receipt. All32 CSV rows pass exact effective geometry and exact native formatting: 16 insertion +16 reopen. Raster results are28 exact and4 within the approved antialias bounds. The runner also passed native text edit persistence, native builds/transitions/provenance and unchanged source hashes before returning success.

Scope accepted: these16 source templates and these exercised native scenarios on this desktop. Full497-source coverage, other Office builds/products, remaining element/demo regressions, the x64 scene provenance hold and full feature parity remain open. No new execution is attributed to the AppContainer agent; this is the owner's desktop run, corroborated by saved artifacts.

Preserved acceptance evidence: artifacts/template-fidelity-accepted-20261028-115130/acceptance.json, original-fidelity.csv and launch.json. The acceptance manifest records source fingerprint, runner build key/hash and evidence deck hash. Source SHA256 remains221739FB04E84F42CEE80D8D7169BDF03F6EBCF177D8F2886B6B781C4AAF09C6. This confirmation supersedes earlier pending-native status for this specific16-template scope.
