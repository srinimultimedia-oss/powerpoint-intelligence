# Native element controls — 28 October 2026

Status: source integration and local contracts implemented; native Office acceptance **pending**. This is an independent implementation informed by the public manuals and read-only installed UI observation. It is not full think-cell parity and does not import vendor binaries or bundled artwork.

Build identity: `PPI-DEMO-RIBBON-20261028-v17-elements1`.

## What changed

The previous Elements commands inserted fixed sample text, four static chevrons, or Unicode indicator glyphs. There was no managed universal connector. Those paths now use typed models and native editable objects. Existing 90 command identities remain intact; Element Controls adds 29 commands. Tools exposes the independently implemented local utilities and existing data-link/style/export/review entry points. All 194 icon-bearing ribbon controls render through the original artwork callback.

| Element | Implemented source behavior | Limits / remaining acceptance |
|---|---|---|
| Text Box | Empty editable native transparent AutoShape; centered paragraph/middle vertical defaults; edge locks; content fit grows/shrinks unlocked edges without shrinking fonts; pt/mm/cm/in/percentage/fractional slide dimensions; directional duplicate; in-place kind conversion; native formatting controls. | Native AutoShapes only; rotated/flipped/grouped and arbitrary native text-box conversion rejected. Selection-debounced fitting is not a complete multi-object dynamic layout solver. Native resize overlays, canvas navigation and relation-breaking gestures are absent. |
| Right pentagon / chevron | Separate insert commands, native geometry adjustment, directional duplicate (right of a pentagon becomes chevron), shared text/format controls. | One native adjustment factor; independent tip/notch manipulation is absent. No custom on-canvas handles. |
| Rounded rectangle | Native editable rounded rectangle, adjustment, text-fit locks, conversion and native formatting. | Native Office adjustment geometry, not a claim of vendor radius behavior. |
| Process grid | Empty 2x3 grid of native editable header/body cells; add/hide rows/columns, reorder columns, alternating fills and whole-grid resize. Typed shared grid metadata and stable shape identities. | 20 rows/12 columns; bounds/overflow preflight. Hidden removed cells are retained as inactive history. Copied grids and individually moved cells fail closed. No automatic full-grid content reflow. Individual cell geometry conversions are disabled in the ribbon to protect grid invariants. |
| Universal connector | Native straight/elbow/curve connectors with shape-site, measured native chart-point or free-coordinate endpoints; line colors/weight/dash/start/end arrows; repeat-insert form. Reflow adopts actual native detach/reattach operations. All-connector preflight and rollback. | Top-level native targets; chart-point geometry must be available; rotated/flipped charts and rotated managed connectors rejected. No general obstacle solver or ribbon double-click insertion lock. Legacy untracked v1 chart points become free on first reconciliation to avoid snapping an unknown manual edit; explicitly reselect a chart point to restore tracking. |
| Harvey ball | Editable native vector group; 1–100 segments, fill fraction/count, colors, size, rotation, increment/decrement and indicator conversion. | Values clamp at endpoints. For more than eight segments, digit k uses nearest k/8 fraction, midpoint away from zero. This is our explicit policy, not a verified vendor rounding contract. |
| Checkbox | Empty/check/cross states drawn as native vectors, colors/size/rotation, optional custom text symbol, conversion and cycle. | Custom symbol uses a native text shape/font; it is not vector outline conversion. Manual per-child formatting survives ordinary save/reopen but Configure regenerates child formatting from canonical settings. |
| Indicator interaction | Double-click event handles only our single selected indicator group. Focused keyboard panel offers +/−, digits, Space/V/X and Enter; Escape closes. | No global key hook or promise of canvas-wide shortcuts. No source-to-indicator Excel link, update/revert or shared datasheet yet. |

Text element settings use bounded strict XML in shape CustomerData (`urn:powerpoint-intelligence:text-element:v1`). Grid metadata uses presentation-level `urn:powerpoint-intelligence:process-grid:v1`. Indicators use `urn:powerpoint-intelligence:indicator:v1`. Connector namespace remains `urn:powerpoint-intelligence:connector:v1`, with explicit payload version 2 and read support for version 1. DTDs, unknown schema fields, nonfinite values, invalid dimensions and oversized records are rejected. Corrupt/unowned objects are preserved.

Text fitting skips hidden grid history and process cells. It compares bounded snapshots of current text/bounds/margins before creating temporary measurement probes, and avoids no-op geometry writes. No workbook contents or slide text are written to diagnostic logs.

## Tools inventory and local equivalents

The installed reference Tools menu was observed on 28 October 2026 in Presentation1/window1708252. Only menu inspection was performed; no vendor command was executed. See [reference study](elements-reference-study.md) for the exact grouped inventory and primary public documentation links.

Implemented local Tools menu:

- Searchable Unicode symbol catalog, inserting at selected native text or into a new textbox.
- Reviewed decimal-symbol substitution, changing individual characters to retain mixed formatting. Scope excludes tables/chart labels and explicitly warns that digit-separated dates/thousands separators also match. Legacy decimal command now uses the same character-preserving engine without the new review dialog.
- Selected-object PNG picture copy, retaining the original editable object. This is explicitly raster output, not a new editable chart or a vendor image container. Unrotated objects only; maximum 4096 pixels per side; temporary export is removed.
- Chart-specific workbook manager, opening actual embedded chart data. It is not a unified shared slide workbook.
- Existing mapped Data Link Manager, selected-slide PPTX export, selected-object original JSON style, read-only presentation review, and our own element help.

Still missing: comprehensive reviewed cleanup/sanitization, shared slide workbook/cell mappings, a persistent default for new-chart data tables, full presentation-level current-style management, send-slide packaging options, external indicator data links, and a public-format-compatible Mekko Graphics object importer. Vendor licensing, trial invitations and vendor add-in lifecycle administration are not duplicated or altered. No email is sent and no organization-defined programs/scripts are executed.

## Tests and evidence

The existing stable `DemoContractTests.exe` and `DemoOfficeTests.exe` are reused. No new standalone element test EXE was created. The contracts include:

- 192 text-shape assertions (units/fractions, fitting, all edge-lock combinations, strict codecs).
- 532 process-grid planning/structural/history/schema assertions.
- 91 connector state/endpoint/schema/reconciliation assertions.
- 25 indicator model/shortcut/schema assertions.
- 24 productivity search/substitution/bounds assertions.

Total: **864 new element/productivity pure assertions passed**. These are model and contract checks, not native PowerPoint acceptance. Ribbon tests verify all new callbacks, every Element Controls binding and local tool ID, icon availability/transparency/cache, all 90 legacy bindings and 29 existing workflow icons.

All three .NET Framework 4.8 architecture compile targets pass. Fresh forced root compilations separately verify the actual intermediate DLL contains the new marker (`artifacts/elements-force-*.log`, `artifacts/elements-payload-verification.log`). AnyCPU output is under `obj/Release`; x86/x64 are under their platform directories. Compilation is not execution of native Office scenarios.

`Test-DemoFeatures.ps1` compiles native fixtures for text shapes, process grids, indicators, connectors and productivity tools into the existing Office runner. They cover direct editing, metadata, rejected mutations, duplication, mixed formatting and native save/reopen. `-ElementsOnly` runs these five suites; `-AgendaOnly` remains available. Installed-ribbon validation now reads the current marker dynamically and checks new element/tool bindings, rather than accepting only the old agenda identity.

The local full-suite result and final source snapshot are recorded below after the final run. The x64 scene runner's existing provenance/quarantine hold is retained; an antivirus detection is not inferred solely from a provenance failure. No antivirus setting, exclusion, security permission or approval receipt was modified.

## Required desktop acceptance

This agent session's previously confirmed AppContainer token cannot activate PowerPoint COM or create the signing certificate (0x80070005); the owner's normal desktop token previously passed activation. A UI menu observation does not remove that COM boundary or prove new source is installed. No new native test is marked passed here.

Save and close personal presentations, keep antivirus active, then run from normal Windows PowerShell:

```powershell
.\scripts\Install-Local.ps1
.\scripts\Test-DemoFeatures.ps1 -InteractiveOffice -ElementsOnly
```

Complete the visible publisher prompt. Verify the installed Intelligence > Elements / Element Controls / Tools menus, the six element types, Edit settings, double-click indicator behavior and the focused shortcut panel. Edit native fields, save and reopen in the same presentation. Review generated `features.csv` and evidence PPTX files. After focused suites pass, run `Test-All.ps1 -InteractiveOffice` for regression coverage. Office 2016, Office 2019 and Microsoft 365 runtime matrices remain separate unverified acceptance gates.

Full parity and final acceptance remain **open** until the listed source gaps and native evidence gates are resolved.

## Final executed local result

Final suite: artifacts/validation-20261028-102634-5b993a6c/results.json. Source SHA256: 29AF3E3131FB3F93117EDD05034FAA5B93AF5E6B10CE251C7793CDDF93206D41. **20 PASS / 1 FAIL**; the only failure is scene-contracts-x64, whose existing runner provenance no longer matches its build inputs. The guard retained the existing executable and required a reviewed rebuild. Its previously recorded quarantine association is not treated as proof of a new scanner event. Renderer core remains **41/41 PASS**. The new element/productivity pure assertions remain **864 PASS**. All AnyCPU/x86/x64 architecture compilation checks passed.

Environment: Windows; 64-bit OS=True, 64-bit validation shell=True. Production target .NET Framework 4.8; Office15 PIAs. Target PowerPoint architecture/native actions were not executed by this agent. Fresh root payload marker/hash checks are in artifacts/elements-payload-verification.log. Final source fingerprint was checked against the full-suite snapshot after all code/script changes. Documentation does not affect this fingerprint. git diff --check passed.

Acceptance: local integration ready for desktop verification; full feature parity and native acceptance remain OPEN. No deployment receipt was forged or updated and no installation is claimed for this source.

## Follow-up desktop evidence and repair

The owner's 28 October desktop run confirmed 16 native indicator assertions and installed-ribbon loading. It also exposed XML update, frame, workbook-close and fixture defects. Current repair marker is `PPI-DEMO-RIBBON-20261028-v17-elements2`; see [Office regression repair](office-regression-elements2.md) for exact failures, changes, final local evidence and outstanding native acceptance. The earlier element-controls source marker/test counts describe the initial increment; the follow-up adds eight pure assertions and retains strict template-pixel failure status.
