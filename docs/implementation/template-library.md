# Editable template library

## Scope and design

On 23 October 2026 the owner requested a Library element with more than 500 distinct slide templates and approved optional native PowerPoint builds/transitions. This increment implements **520 original procedural layouts in 26 families**, each with 20 structural configurations. The catalog uses different compositions, stage counts, relationships and table dimensions. Themes, color swaps and text substitutions are not counted as new templates. A geometry fingerprint excludes text and colors and currently distinguishes every entry. This demonstrates structural uniqueness, not 520 independently implemented business applications or 520 individually hand-crafted designs. Templates share 26 deterministic layout builders. The supplied screenshots were visual references; no proprietary template files, artwork or binaries were imported.

Families: agendas, timelines, processes, cycles, pyramids, funnels, organization, decisions, matrices, SWOT, comparisons, roadmaps, swimlanes, customer journeys, portfolios, scorecards, teams, quotes/evidence, opening slides, section navigation, tables, architecture, ecosystems, cubes, overlapping sets and experiment plans. Each named entry has an ID, category, purpose and versioned editable fields.

## User flow

1. Open a presentation and select a slide in Normal view.
2. Select **Intelligence → Library** (ribbon marker `PPI-DEMO-RIBBON-20261023-v7`).
3. Filter by family or search by name/purpose. The library shows 40 previews per page, with Previous/Next controls.
4. Select a template. Edit its title, subtitle and named content fields in the field grid; use Update preview to see the edits.
5. Optionally enable click-to-reveal builds and choose Fade, Wipe or Fly in. Optionally choose a None/Fade/Wipe transition. Defaults are static with no transition.
6. Insert into current presentation. A new slide is placed immediately after the selected slide. No additional PowerPoint presentation is opened.
7. Edit native text, shapes, colors, points, dimensions and groups in PowerPoint. Edit effects in Animation Pane and Transitions. Save the presentation normally.

Default labels are clearly marked placeholders, not claimed business data. This library does not create sample-data charts and does not change the linked-source chart workflow. These are slide/diagram scaffolds, not live calculations or linked native charts. Cube faces are editable vector polygons, not a 3D modeling system. Matrix/table cells are editable native shape/text fields, not native PowerPoint table objects. Organization/decision connectors are editable lines, not automatically rerouting SmartArt. Calendar scheduling, geographic maps and arbitrary diagram recalculation are not implemented by these templates.

## Native representation and safeguards

`TemplateCatalog.cs` defines typed internal layout nodes on a 960×540 point canvas. The GDI+ previews and COM adapter use the same node coordinates, text and colors. Preview fonts and rounded-corner treatment may differ slightly from Office. Native text boxes, shapes, lines and freeforms remain editable. No raster screenshot is inserted into the presentation. The slide is fit proportionally to the current presentation's dimensions and centered, including 4:3 decks.

`TemplateLibrary.cs` implements the browser and insertion adapter. The adapter validates the catalog ID, supported animation/transition choice, field names and field length (2,000 characters maximum) before slide mutation. It creates only one owned slide; insertion failure deletes that slide rather than leaving partial content. It preserves the existing presentation and uses a PowerPoint undo entry. Text can auto-fit in PowerPoint, so extreme user-entered content may shrink and requires visual review.

Each stage can be grouped and animated with a native 0.35-second effect triggered on click. Group contents remain editable in PowerPoint. There are no timers, external media, VBA, Morph-only features or automatic slide advance. These conservative API choices target Office 2016/2019/Microsoft 365, but product/build-specific animation behavior is not certified until executed on those versions.

Slides and objects carry template/field/stage tags. A v1 Custom XML record (`urn:powerpoint-intelligence:template:v1`) attached to the owned title object stores the template identity, insertion-time field values and chosen build/transition options. That record is an insertion snapshot; subsequent manual edits remain native PowerPoint content and do not automatically update the snapshot. There is no currently implemented reapply/regenerate command that could overwrite those edits. Contents are not written to diagnostic logs or sent to external services.

## Checks and reproduction

- `scripts/Test-TemplateLibrary.ps1`: compiles a genuine x64 test executable, builds all 520 layouts twice, rejects duplicate geometry signatures, checks bounds/field IDs, renders 520 previews plus 26 family contact sheets, exercises search/category filtering and constructs the browser controls without Office activation.
- `scripts/Test-TemplateLibrary.ps1 -InteractiveOffice`: additionally creates an owned test deck, inserts all 520 templates into that same presentation, mixes static/animated options and native effect styles, edits title and content fields including grouped fields, saves/reopens, and verifies retained edits, fields and animation presence. It refuses to proceed when the Office instance has user presentations.
- `scripts/Test-All.ps1 -InteractiveOffice`: includes library local and Office checks alongside existing chart/domain/compatibility/ribbon tests.
- `scripts/Test-InstalledRibbon.ps1`: now requires the v7 build and Library command. Compile/structural ribbon checks do not establish actual runtime visibility.

The all-template Office test creates 521 slides (one initial blank plus the 520 templates) and can take time. It is intended for an interactive Windows desktop with user documents closed. It is not a server-side Office workload.

## Visual review and limitations

All 26 family contact sheets were inspected at consistent scale. Review led to corrections for self-crossing pyramid faces, label/face collisions, team-card text overlap, branching connector z-order and set-label occlusion. Automated geometry uniqueness/bounds do not establish semantic correctness for every business situation. A full visual review in PowerPoint, all-template save/reopen and animation playback remain explicit acceptance gates. An off-screen WinForms window-render attempt stalled in the restricted agent session and was stopped; no blank browser capture is presented as evidence. GDI+ slide previews are design evidence only.

Current acceptance: source implementation and local validation only; not yet accepted as a fully validated installed Office library.

## Validation evidence — 23 October 2026

Final full run: `artifacts/validation-20261023-100734-d041011d`.
Source SHA256: `E62EF283165E684801AE3C29675998DDCF4C260EEEC9DFEF5AD685FFE88F0690` (confirmed equal to current executable/test sources after validation).

**21/21 local checks pass**, including template-library (genuine Windows x64 executable), all 520 deterministic layouts and unique geometry fingerprints, 520 design previews, 26 family sheets, existing 68 data-source checks, native-renderer core, ribbon bindings and AnyCPU/x86/x64 compile targets. The UI adds a Library button using the existing original library-browser icon. The template check adds one local check to the previous 20-check baseline.

**Five Office checks fail**: template-library-office, installed-ribbon, demo-feature-runtime, native-shapes-office and office-1. The new template test fails while retrieving the PowerPoint COM class factory with `0x80070005`, before inserting any templates. The underlying agent token limitation and desktop comparison are documented in the activation diagnostics from earlier work. No security controls were weakened. Native animation playback and all-template save/reopen are NOT marked passed. The current library source has not been installed; the existing package predates marker v7.

Review artifacts outside the immutable run: `artifacts/template-library-x64/catalog.tsv` lists all 520 entries, field/node counts and structure hashes; `catalog.html` is an offline searchable self-contained preview gallery; `featured-00.png` through `featured-25.png` are 960×540 family examples. The gallery is a design-preview artifact, not a replacement for the native add-in browser or Office validation.

To validate on the approved normal desktop, save/close PowerPoint, run `scripts/Install-Local.ps1` (without SkipBuild because executable source changed), complete the normal publisher prompt, then run `scripts/Test-All.ps1 -InteractiveOffice`. Confirm the Library button visually; test search, pagination, field edits, static insertion, all three build styles, both transitions, native point/text edits and save/reopen. Also verify proportional insertion in a 4:3 presentation and cancellation leaves the deck untouched. The automated template round-trip covers all 520 designs, native field edits including grouped content, static/animated variants and animation persistence; it does not by itself prove visual playback quality or certify all Office product builds.

## Library palettes and recoloring — 23 October 2026

The supplied `powerpoint_subtle_color_palettes.md` defines 20 named palettes with ten colors each. All names and all 200 HEX values were compared against the imported definitions; they match exactly. The portable source table is retained in `docs/implementation/library-palettes.csv`; runtime definitions are in `TemplateCatalog.cs`. Document prose was treated as reference material; the implementation follows the owner's request.

**Executive Slate (01) is now the default** for library thumbnails, large previews, slide insertion and preview exports. Its colors are `#34495E #536779 #708494 #90A1AE #B1BEC7 #D1D9DF #E9EDF0 #718B91 #A6B5B0 #F5F7F8`. Colors are assigned to semantic surfaces, headings, accents and connectors. All 20 choices are available in the library's Color palette selector, with ten-color swatches in its dropdown. Changing the palette refreshes thumbnails/preview and retains edited field values before insertion. Themes are not counted as additional templates.

After insertion, select the library slide and choose **Intelligence → Template Colors**. The ribbon KeyTip sequence is **Alt → PI → CP** (press the letters in order). This opens the palette selector; Cancel performs no mutation. Apply changes tracked library text/fill/line colors and the slide background, including objects inside animated groups. It retains text, geometry, stage grouping and animation effects and updates the palette identity. It targets the current slide only; manually added objects without template tags are left unchanged. Per-slide palette selection is stored in tags and template metadata. Library slides from the preceding v7 build are supported through retained template IDs and `ppi_template_N` node names. Unknown/untracked slides are rejected instead of recolored blindly.

The recolor command records previous RGB colors, including per-character rich-text colors, and attempts reverse-order restoration if COM rejects a change. It reports incomplete rollback and advises PowerPoint Undo if restoration itself fails. This is not a claim of executed COM rollback acceptance. Palette changes intentionally replace the managed template colors, including manually changed colors on those objects; they do not replace native text or redraw shapes. Manually replacing/deleting field tags or renaming legacy untagged nodes can prevent those objects from being recognized.

Text uses the palette's dark/light tone when it reaches 4.5:1 calculated contrast against its intended generated surface; black/white are used when neither supplied tone reaches that value. Local tests cover all 20 palettes' surface/foreground calculations. This is not a WCAG certification: manual background changes, transparent/overlapping geometry, custom content and Office rendering still need visual review.

Tests now cover all 10,400 template/palette combinations for unchanged geometry and field text, default selection, 20×10 definitions and unknown-palette rejection. Twenty 960×540 example renders are exported as `palette-01.png` through `palette-20.png`, with an offline comparison gallery at `artifacts/template-library-x64/palettes.html`. The Office suite now recolors all 520 inserted templates, checks object/effect counts, position and text, and checks palette IDs after save/reopen. Runtime execution remains an explicit acceptance gate.

Reproduction: rebuild/install current source (marker v8), choose Library → template → Sage & Stone, edit a field and change palette again before insertion to verify the edit remains. Insert with animations, edit native text in PowerPoint, then use Template Colors / Alt PI CP to apply Coastal Blue. Check text, shape positions and effects remain; save/reopen and repeat. Verify cancel and a non-library slide. Run `scripts/Test-All.ps1 -InteractiveOffice` in the approved normal desktop environment.

Palette validation run: `artifacts/validation-20261023-145142-13784d89`; source SHA256 `C9011326A93CE20F050127115B6114490E3ED3C9C936B6F1460E05264570B1BA`. All 21 local checks pass, including the palette combinations and AnyCPU/x86/x64 builds. The five Office integration checks fail at activation with 0x80070005. Post-insertion palette changes, rollback, native UI selection and save/reopen therefore remain unverified in Office. This update requires rebuilding/installing the v8 source; the previous deployment does not contain it.


## Text alignment repair — 23 October 2026

The supplied `Presentation2.pptx` contains 512 slides. Its saved slide XML confirmed that many text frames had no explicit paragraph alignment or vertical anchor, inheriting left/top defaults. The insertion adapter previously specified font and wrapping but not alignment. Some labels also used short text frames near the upper edge of their parent shapes, so centering their text alone would not center the label within the shape.

New library insertion and previews default to horizontal **Center** and vertical **Middle** for all template text. A single label within a shape is fitted to the shape's padded interior; chevrons, diamonds and polygons use safer horizontal insets. Multi-field cards keep their separate text-frame locations. The library offers Left/Center/Right and Top/Middle/Bottom before insertion. Center means centered paragraph alignment, not fully justified word spacing.

For inserted slides, **Intelligence → Template Alignment** (KeyTips **Alt → PI → TA**) changes the alignment of selected tracked library objects/groups. With no shape selection it applies to all tracked template text on the current slide. It preserves text, geometry and animation objects. Cancel does not mutate the deck. A failed COM operation attempts to restore previous paragraph/vertical alignment and reports incomplete rollback. This command aligns existing text frames; it does not undo manual geometry edits or automatically refit legacy frames. New insertion uses the corrected frame geometry.

A repaired, separate copy is at `artifacts/Presentation2-centered.pptx`. The original file was not overwritten. `scripts/Repair-TemplateAlignment.py` reads the saved input package and writes a new offline package; it never edits an open presentation. It centers 11,398 text frames and fits 3,325 single labels to enclosing shape bounds. It verifies all 512 slides, every original text value, unchanged timing/transition XML, identical unrelated ZIP entries and an unchanged input SHA256. Evidence: `artifacts/Presentation2-centered.validation.json`. These are structural checks; they do not establish actual PowerPoint rendering. The original saved package's thumbnail is preserved and may refresh only when PowerPoint saves the repaired copy.

Tests cover single-label fitting, multi-field position preservation, the nine supported alignment combinations, invalid values, preview rendering, the new ribbon binding and native default/apply alignment assertions in the Office suite. All 21 local checks pass in `artifacts/validation-20261023-150430-948ffdc2`, source SHA256 `8F2014D62BDD119876A1FF9368557AF28EF5896E73CF7AF2F3BE1E4EF0911218`. The five Office checks still fail on PowerPoint activation with 0x80070005 before native alignment assertions. Actual visual playback, UI controls and save/reopen of the updated insertion path remain unverified. The original deck cannot be fully rendered by PowerPoint in this restricted execution context, and no LibreOffice renderer was found; design previews are not presented as native PowerPoint renders.

Apply the source update with `scripts/Install-Local.ps1` (fresh build, marker v9), then test the repaired copy and new templates in normal desktop PowerPoint. Verify labels at normal zoom, change one selected field to Left/Top, change the whole slide to Center/Middle, cancel once, save and reopen. Native Office controls remain available for independent alignment changes.

## Desktop regression diagnosis — 23 October 2026, 15:09 run

The owner's normal desktop run `artifacts/validation-20261023-150939-ae651824` provides new runtime evidence: **installed-ribbon PASS** and **native-shapes-office PASS**, including native bar edits after save/reopen, canonical data retention, duplicate handling and persistence rollback. Desktop PowerPoint activation succeeds. Earlier 0x80070005 results apply to the restricted agent context; they do not describe this desktop run.

The template check failed at `CustomXMLPart.LoadXML` during recoloring: Office reports “This custom XML part has already been loaded.” The rollback repeated the same unsupported reload and reported incomplete rollback. The implementation now edits the existing document element's palette attribute through `CustomXMLNode.NodeValue`, creates it through AppendChildNode only when absent, and rolls back the attribute value/deletion. It does not reload an already-loaded part. Initial creation still loads XML into a fresh part, which is a separate supported path. Local compilation passes; the updated native operation has not yet executed successfully in the restricted agent context.

Demo runtime first reported point-count mismatches for bar, line and profile. Later it disconnected during scatter; 81 aggregate failures included many commands attempted against the dead server. The captured crash records identify POWERPNT.EXE / chart.dll / c0000005 and Excel c0000409 records. These establish native process crashes but do not by themselves establish their triggering cause. Source changes now wait up to eight seconds for committed chart series counts/point counts/nonmissing numeric values before returning; call-rejected errors are retried within that bound, RPC disconnections are not. Pie/doughnut/bubble expect one series and profile preserves row-oriented dimensions. Scatter/bubble write numeric X cells and configure typed series before embedded workbook close, avoiding transient post-close series rewrites. These are proposed runtime fixes requiring desktop confirmation, not a verified resolution of chart.dll crashes.

The demo runner stops executing subsequent commands after an RPC disconnection and writes BLOCKED rows instead of dozens of derivative feature failures. S0 in the owner's run never reached fixture insertion: all three runs refused to proceed because presentations were already open. The guard is retained; the suite now avoids repeating that same precondition failure. No user presentation is closed automatically and no security controls are changed.

Post-change run: `artifacts/validation-20261023-151533-42603dab`, source SHA256 `D68345A8B4E32F0CA87444BDDB695B46690AC540CB2329C66FD6FB6BB9D35031`. All **21 local checks pass**, including AnyCPU/x86/x64 compilation. The agent's five Office checks still fail at activation with 0x80070005, so this run cannot supersede the owner's desktop integration evidence or verify the proposed fixes. Rebuild/install current source in normal desktop PowerShell, with user presentations saved/closed, and rerun the interactive suite. Template, demo/chart stability and S0 acceptance remain pending.

## Library controls increment — 26 October 2026

The Library now offers search/counts, thumbnail zoom/sort, retained drafts, matching-field preview highlights, object-only insertion, destination-theme adaptation and a **My PowerPoint files** browser. The separate **Layout Links** ribbon command provides opt-in native alignment/gap/edge constraints and text-fit controls. See [evidence, behavior and limits](thinkcell-observable-study.md). The Remember checkbox stores an opt-in searchable local index; unchecking removes it. User-file listings expose text/metadata rather than native thumbnail images. Procedural template instances support palette/alignment targeting; object-only insertion rejects staged builds and retains the current slide background/transition.

Desktop PowerPoint activation is verified by the owner's 26 October comparison. Installed v11 interactions, native layout-link persistence, object-only round trips and the added integration assertions remain acceptance gates. Compile/pure-test results do not close those gates.

## Original source collection (v12)

The Library command now defaults to the installed external `.potx` collection when available: 62 files / 497 slides on the observed machine. This replaces procedural layouts as the default for the owner's unchanged-design requirement. Native previews and source-design/custom-layout import use approved hidden read-only access, with exact slide-size checks. No source assets are bundled into the product repository. The original files must remain accessible; think-cell-specific editing functionality is not independently recreated. Pixel-level import/save-reopen verification is pending the focused original-source Office scenario. Procedural layouts remain a separately labelled fallback.

## Thumbnail gallery (v15)

Original-file browsing now shows paged native slide thumbnails instead of the text list. Select a card for the large preview; double-click or press Enter to insert. Thumbnail size and details visibility are adjustable. Previews load on the Office UI thread for the current 24-slide page, with a 96-image memory cache. The gallery screenshot in the observable-study artifacts uses seven captured native renders for local UI verification; installed Office interaction is a separate pending gate.
