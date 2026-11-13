# Template insertion, source-original sessions and layout links

<!-- generated-toc -->
## Contents

- [Purpose and implementation status](#purpose-and-implementation-status)
- [Source locations and internal components](#source-locations-and-internal-components)
- [Inputs, outputs, data structures and dependencies](#inputs-outputs-data-structures-and-dependencies)
- [Internal execution and external integration](#internal-execution-and-external-integration)
- [Configuration, validation and failure behavior](#configuration-validation-and-failure-behavior)
- [Security considerations and cleanup](#security-considerations-and-cleanup)
- [Test references](#test-references)
- [Known limitations and technical debt](#known-limitations-and-technical-debt)
- [Extension points](#extension-points)
- [Classes, interfaces and method navigation](#classes-interfaces-and-method-navigation)
<!-- end-generated-toc -->

**VERIFIED — source implementation. Current runtime acceptance is not established by this design review.**

## Purpose and implementation status

Insert editable designed templates or licensed original slides and expose palette/alignment/native-animation controls.

Implemented local source is distinguished from supported subset/absent external behavior below. See current test evidence in the export testing chapter; do not treat historical logs or method existence as a current runtime pass.

## Source locations and internal components

- `apps/powerpoint-vsto/TemplateLibrary.cs:1` — TemplatePreview, TemplateInserter, AlignmentPicker, PalettePicker, TemplateLibrary, LocalSlideAsset, LocalSlideIndex, LayoutRule, LibraryLayoutLinks, TemplateSources, SourceTemplateSession.

## Inputs, outputs, data structures and dependencies

Inputs: catalog template/fields/palette/alignment/build options or LocalSlideAsset; outputs: editable slide/group, provenance/XML and optional native effects. Dependencies: TemplateCatalog, Office Slides/Shapes/Designs, PIIDX references and System.Drawing.

## Internal execution and external integration

Designed insertion validates fields/options and scales960x540nodes into active presentation. New slide follows active Design; native text/shapes/freeforms get instance/node/field tags and XML. Optional grouped stages use click-triggered Fade/Wipe/Fly-in0.35s; transitions None/Fade/Wipe, click advance only. Object-only mode groups into active slide and rejects staged builds. Palette updates target instance/whole designed slide while preserving text; alignment updates native paragraphs/anchors. SourceTemplateSession opens owned hidden read-only source, rejects destination/already-user-open source, validates fingerprint/native identity and exports PNG. Keep-source insertion requires matching size, uses InsertFromFile, clones Design and retains CustomLayout/background. Layout links encode edge/offset dependency rules and reflow followers.

## Configuration, validation and failure behavior

TemplateLibraryFolder resolves beside DLL or explicitly configured folder; no vendor-install fallback. Keep-source size tolerance.05pt. Original assets require the owner license. Preview temp PNG is removed; Dispose closes only owned hidden source.

## Security considerations and cleanup

Local native calls run in the current Office user process and require supported selection/ownership; a mutex/selection key is not authentication or RBAC. File/metadata content must remain bounded and error text must not expose data or credentials. Explicit owned-resource cleanup and error/rollback paths are part of the cited implementation, not a guarantee that a failed COM server can always be restored.

## Test references

Lexical module-class matches in test source (not instrumentation coverage; no test is claimed executed here):

- `tests/library/GalleryTests.cs:17`.
- `tests/library/LegacyLibraryCacheFixture.cs:11`.
- `tests/library/LibraryBrowserTests.cs:57`.
- `tests/library/PiidxGalleryTests.cs:35`.
- `tests/library/TemplateLibraryTests.cs:140`.

## Known limitations and technical debt

Palette/alignment features target generated tagged templates, not arbitrary originals. Source changes/layout mismatch stop insertion; rollback deletes only owned inserted slide/design/group and reports incomplete cleanup. Grouped stage ownership/fidelity requires native testing across assets; historical16slide fidelity sample is not universal certification.

## Extension points

Add modes only with explicit source/destination ownership, native editability and design/layout/hash/save-reopen regression tests.

## Classes, interfaces and method navigation

Method names below are source declarations; private/internal members are not public API promises. The source hash/review manifest gives exact reviewed bytes.

| Declaration | Kind | Source |
|---|---|---|
| `TemplatePreview` | class | `apps/powerpoint-vsto/TemplateLibrary.cs:19` |
| `ColorOf` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:21` |
| `Render` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:22` |
| `TemplateInserter` | class | `apps/powerpoint-vsto/TemplateLibrary.cs:60` |
| `Insert` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:63` |
| `ApplyTheme` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:122` |
| `ThemeColor` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:127` |
| `Paragraph` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:131` |
| `Anchor` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:132` |
| `ChangeAlignment` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:133` |
| `CollectAll` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:152` |
| `ChangePalette` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:153` |
| `SelectedInstance` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:195` |
| `Collect` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:203` |
| `Rgb` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:211` |
| `AlignmentPicker` | class | `apps/powerpoint-vsto/TemplateLibrary.cs:213` |
| `Choose` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:215` |
| `PalettePicker` | class | `apps/powerpoint-vsto/TemplateLibrary.cs:226` |
| `Configure` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:228` |
| `Choose` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:234` |
| `TemplateLibrary` | class | `apps/powerpoint-vsto/TemplateLibrary.cs:245` |
| `SaveDraft` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:292` |
| `Filter` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:293` |
| `ShowPage` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:302` |
| `SelectTemplate` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:311` |
| `ShowTemplate` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:313` |
| `Values` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:318` |
| `ReplacePreview` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:319` |
| `UpdatePreview` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:320` |
| `Insert` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:321` |
| `Dispose` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:327` |
| `LocalSlideAsset` | class | `apps/powerpoint-vsto/TemplateLibrary.cs:334` |
| `ToString` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:341` |
| `LocalSlideIndex` | class | `apps/powerpoint-vsto/TemplateLibrary.cs:343` |
| `Read` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:345` |
| `SlideText` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:351` |
| `Value` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:357` |
| `Resolve` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:358` |
| `Relation` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:363` |
| `IsWithin` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:367` |
| `Discover` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:368` |
| `ReadFile` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:374` |
| `Filter` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:390` |
| `LayoutRule` | class | `apps/powerpoint-vsto/TemplateLibrary.cs:400` |
| `Encode` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:404` |
| `Decode` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:405` |
| `Validate` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:406` |
| `Horizontal` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:407` |
| `Edge` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:408` |
| `Position` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:409` |
| `LibraryLayoutLinks` | class | `apps/powerpoint-vsto/TemplateLibrary.cs:411` |
| `Rect` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:414` |
| `Set` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:415` |
| `Reflow` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:428` |
| `Controls` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:435` |
| `TemplateSources` | class | `apps/powerpoint-vsto/TemplateLibrary.cs:444` |
| `ReadFolderSetting` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:446` |
| `ResolveFolder` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:450` |
| `RequireSameSize` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:459` |
| `SourceTemplateSession` | class | `apps/powerpoint-vsto/TemplateLibrary.cs:462` |
| `Source` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:476` |
| `Preview` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:484` |
| `InsertOriginal` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:489` |
| `Dispose` | method | `apps/powerpoint-vsto/TemplateLibrary.cs:499` |
