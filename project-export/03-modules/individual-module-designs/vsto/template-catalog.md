# Designed template layouts, editable fields and palettes

**VERIFIED — source implementation. Current runtime acceptance is not established by this design review.**

## Purpose and implementation status

Generate deterministic native-oriented layout specifications, fields and palettes separately from licensed original slide packages.

Implemented local source is distinguished from supported subset/absent external behavior below. See current test evidence in the export testing chapter; do not treat historical logs or method existence as a current runtime pass.

## Source locations and internal components

- `apps/powerpoint-vsto/TemplateCatalog.cs:1` — LibraryDrafts, TemplateTheme, TemplateAlignment, TemplatePalette, TemplatePalettes, TemplateNode, SlideTemplate, TemplateCatalog, Composer.

## Inputs, outputs, data structures and dependencies

Inputs: template/palette and draft field values. Outputs: SlideTemplate/TemplateNode geometry for preview/native insertion. Dependencies: TemplateLibrary/Inserter and System.Drawing preview; no COM in pure catalog.

## Internal execution and external integration

Catalog enumerates categories/variants and composes text/box/line/polygon nodes on960x540canvas, with unique IDs/field keys and staged builds. Validators check geometry/bounds/keys. Palette mapping remaps semantic colors and derives text contrast against inferred background; default is Executive Slate01. Draft values are per-template form-lifetime state. Alignment validation/preparation distinguishes horizontal and vertical modes.

## Configuration, validation and failure behavior

Twenty supplied palettes; default ID01; contrast target4.5 uses local-node background heuristic. Node validators constrain supported primitives/points; insertion fields<=2000 chars.

## Security considerations and cleanup

Local native calls run in the current Office user process and require supported selection/ownership; a mutex/selection key is not authentication or RBAC. File/metadata content must remain bounded and error text must not expose data or credentials. Explicit owned-resource cleanup and error/rollback paths are part of the cited implementation, not a guarantee that a failed COM server can always be restored.

## Test references

Lexical module-class matches in test source (not instrumentation coverage; no test is claimed executed here):

- `tests/library/LibraryBrowserTests.cs:36`.
- `tests/library/TemplateLibraryTests.cs:30`.

## Known limitations and technical debt

Historical tests enumerate520layouts and10400palette/layout combinations; this leaf does not re-run tests. Distinct IDs/geometry do not prove every layout has a unique business function or source-original fidelity. Drafts are not a durable remote template service.

## Extension points

Add unique IDs/fields and explicit layout geometry, bounds/preview/contrast tests and independent semantic design review. Do not inflate feature counts with palette variations.

## Classes, interfaces and method navigation

Method names below are source declarations; private/internal members are not public API promises. The source hash/review manifest gives exact reviewed bytes.

| Declaration | Kind | Source |
|---|---|---|
| `LibraryDrafts` | class | `apps/powerpoint-vsto/TemplateCatalog.cs:7` |
| `Load` | method | `apps/powerpoint-vsto/TemplateCatalog.cs:10` |
| `Save` | method | `apps/powerpoint-vsto/TemplateCatalog.cs:16` |
| `TemplateTheme` | class | `apps/powerpoint-vsto/TemplateCatalog.cs:23` |
| `Binding` | method | `apps/powerpoint-vsto/TemplateCatalog.cs:26` |
| `TemplateAlignment` | class | `apps/powerpoint-vsto/TemplateCatalog.cs:35` |
| `Validate` | method | `apps/powerpoint-vsto/TemplateCatalog.cs:39` |
| `Prepare` | method | `apps/powerpoint-vsto/TemplateCatalog.cs:41` |
| `TemplatePalette` | class | `apps/powerpoint-vsto/TemplateCatalog.cs:55` |
| `ToString` | method | `apps/powerpoint-vsto/TemplateCatalog.cs:59` |
| `TemplatePalettes` | class | `apps/powerpoint-vsto/TemplateCatalog.cs:61` |
| `Get` | method | `apps/powerpoint-vsto/TemplateCatalog.cs:86` |
| `Map` | method | `apps/powerpoint-vsto/TemplateCatalog.cs:87` |
| `Luminance` | method | `apps/powerpoint-vsto/TemplateCatalog.cs:94` |
| `Contrast` | method | `apps/powerpoint-vsto/TemplateCatalog.cs:96` |
| `TextColor` | method | `apps/powerpoint-vsto/TemplateCatalog.cs:97` |
| `Apply` | method | `apps/powerpoint-vsto/TemplateCatalog.cs:99` |
| `TemplateNode` | class | `apps/powerpoint-vsto/TemplateCatalog.cs:112` |
| `SlideTemplate` | class | `apps/powerpoint-vsto/TemplateCatalog.cs:120` |
| `Build` | method | `apps/powerpoint-vsto/TemplateCatalog.cs:124` |
| `ToString` | method | `apps/powerpoint-vsto/TemplateCatalog.cs:125` |
| `TemplateCatalog` | class | `apps/powerpoint-vsto/TemplateCatalog.cs:129` |
| `Create` | method | `apps/powerpoint-vsto/TemplateCatalog.cs:162` |
| `Search` | method | `apps/powerpoint-vsto/TemplateCatalog.cs:169` |
| `Build` | method | `apps/powerpoint-vsto/TemplateCatalog.cs:175` |
| `Validate` | method | `apps/powerpoint-vsto/TemplateCatalog.cs:278` |
| `Composer` | class | `apps/powerpoint-vsto/TemplateCatalog.cs:284` |
| `Text` | method | `apps/powerpoint-vsto/TemplateCatalog.cs:287` |
| `Shape` | method | `apps/powerpoint-vsto/TemplateCatalog.cs:288` |
| `Line` | method | `apps/powerpoint-vsto/TemplateCatalog.cs:289` |
| `Polygon` | method | `apps/powerpoint-vsto/TemplateCatalog.cs:290` |
