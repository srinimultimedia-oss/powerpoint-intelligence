# Ribbon catalog, dispatch, icons, browser and automation facade

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
- [Per-instance input ports (testability boundary)](#per-instance-input-ports-testability-boundary)
<!-- end-generated-toc -->

**VERIFIED — source implementation. Current runtime acceptance is not established by this design review.**

## Purpose and implementation status

Expose NexWave commands and searchable local tools while keeping descriptive requirements separate from executable commands.

Implemented local source is distinguished from supported subset/absent external behavior below. See current test evidence in the export testing chapter; do not treat historical logs or method existence as a current runtime pass.

## Source locations and internal components

- `apps/powerpoint-vsto/DemoRibbon.cs:1` — DemoRibbon.
- `apps/powerpoint-vsto/DemoTools.cs:1` — GridPlanner, DemoTool, DemoTools, DemoStyle.
- `apps/powerpoint-vsto/RibbonIcons.cs:1` — RibbonIcons, PictureBridge.
- `apps/powerpoint-vsto/DemoWorkspace.cs:1` — DemoWorkspace.
- `apps/powerpoint-vsto/FeatureInventory.cs:1` — FeatureInventoryEntry, FeatureInventory.
- `apps/powerpoint-vsto/DemoInputs.cs:1` — DemoInputs.
- `apps/powerpoint-vsto/DemoApi.cs:1` — IDemoAutomation, DemoAutomation.

## Inputs, outputs, data structures and dependencies

Inputs: command tags, current selection, prompts/files or fixture delegates and S0 JSON. Outputs: native edits, forms, local exports and COM observation. Dependencies: managed element/data/template/format/quality adapters and S0 host.

## Internal execution and external integration

DemoRibbon.Create builds Office2009 customUI XML; IntelligenceRibbon forwards callbacks/tags. DemoTools.BuildCatalog binds IDs/delegates and generated arrange modes; Execute opens a native undo boundary. DemoWorkspace displays tools and requirement statuses. RibbonIcons draws local16/32px bitmap glyphs and bridges IPictureDisp. DemoInputs supplies per-instance fixture/UI ports. IDemoAutomation exposes build/version/ribbon/immutable properties state and bounded S0 JSON insertion, not HTTP.

## Configuration, validation and failure behavior

BuildMarker is deployment freshness data. GetContextPropertiesState reads a volatile UI-published string. InsertChartJson requires the creating UI STA, validates <=1 MiB JSON/model and assigns fresh GUID/revision. No arbitrary AI-to-COM dispatch is present.

## Security considerations and cleanup

Local native calls run in the current Office user process and require supported selection/ownership; a mutex/selection key is not authentication or RBAC. File/metadata content must remain bounded and error text must not expose data or credentials. Explicit owned-resource cleanup and error/rollback paths are part of the cited implementation, not a guarantee that a failed COM server can always be restored.

## Test references

Lexical module-class matches in test source (not instrumentation coverage; no test is claimed executed here):

- `tests/library/LibraryBrowserTests.cs:30`.
- `tests/s0/ContextPropertiesOfficeTests.cs:60`.
- `tests/s0/ContextPropertiesTests.cs:29`.
- `tests/s0/DemoContractTests.cs:76`.
- `tests/s0/DemoOfficeTests.cs:35`.
- `tests/s0/DemoScenario.cs:107`.

## Known limitations and technical debt

FeatureInventory labels are historical descriptive claims, not runtime proof. Simple arrange operations may rely on Undo rather than complete per-object rollback. Linked report adds seven chart slides to the same active deck with one linked source and no sample fallback. Office JavaScript/HTML task pane and SmartArt authoring are absent from the compiled graph.

## Extension points

Add catalog/callback/icon/test entries together. Maintain separate counts for executable routes and requirement inventory; preserve stable internal names while visible brand is NexWave.

## Classes, interfaces and method navigation

Method names below are source declarations; private/internal members are not public API promises. The source hash/review manifest gives exact reviewed bytes.

| Declaration | Kind | Source |
|---|---|---|
| `DemoRibbon` | class | `apps/powerpoint-vsto/DemoRibbon.cs:9` |
| `Create` | method | `apps/powerpoint-vsto/DemoRibbon.cs:12` |
| `ProductivityMenu` | method | `apps/powerpoint-vsto/DemoRibbon.cs:49` |
| `ElementMenu` | method | `apps/powerpoint-vsto/DemoRibbon.cs:55` |
| `Workflow` | method | `apps/powerpoint-vsto/DemoRibbon.cs:67` |
| `Tools` | method | `apps/powerpoint-vsto/DemoRibbon.cs:68` |
| `Group` | method | `apps/powerpoint-vsto/DemoRibbon.cs:75` |
| `Button` | method | `apps/powerpoint-vsto/DemoRibbon.cs:94` |
| `GridPlanner` | class | `apps/powerpoint-vsto/DemoTools.cs:11` |
| `Plan` | method | `apps/powerpoint-vsto/DemoTools.cs:13` |
| `Finite` | method | `apps/powerpoint-vsto/DemoTools.cs:27` |
| `DemoTool` | class | `apps/powerpoint-vsto/DemoTools.cs:29` |
| `ToString` | method | `apps/powerpoint-vsto/DemoTools.cs:33` |
| `DemoTools` | class | `apps/powerpoint-vsto/DemoTools.cs:38` |
| `Selected` | method | `apps/powerpoint-vsto/DemoTools.cs:44` |
| `Chart` | method | `apps/powerpoint-vsto/DemoTools.cs:52` |
| `BuildCatalog` | method | `apps/powerpoint-vsto/DemoTools.cs:59` |
| `Execute` | method | `apps/powerpoint-vsto/DemoTools.cs:138` |
| `CreateTour` | method | `apps/powerpoint-vsto/DemoTools.cs:139` |
| `Source` | method | `apps/powerpoint-vsto/DemoTools.cs:151` |
| `Text` | method | `apps/powerpoint-vsto/DemoTools.cs:153` |
| `Box` | method | `apps/powerpoint-vsto/DemoTools.cs:159` |
| `InsertChart` | method | `apps/powerpoint-vsto/DemoTools.cs:165` |
| `InsertCombo` | method | `apps/powerpoint-vsto/DemoTools.cs:186` |
| `GroupNew` | method | `apps/powerpoint-vsto/DemoTools.cs:187` |
| `ProcessFlow` | method | `apps/powerpoint-vsto/DemoTools.cs:196` |
| `Diagram` | method | `apps/powerpoint-vsto/DemoTools.cs:197` |
| `InsertGantt` | method | `apps/powerpoint-vsto/DemoTools.cs:219` |
| `InsertMekko` | method | `apps/powerpoint-vsto/DemoTools.cs:220` |
| `InsertWaterfall` | method | `apps/powerpoint-vsto/DemoTools.cs:221` |
| `Arrange` | method | `apps/powerpoint-vsto/DemoTools.cs:222` |
| `Grid` | method | `apps/powerpoint-vsto/DemoTools.cs:245` |
| `InsertTable` | method | `apps/powerpoint-vsto/DemoTools.cs:252` |
| `Agenda` | method | `apps/powerpoint-vsto/DemoTools.cs:258` |
| `PasteText` | method | `apps/powerpoint-vsto/DemoTools.cs:259` |
| `Font` | method | `apps/powerpoint-vsto/DemoTools.cs:260` |
| `SwitchDecimals` | method | `apps/powerpoint-vsto/DemoTools.cs:261` |
| `Language` | method | `apps/powerpoint-vsto/DemoTools.cs:262` |
| `ChartStyle` | method | `apps/powerpoint-vsto/DemoTools.cs:263` |
| `ValueLine` | method | `apps/powerpoint-vsto/DemoTools.cs:264` |
| `Difference` | method | `apps/powerpoint-vsto/DemoTools.cs:265` |
| `Cagr` | method | `apps/powerpoint-vsto/DemoTools.cs:266` |
| `Scan` | method | `apps/powerpoint-vsto/DemoTools.cs:267` |
| `InsertImage` | method | `apps/powerpoint-vsto/DemoTools.cs:268` |
| `ImportSlides` | method | `apps/powerpoint-vsto/DemoTools.cs:269` |
| `ApplyTemplate` | method | `apps/powerpoint-vsto/DemoTools.cs:270` |
| `ExportSlide` | method | `apps/powerpoint-vsto/DemoTools.cs:271` |
| `Brand` | method | `apps/powerpoint-vsto/DemoTools.cs:272` |
| `AxisTitle` | method | `apps/powerpoint-vsto/DemoTools.cs:273` |
| `CleanComments` | method | `apps/powerpoint-vsto/DemoTools.cs:274` |
| `SaveSlides` | method | `apps/powerpoint-vsto/DemoTools.cs:280` |
| `RoundData` | method | `apps/powerpoint-vsto/DemoTools.cs:292` |
| `StyleSheet` | method | `apps/powerpoint-vsto/DemoTools.cs:301` |
| `ImportJson` | method | `apps/powerpoint-vsto/DemoTools.cs:302` |
| `ExportJson` | method | `apps/powerpoint-vsto/DemoTools.cs:314` |
| `MiniToolbar` | method | `apps/powerpoint-vsto/DemoTools.cs:321` |
| `TableData` | method | `apps/powerpoint-vsto/DemoTools.cs:328` |
| `Library` | method | `apps/powerpoint-vsto/DemoTools.cs:348` |
| `AgendaCurrent` | method | `apps/powerpoint-vsto/DemoTools.cs:362` |
| `StyleFile` | method | `apps/powerpoint-vsto/DemoTools.cs:363` |
| `ExcelCagr` | method | `apps/powerpoint-vsto/DemoTools.cs:373` |
| `Prompt` | method | `apps/powerpoint-vsto/DemoTools.cs:379` |
| `OpenFile` | method | `apps/powerpoint-vsto/DemoTools.cs:389` |
| `SaveFile` | method | `apps/powerpoint-vsto/DemoTools.cs:394` |
| `ChooseFolder` | method | `apps/powerpoint-vsto/DemoTools.cs:399` |
| `Confirm` | method | `apps/powerpoint-vsto/DemoTools.cs:404` |
| `Notify` | method | `apps/powerpoint-vsto/DemoTools.cs:406` |
| `InsertAsset` | method | `apps/powerpoint-vsto/DemoTools.cs:408` |
| `DemoStyle` | class | `apps/powerpoint-vsto/DemoTools.cs:416` |
| `RibbonIcons` | class | `apps/powerpoint-vsto/RibbonIcons.cs:13` |
| `Key` | method | `apps/powerpoint-vsto/RibbonIcons.cs:22` |
| `Load` | method | `apps/powerpoint-vsto/RibbonIcons.cs:23` |
| `PictureBridge` | class | `apps/powerpoint-vsto/RibbonIcons.cs:44` |
| `Convert` | method | `apps/powerpoint-vsto/RibbonIcons.cs:47` |
| `Render` | method | `apps/powerpoint-vsto/RibbonIcons.cs:49` |
| `Line` | method | `apps/powerpoint-vsto/RibbonIcons.cs:65` |
| `Rect` | method | `apps/powerpoint-vsto/RibbonIcons.cs:67` |
| `Outline` | method | `apps/powerpoint-vsto/RibbonIcons.cs:69` |
| `Circle` | method | `apps/powerpoint-vsto/RibbonIcons.cs:71` |
| `Polygon` | method | `apps/powerpoint-vsto/RibbonIcons.cs:73` |
| `Letter` | method | `apps/powerpoint-vsto/RibbonIcons.cs:75` |
| `Arrow` | method | `apps/powerpoint-vsto/RibbonIcons.cs:81` |
| `Axes` | method | `apps/powerpoint-vsto/RibbonIcons.cs:91` |
| `Chart` | method | `apps/powerpoint-vsto/RibbonIcons.cs:92` |
| `Grid` | method | `apps/powerpoint-vsto/RibbonIcons.cs:151` |
| `Document` | method | `apps/powerpoint-vsto/RibbonIcons.cs:158` |
| `Slides` | method | `apps/powerpoint-vsto/RibbonIcons.cs:160` |
| `Folder` | method | `apps/powerpoint-vsto/RibbonIcons.cs:162` |
| `Palette` | method | `apps/powerpoint-vsto/RibbonIcons.cs:164` |
| `Refresh` | method | `apps/powerpoint-vsto/RibbonIcons.cs:166` |
| `Magnify` | method | `apps/powerpoint-vsto/RibbonIcons.cs:168` |
| `Chain` | method | `apps/powerpoint-vsto/RibbonIcons.cs:170` |
| `Align` | method | `apps/powerpoint-vsto/RibbonIcons.cs:172` |
| `TextLines` | method | `apps/powerpoint-vsto/RibbonIcons.cs:189` |
| `TextFrameIcon` | method | `apps/powerpoint-vsto/RibbonIcons.cs:194` |
| `Gallery` | method | `apps/powerpoint-vsto/RibbonIcons.cs:206` |
| `Gear` | method | `apps/powerpoint-vsto/RibbonIcons.cs:211` |
| `WorkflowIcon` | method | `apps/powerpoint-vsto/RibbonIcons.cs:216` |
| `Draw` | method | `apps/powerpoint-vsto/RibbonIcons.cs:246` |
| `DemoWorkspace` | class | `apps/powerpoint-vsto/DemoWorkspace.cs:8` |
| `RefreshList` | method | `apps/powerpoint-vsto/DemoWorkspace.cs:40` |
| `Selection` | method | `apps/powerpoint-vsto/DemoWorkspace.cs:53` |
| `FeatureInventoryEntry` | class | `apps/powerpoint-vsto/FeatureInventory.cs:2` |
| `FeatureInventory` | class | `apps/powerpoint-vsto/FeatureInventory.cs:3` |
| `DemoInputs` | class | `apps/powerpoint-vsto/DemoInputs.cs:9` |
| `IDemoAutomation` | interface | `apps/powerpoint-vsto/DemoApi.cs:12` |
| `DemoAutomation` | class | `apps/powerpoint-vsto/DemoApi.cs:22` |
| `GetVersion` | method | `apps/powerpoint-vsto/DemoApi.cs:29` |
| `GetBuildInfo` | method | `apps/powerpoint-vsto/DemoApi.cs:30` |
| `GetRibbonXml` | method | `apps/powerpoint-vsto/DemoApi.cs:31` |
| `GetRibbonLoaded` | method | `apps/powerpoint-vsto/DemoApi.cs:32` |
| `GetContextPropertiesState` | method | `apps/powerpoint-vsto/DemoApi.cs:34` |
| `InsertChartJson` | method | `apps/powerpoint-vsto/DemoApi.cs:35` |

## Per-instance input ports (testability boundary)

`apps/powerpoint-vsto/DemoInputs.cs:9` defines these delegate ports. Null means use the production native dialog or behavior, rather than simulated success. Supplying a port exercises command logic, not the native picker's appearance.

| Port | Input / output contract |
|---|---|
| Prompt | prompt/default strings -> chosen string or cancellation |
| OpenFile | extension-filter string -> selected path |
| SaveFile | filter/default filename -> output path |
| Folder | no arguments -> selected folder |
| Confirm | message/caption -> boolean |
| Notify | message/caption -> void |
| EditTable | two-dimensional string values -> edited matrix |
| ChooseAsset | candidate string list -> selected index |
| ShowToolbar | Form -> void (test/UI handoff) |
| PresentationCreated | native Presentation -> void (observation hook) |

These are internal dependency-injection seams, not a generalized plugin execution engine or externally authenticated API.
