# Managed chapters/contents and legacy presentation agenda

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

Keep editable agenda slides synchronized with chapter hierarchy, page references, hyperlinks, scope and native formatting.

Implemented local source is distinguished from supported subset/absent external behavior below. See current test evidence in the export testing chapter; do not treat historical logs or method existence as a current runtime pass.

## Source locations and internal components

- `apps/powerpoint-vsto/ChapterAgenda.cs:1` — ChapterAgenda, Stamp, ViewState, CaptureState.
- `apps/powerpoint-vsto/ChapterAgendaModel.cs:1` — ChapterRecord, AgendaView, ChapterCaptionEdit, ChapterAgendaState, ChapterAgendaEntry, ChapterAgendaPlan, ChapterMovePlan, ChapterAgendaPlanner, ChapterAgendaCodec.
- `apps/powerpoint-vsto/PresentationAgenda.cs:1` — AgendaSlideInfo, AgendaSectionInfo, AgendaEntry, AgendaPlan, AgendaPlanner, PresentationAgenda.

## Inputs, outputs, data structures and dependencies

Inputs: managed slides/field captions, hierarchy1–4, scope divider/order. Outputs: editable fields, native hyperlinks, slide tags and deck-level ChapterAgendaCodec XML. Dependencies: pure planners/codecs, XmlPartUpdates, slide duplicate/move and add-in debounce/BeforeSave.

## Internal execution and external integration

ChapterAgenda captures schema2managed chapter/views, resolves caption changes against baseline and validates owned fields. RefreshAll plans all views before writes, applies rows/page/hyperlink changes and verifies deck registry; immediate parent-child agenda ordering can auto-hide parent chapter views. Insert duplicates a managed prototype with new owner or creates blank using active Design. Manager edits captions/levels/show-all/layout frame and moves chapter/content subtrees. SplitHere marks a separate non-agenda divider. Legacy schema1 PresentationAgenda derives from slide titles/sections and refreshes through a separate dispatch path.

## Configuration, validation and failure behavior

Caption<=240 chars; <=4000 synchronized fields; layout frame must fit slide/minimum geometry. Automatic refresh skips active text editing. Scope anchors use stable native SlideID, not slide index alone.

## Security considerations and cleanup

Local native calls run in the current Office user process and require supported selection/ownership; a mutex/selection key is not authentication or RBAC. File/metadata content must remain bounded and error text must not expose data or credentials. Explicit owned-resource cleanup and error/rollback paths are part of the cited implementation, not a guarantee that a failed COM server can always be restored.

## Test references

Lexical module-class matches in test source (not instrumentation coverage; no test is claimed executed here):

- `tests/s0/ChapterAgendaOfficeTests.cs:41`.
- `tests/s0/ChapterAgendaTests.cs:11`.
- `tests/s0/DemoScenario.cs:147`.
- `tests/s0/PresentationAgendaTests.cs:11`.
- `tests/s0/WorkflowOfficeTests.cs:20`.

## Known limitations and technical debt

Caption conflicts, owned fields grouped/foreign/missing and invalid hierarchy/layout fail closed. Two agenda schemas coexist. Moves/updates compensate ordering/fields/XML; incomplete restoration directs Undo. Full reference-product agenda parity is not established.

## Extension points

Extend hierarchy/planner and schema migration first; test copied views, conflicts, scopes, page/hyperlink updates, formatting, movement rollback and save/reopen.

## Classes, interfaces and method navigation

Method names below are source declarations; private/internal members are not public API promises. The source hash/review manifest gives exact reviewed bytes.

| Declaration | Kind | Source |
|---|---|---|
| `ChapterAgenda` | class | `apps/powerpoint-vsto/ChapterAgenda.cs:12` |
| `Stamp` | class | `apps/powerpoint-vsto/ChapterAgenda.cs:16` |
| `ViewState` | class | `apps/powerpoint-vsto/ChapterAgenda.cs:18` |
| `CaptureState` | class | `apps/powerpoint-vsto/ChapterAgenda.cs:19` |
| `Active` | method | `apps/powerpoint-vsto/ChapterAgenda.cs:20` |
| `IsManaged` | method | `apps/powerpoint-vsto/ChapterAgenda.cs:22` |
| `GroupContainsOwned` | method | `apps/powerpoint-vsto/ChapterAgenda.cs:23` |
| `Fields` | method | `apps/powerpoint-vsto/ChapterAgenda.cs:28` |
| `Capture` | method | `apps/powerpoint-vsto/ChapterAgenda.cs:33` |
| `NewField` | method | `apps/powerpoint-vsto/ChapterAgenda.cs:48` |
| `Retain` | method | `apps/powerpoint-vsto/ChapterAgenda.cs:52` |
| `Apply` | method | `apps/powerpoint-vsto/ChapterAgenda.cs:57` |
| `RefreshAll` | method | `apps/powerpoint-vsto/ChapterAgenda.cs:70` |
| `InsertChapter` | method | `apps/powerpoint-vsto/ChapterAgenda.cs:84` |
| `InsertContents` | method | `apps/powerpoint-vsto/ChapterAgenda.cs:85` |
| `Insert` | method | `apps/powerpoint-vsto/ChapterAgenda.cs:86` |
| `SplitHere` | method | `apps/powerpoint-vsto/ChapterAgenda.cs:95` |
| `Read` | method | `apps/powerpoint-vsto/ChapterAgenda.cs:97` |
| `ChangeChapter` | method | `apps/powerpoint-vsto/ChapterAgenda.cs:98` |
| `ShowManager` | method | `apps/powerpoint-vsto/ChapterAgenda.cs:105` |
| `Move` | method | `apps/powerpoint-vsto/ChapterAgenda.cs:120` |
| `Automatic` | method | `apps/powerpoint-vsto/ChapterAgenda.cs:125` |
| `ChapterRecord` | class | `apps/powerpoint-vsto/ChapterAgendaModel.cs:10` |
| `AgendaView` | class | `apps/powerpoint-vsto/ChapterAgendaModel.cs:11` |
| `ChapterCaptionEdit` | class | `apps/powerpoint-vsto/ChapterAgendaModel.cs:12` |
| `ChapterAgendaState` | class | `apps/powerpoint-vsto/ChapterAgendaModel.cs:13` |
| `ChapterAgendaEntry` | class | `apps/powerpoint-vsto/ChapterAgendaModel.cs:18` |
| `ChapterAgendaPlan` | class | `apps/powerpoint-vsto/ChapterAgendaModel.cs:25` |
| `ChapterMovePlan` | class | `apps/powerpoint-vsto/ChapterAgendaModel.cs:26` |
| `ChapterAgendaPlanner` | class | `apps/powerpoint-vsto/ChapterAgendaModel.cs:27` |
| `Move` | method | `apps/powerpoint-vsto/ChapterAgendaModel.cs:29` |
| `Caption` | method | `apps/powerpoint-vsto/ChapterAgendaModel.cs:52` |
| `Validate` | method | `apps/powerpoint-vsto/ChapterAgendaModel.cs:58` |
| `Build` | method | `apps/powerpoint-vsto/ChapterAgendaModel.cs:66` |
| `ResolveEdits` | method | `apps/powerpoint-vsto/ChapterAgendaModel.cs:82` |
| `ChapterAgendaCodec` | class | `apps/powerpoint-vsto/ChapterAgendaModel.cs:90` |
| `State` | method | `apps/powerpoint-vsto/ChapterAgendaModel.cs:94` |
| `Encode` | method | `apps/powerpoint-vsto/ChapterAgendaModel.cs:98` |
| `Shape` | method | `apps/powerpoint-vsto/ChapterAgendaModel.cs:102` |
| `Number` | method | `apps/powerpoint-vsto/ChapterAgendaModel.cs:106` |
| `Decode` | method | `apps/powerpoint-vsto/ChapterAgendaModel.cs:107` |
| `AgendaSlideInfo` | class | `apps/powerpoint-vsto/PresentationAgenda.cs:10` |
| `AgendaSectionInfo` | class | `apps/powerpoint-vsto/PresentationAgenda.cs:16` |
| `AgendaEntry` | class | `apps/powerpoint-vsto/PresentationAgenda.cs:21` |
| `AgendaPlan` | class | `apps/powerpoint-vsto/PresentationAgenda.cs:29` |
| `AgendaPlanner` | class | `apps/powerpoint-vsto/PresentationAgenda.cs:35` |
| `Build` | method | `apps/powerpoint-vsto/PresentationAgenda.cs:37` |
| `Entry` | method | `apps/powerpoint-vsto/PresentationAgenda.cs:55` |
| `Finite` | method | `apps/powerpoint-vsto/PresentationAgenda.cs:57` |
| `PresentationAgenda` | class | `apps/powerpoint-vsto/PresentationAgenda.cs:59` |
| `Capture` | method | `apps/powerpoint-vsto/PresentationAgenda.cs:62` |
| `Insert` | method | `apps/powerpoint-vsto/PresentationAgenda.cs:69` |
| `RefreshSelected` | method | `apps/powerpoint-vsto/PresentationAgenda.cs:75` |
| `Active` | method | `apps/powerpoint-vsto/PresentationAgenda.cs:79` |
| `ContainsOwnedField` | method | `apps/powerpoint-vsto/PresentationAgenda.cs:81` |
| `Apply` | method | `apps/powerpoint-vsto/PresentationAgenda.cs:83` |
| `Upsert` | method | `apps/powerpoint-vsto/PresentationAgenda.cs:96` |
