# Workflow inventory

**VERIFIED:** routes below exist in source. Execution evidence is recorded separately; source-supported diagrams are not fresh Office test passes.

| Workflow | Trigger and input | Output | Primary source | Failure boundary |
|---|---|---|---|---|
| Startup | VSTO Startup | Host, timers, compatibility and context controller | `apps/powerpoint-vsto/ThisAddIn.cs:37` | Invalid config disables commands |
| Ribbon action | Callback plus Tag/ID/current selection | Native document mutation or dialog | `apps/powerpoint-vsto/ThisAddIn.cs:133` | Busy/lease/compatibility and status catches |
| Context property update | Selection event/timer and foreground policy | Immutable UI-state snapshot and applicable controls | `apps/powerpoint-vsto/ContextPropertiesController.cs:1` | Stale selection or absent foreground prevents mutation |
| Linked chart write | Valid local table and explicit source link | Embedded workbook, native series and attached source metadata | `apps/powerpoint-vsto/DataSources.cs:166` | Native-value acknowledgement/timeout; no sample fallback |
| Table insertion | Linked rows/columns | Editable native PowerPoint table and link metadata | `apps/powerpoint-vsto/DemoTools.cs:247` | Partial owned shape cleaned up on failure |
| Native shape rendering | v2 canonical chart plus verified scene | Grouped editable native shapes and canonical binding metadata | `src/Chart.Renderers/NativeRenderingCore.cs:158` | Preflight then owned-object rollback; incomplete rollback explicit |
| Managed element editing | Element command/double click | Updated text/process/indicator/connector canonical/native state | `apps/powerpoint-vsto/ElementCommands.cs:1` | Unsupported/ambiguous metadata preserved |
| Agenda refresh | Chapter controls, debounce or before-save | Chapter/contents text/hierarchy/links | `apps/powerpoint-vsto/ChapterAgenda.cs:1` | Invalid state reports paused update |
| Library indexing | Registered root, refresh/watch request | Transactionally refreshed PIIDX catalog | `src/Library.Core/PiidxCatalog.cs:1` | Exact schema/integrity/path parsing guards |
| Library preview/import | Selected library slide and destination | Native preview or inserted source-formatted slide | `apps/powerpoint-vsto/LocalLibraryGallery.cs:1` | Missing/stale source and owned hidden-source cleanup |
| Save/reopen/recover | BeforeSave or explicit inspect | Associated chart data/revisions and duplicate identity handling | `apps/powerpoint-vsto/PowerPointHost.cs:48` | Preserve unsupported objects with diagnostics |
| Service profile preparation | Configure services | Validated local profile/status | `apps/powerpoint-vsto/ServiceConnections.cs:1` | Not an authenticated network call |
| Shutdown | VSTO Shutdown | Disposed timers/controls and event detachment | `apps/powerpoint-vsto/ThisAddIn.cs:77` | Scoped disposal, no forced Office termination |

The [sequence diagrams](sequence-diagrams.md) and [activity diagrams](activity-diagrams.md) explain significant paths. AI inference has no implemented runtime workflow; its absence and only actual profile-preparation path are documented in [AI workflows](../06-ai-architecture/ai-workflows.md).
