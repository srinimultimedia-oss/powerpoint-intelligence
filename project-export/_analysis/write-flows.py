from pathlib import Path
E=Path('project-export')
def put(path,title,s):(E/path).write_text('# '+title+'\n\n'+s.strip()+'\n',encoding='utf-8')
put('10-workflows/workflow-inventory.md','Workflow inventory', '''**VERIFIED:** routes below exist in source. Execution evidence is recorded separately; source-supported diagrams are not fresh Office test passes.

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

The [sequence diagrams](sequence-diagrams.md) and [activity diagrams](activity-diagrams.md) explain significant paths. AI inference has no implemented runtime workflow; its absence and only actual profile-preparation path are documented in [AI workflows](../06-ai-architecture/ai-workflows.md).''')
put('10-workflows/sequence-diagrams.md','Execution sequence diagrams', '''## Add-in initialization and teardown

```mermaid
sequenceDiagram
  participant P as PowerPoint VSTO
  participant A as ThisAddIn
  participant H as PowerPointHost
  participant C as ContextPropertiesController
  P->>A: Startup
  A->>H: Construct on UI STA
  A->>A: ConfigureCompatibility and create timers
  A->>C: Construct guarded context controller
  A->>P: Subscribe selection open save activate close
  P->>A: User selection changed
  A->>C: Request observation
  A->>A: Invalidate ribbon and restart debounce
  P->>A: Shutdown
  A->>C: Dispose
  A->>A: Dispose timers and layout controls
  A->>P: Unsubscribe events and remove owned task pane
```

Source: `apps/powerpoint-vsto/ThisAddIn.cs:37`, `apps/powerpoint-vsto/ThisAddIn.cs:77`. Compatibility does not validate Office edition licensing. Context observation and host mutation have different permissions/lifetimes.

## Ribbon action and native table

```mermaid
sequenceDiagram
  actor U as Author
  participant R as Ribbon callback
  participant A as ThisAddIn
  participant L as Mutation lease
  participant D as LinkedChartData
  participant P as Active PowerPoint slide
  U->>R: Choose editable table
  R->>A: RunDemoTool with command tag
  A->>L: TryEnter current Office PID
  alt Busy or incompatible
    A-->>U: No mutation or compatibility status
  else Allowed
    A->>D: Choose and load linked source
    D-->>A: Validated table and selected mapping
    A->>P: Shapes.AddTable and fill native cells
    A->>D: Persist shape link metadata
    A->>P: Select created object
    A-->>U: Completed or scoped failure status
  end
```

The native table path writes editable cells with central vertical anchoring and centered paragraphs. It does not rasterize the result or create another report presentation. `apps/powerpoint-vsto/DemoTools.cs:247`; `apps/powerpoint-vsto/ThisAddIn.cs:133`.

## Linked chart workbook commit

```mermaid
sequenceDiagram
  participant C as LinkedChartData
  participant E as Embedded Excel workbook
  participant N as Native PowerPoint chart
  participant X as Shape custom XML
  C->>E: Activate and write validated linked table
  C->>N: Bind source and set typed series values
  C->>N: Wait for expected native series and point values
  alt Native values acknowledged
    C->>E: Calculate and close owned workbook
    C->>N: Verify committed native values
    C->>X: Persist source association
  else Timeout or COM failure
    C-->>C: Throw stage-specific diagnostic
  end
```

`apps/powerpoint-vsto/DataSources.cs:166`. Closing an Excel workbook alone is not evidence that PowerPoint committed values; native chart observation is the acceptance condition. Fixed diagnostics deliberately omit source values. No sample-data fallback is permitted.

## Grouped-shape renderer transaction

```mermaid
sequenceDiagram
  participant R as NativeSceneRenderer
  participant V as ClusteredProjection
  participant S as IEditableSceneSurface
  R->>V: Validate canonical JSON and scene equality
  V-->>R: Validated projection
  R->>S: Begin owned batch
  loop Drawable semantic nodes
    R->>S: Create native shape and bind semantic ID
  end
  R->>S: Group and persist canonical scene bindings
  alt Success
    S-->>R: Verified group record
  else Mutation or persistence failure
    R->>S: Roll back only newly owned objects
    S-->>R: Rollback completeness
    R-->>R: Throw NativeRenderException
  end
```

`src/Chart.Renderers/NativeRenderingCore.cs:158`; `src/Chart.Renderers/PowerPointSceneSurface.cs:14`. Zero/missing values can have semantic bindings without drawable bars. Incomplete cleanup remains an error, rather than a success or permission to delete unrelated shapes.

## Library refresh, query and insertion

```mermaid
sequenceDiagram
  actor U as Author
  participant G as Gallery workspace
  participant C as PiidxCatalog
  participant F as Local source packages
  participant P as PowerPoint
  U->>G: Register source or refresh
  G->>C: Refresh with source registration
  C->>F: Read bounded Open XML and fingerprints
  C->>C: Transactionally update catalog and search index
  U->>G: Query filter and select thumbnail
  G->>C: Bounded search page and source reference
  G->>P: Render preview from owned read-only source if needed
  U->>G: Insert selected slide
  G->>P: Insert source slide into current destination
  G->>P: Close only owned source presentation
```

`src/Library.Core/PiidxCatalog.cs:1`; `src/Library.Core/PresentationSource.cs:1`; `apps/powerpoint-vsto/LocalLibraryGallery.cs:1`. PIIDX extraction does not itself render Office pixels. Source freshness and native-format fidelity are separate checks.

## Save and recovery

```mermaid
sequenceDiagram
  participant P as PowerPoint
  participant A as ThisAddIn
  participant H as PowerPointHost
  participant X as Custom XML and tags
  P->>A: BeforeSave active presentation
  A->>A: Refresh managed agenda and layout links
  A->>H: InspectAndRecover capture enabled
  H->>X: Resolve shape-owned metadata namespace
  alt Supported valid chart
    H->>P: Capture native data and reconcile identity
    H->>X: Persist validated metadata
  else Missing ambiguous or unsupported metadata
    H-->>A: Fixed diagnostic and preserve native object
  end
  P->>P: Save presentation
```

`apps/powerpoint-vsto/ThisAddIn.cs:318`; `apps/powerpoint-vsto/PowerPointHost.cs:48`. Open/window activation do not unconditionally mutate all decks. Do not infer that event-level catching means a failed mutation was fully transactional.''')
put('10-workflows/activity-diagrams.md','Activity diagrams', '''## Validation before document mutation

```mermaid
flowchart TD
  A[User action] --> B{Host available and compatible}
  B -- No --> X[Report disabled or unavailable]
  B -- Yes --> C{Mutation lease acquired}
  C -- No --> Y[Return without competing write]
  C -- Yes --> D[Resolve current presentation and selection]
  D --> E[Read and validate canonical data or input table]
  E --> F{Supported and unambiguous}
  F -- No --> Z[Preserve native objects and report]
  F -- Yes --> G[Start native undo boundary where implemented]
  G --> H[Perform scoped native edits]
  H --> I[Verify commit and persist association]
  I --> J[Update status and context observation]
```

`apps/powerpoint-vsto/ThisAddIn.cs:133`; `apps/powerpoint-vsto/PowerPointHost.cs:48`. Undo grouping and transaction rollback differ by module; the graph is a shared validation pattern, not a claim every handler implements the same atomicity.

## Context properties and stale target prevention

```mermaid
flowchart TD
  A[Selection or timer request] --> B[Capture on Office UI STA]
  B --> C{PowerPoint owner foreground}
  C -- No --> H[Hide property box]
  C -- Yes --> D[Classify applicable properties and target identity]
  D --> E[Publish immutable observation snapshot]
  E --> F[User chooses property edit]
  F --> G{Selection identity still matches}
  G -- No --> R[Reject stale mutation]
  G -- Yes --> M[Apply only supported properties]
```

`apps/powerpoint-vsto/ContextPropertiesController.cs:1`; `apps/powerpoint-vsto/ContextSelectionBridge.cs:1`; `apps/powerpoint-vsto/ContextActionRouter.cs:1`. Foreground focus is a Windows prerequisite, not a reason to change focus/security policy. Table cell text, group text, chart properties and indicators have different applicability.

## Local catalog refresh

```mermaid
flowchart TD
  A[Configured enabled source root] --> B[Enumerate supported packages within policy]
  B --> C{Fingerprint changed or new}
  C -- No --> D[Retain current indexed presentation]
  C -- Yes --> E[Parse bounded presentation metadata]
  E --> F[Begin catalog transaction]
  F --> G[Replace presentation slides objects tags and FTS rows]
  G --> H[Commit or roll back error]
  H --> I[Report refresh progress and missing files]
```

`src/Library.Core/PiidxCatalog.cs:1`; `src/Library.Core/PresentationSource.cs:1`. Watch notifications/debounce request refresh; they are not a guarantee of enterprise scheduled updates or real-time external synchronization.''')
put('10-workflows/error-handling-flows.md','Error handling and recovery flows', '''**VERIFIED:** the application has several scoped failure policies rather than one global transaction manager.

| Boundary | Error policy | Preserved state | Consequence |
|---|---|---|---|
| Ribbon/user action | Catch cancellation/specific validation and generic operation failures; reset busy/request context | Existing user document when preflight rejects | Generic UI message can hide detailed root cause; use diagnostic tests |
| Office event callback | Catch all and emit exception type/HRESULT | Prevent managed exception escaping native event | Event survival is not proof no partial mutation |
| S0 metadata inspection | Missing/ambiguous/schema-incompatible metadata produces diagnostics | Native chart/object | Manual recovery may be needed |
| XML attachment replacement | Verify replacement association and attempt restore/cleanup | Prior metadata/native objects where rollback succeeds | Explicit replacement-verification failure must remain visible |
| Native scene insertion | Track and rollback newly owned objects; report cleanup completeness | Unrelated objects and original records | Partial rollback never counts as success |
| Embedded Excel | Wait for native values before/after close; fixed stage diagnostics | No invented sample fallback | RPC loss/busy Office is not repaired by retrying arbitrary COM actions |
| Linked source refresh | Stop on changed headers/missing source/unsupported chart | Existing native chart | User must confirm relink/mapping |
| PIIDX | Exact schema/integrity checks and transactional writes | Previous committed catalog rows | Invalid/corrupt file is not silently trusted |
| Context properties | Hide or reject on stale target/foreground mismatch | Selection that is not the observed target | Visual UI prerequisite can block appearance tests |
| Installer | Reject stale/hash/signature failure before publish; retain previous artifact | Prior published package | Requires approved signer/trust, not policy bypass |

Sources: `apps/powerpoint-vsto/PowerPointHost.cs:34`; `apps/powerpoint-vsto/XmlPartUpdates.cs:1`; `apps/powerpoint-vsto/DataSources.cs:166`; `src/Chart.Renderers/NativeRenderingCore.cs:158`; `src/Library.Core/PiidxCatalog.cs:1`; `easySetup/Build-EasySetup.ps1:1`.

```mermaid
flowchart LR
  E[Failure detected] --> P{Preflight or after mutation}
  P -- Preflight --> R[Reject and preserve existing objects]
  P -- After mutation --> O[Use module-owned cleanup or restore]
  O --> V{Cleanup verified}
  V -- Yes --> F[Report failed operation with rollback evidence]
  V -- No --> U[Report incomplete recovery and preserve evidence]
```

There is no implemented universal COM retry policy. Bounded waits are local to specific operations. Retry must respect target lifetime, STA ownership and identity; catching `0x800706BA` cannot safely recreate a disconnected object by itself.''')
put('10-workflows/data-flow-diagrams.md','Data-flow and trust-boundary diagrams', '''```mermaid
flowchart LR
  FILE[User-selected local file] --> PARSE[Bounded CSV JSON XML or Excel reader]
  PARSE --> MAP[Explicit columns and worksheet range]
  MAP --> TABLE[Validated linked table]
  TABLE --> CHART[Native Excel-backed PowerPoint chart]
  TABLE --> SHAPES[Canonical v2 projection and native shapes]
  CHART --> META[Shape-attached custom XML plus tags]
  SHAPES --> META
  CHART --> SAVE[PPTX save]
  META --> SAVE
  SAVE --> READ[Reopen and validated association read]
```

The embedded workbook may contain actual linked data. Canonical metadata and source path provenance are persisted locally with the presentation; encryption/redaction policy is not invented by the codec. Diagnostics intentionally avoid logging workbook values but do not encrypt the deck. `apps/powerpoint-vsto/DataSources.cs:39`; `apps/powerpoint-vsto/MetadataCodec.cs:1`; `src/Chart.Renderers/NativeRenderingCore.cs:71`.

```mermaid
flowchart LR
  PKG[Local PPTX POTX source] --> READ[ZIP and XML metadata reader]
  READ --> IDX[PIIDX SQLite metadata and FTS]
  IDX --> Q[Bounded gallery queries]
  Q --> REF[Source path and slide identity]
  REF --> PPT[Owned read-only PowerPoint source access]
  PPT --> PNG[Validated PNG preview cache]
  PPT --> DEST[Current destination slide insertion]
```

The index can contain slide text, author names, paths and other document metadata. Its purpose is local search, not content anonymization. Preview cache integrity and source-fingerprint checks do not establish confidentiality. `src/Library.Core/PiidxCatalog.cs:1`; `src/Library.Core/PiidxPreviewCache.cs:1`.

There is no verified data flow from presentation content to Azure OpenAI. Do not draw or document an authenticated AI request sequence until the gateway/client actually exists. Local configured-provider profiles are shown separately in [AI data flow](../06-ai-architecture/ai-data-flow.md).''')
print('workflow docs written')
