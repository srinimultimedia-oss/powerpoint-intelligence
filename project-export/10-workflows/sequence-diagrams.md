# Execution sequence diagrams

<!-- generated-toc -->
## Contents

- [Add-in initialization and teardown](#add-in-initialization-and-teardown)
- [Ribbon action and native table](#ribbon-action-and-native-table)
- [Linked chart workbook commit](#linked-chart-workbook-commit)
- [Grouped-shape renderer transaction](#grouped-shape-renderer-transaction)
- [Library refresh, query and insertion](#library-refresh-query-and-insertion)
- [Save and recovery](#save-and-recovery)
<!-- end-generated-toc -->

## Add-in initialization and teardown

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

`apps/powerpoint-vsto/ThisAddIn.cs:318`; `apps/powerpoint-vsto/PowerPointHost.cs:48`. Open/window activation do not unconditionally mutate all decks. Do not infer that event-level catching means a failed mutation was fully transactional.
