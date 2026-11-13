# Native editing and persistence mechanisms

## Three different editable chart/object paths

1. S0 native ChartDocument (`apps/powerpoint-vsto/Contracts.cs:17`) is a narrow clustered-column contract. PowerPointHost inserts a genuine chart with embedded workbook and shape-owned XML/GUID. It reads supported native edits into metadata during Check & Sync/save; it does not project old metadata back onto charts on reopen.
2. Linked native charts (`apps/powerpoint-vsto/DataSources.cs:166`) use DataSourceLink/LinkedTable and native chart source binding. Chart/table/custom diagram object references are different from S0 metadata. Mapped refresh is intentionally narrower than insertion.
3. Scene rendering (`apps/powerpoint-vsto/RendererRibbonCommand.cs:44`) creates validated native editable primitives via the core renderer/surface. Result is a shape group, not a native chart workbook; current adapter emits axis/legend omission diagnostics.

These paths must not be conflated in maintenance or migration. Metadata namespaces/tags identify the appropriate owner. Chart-family conversion on S0 objects is rejected by NativeSeriesControls to preserve its narrower semantics.

## Metadata and copied owners

CustomerData binds an Office custom XML data-store part to native shapes. Shape numeric IDs are document locations, not globally canonical chart identity. S0 uses GUID plus PPI_CHART_ID. Shared/duplicated Office parts can require preserving every copied owner, not just checking an RCW returned from Add.

`apps/powerpoint-vsto/XmlPartUpdates.cs:47` resolves part membership by normalized GUID and reads root namespace from serialized XML instead of trusting cached NamespaceURI. Shape replacement (`:63`) validates one current attachment, scans shared owners, detaches/adds/verifies and restores old owners on failure. Deck replacement (`:76`) requires a unique registry. Semantic XML comparison retains significant text/whitespace; it is not a byte-equality requirement.

```mermaid
sequenceDiagram
  participant Adapter
  participant XML as XmlPartUpdates
  participant Owner as Native owner CustomerData
  participant Shared as Copied owners
  Adapter->>XML: Validated replacement XML and current part
  XML->>Owner: Verify GUID/namespace/payload
  XML->>Shared: Snapshot owners of prior part
  XML->>Owner: Detach old; create new
  XML->>Owner: Resolve live collection; verify final attachment
  XML->>Shared: Restore preserved old payload where needed
  alt failure
    XML->>Owner: Remove owned replacement if attached
    XML->>Shared: Attempt original restoration
    XML-->>Adapter: Failure; incomplete restoration surfaced
  else verified
    XML-->>Adapter: Current verified attached part
  end
```

## Native text/properties

Managed text, process cells, indicators/connectors and template fields are native objects; text remains directly editable. NativeFormatting and ContextFormatting snapshot requested properties and support bounded native groups/tables/selected characters. A selection key prevents a delayed picker from changing a different object. Indicator/connector colors route through canonical specs so native appearance and persisted settings stay aligned.

## Undo/redo and atomicity

Adapters call Application.StartNewUndoEntry around explicit native mutations. There is no separate full-document undo engine or reliable cross-process transaction system. Some complex adapters implement compensation; simple catalog mutations rely on PowerPoint Undo. Compensation failures retain evidence/objects and report Undo guidance. Native user Undo/Redo and other add-ins can change identities; the next read validates rather than trusting cached state.

## Save/reopen boundaries

PowerPoint saves its own native document plus custom XML/tags; adapters never rewrite the ZIP behind the running host. Missing/unknown metadata is preserved and blocks unsafe automatic edits. Cross-presentation explicit S0 copy captures original model before clipboard transfer and writes a new destination attachment when clipboard omits CustomerData (`apps/powerpoint-vsto/PowerPointHost.cs:231`). This helper is not a promise that arbitrary manual clipboard operations retain every custom schema.
