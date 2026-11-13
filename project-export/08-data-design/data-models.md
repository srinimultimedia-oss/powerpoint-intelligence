# Data models

**VERIFIED:** inspected implementation, not a claim of current Office execution. The repository contains several independent versioned models; version numbers across these formats are not interchangeable.

| Model boundary | Classes | Storage/version | Evidence | Meaning and ownership |
| --- | --- | --- | --- | --- |
| Legacy native chart | Host.ChartDocument / ChartSeries | XML schema1; shape CustomerData | `apps/powerpoint-vsto/Contracts.cs:17` | Categories are ordered strings; series are ordered identified names and nullable doubles. ChartId GUID, revision >=0, clustered-column only. Native formatting/layout is not stored by this DTO. |
| Canonical chart v2 | Domain.ChartDocument | JSON schema2; explicit codec/migration | `src/Chart.Domain/ChartDocument.cs:8` | Dataset and categories/series/axes have unique semantic IDs. Two category/value axes, point layout and series-color style. It is not automatically the metadata used by all native demo chart types. |
| Canonical authored v3 | ChartDocumentV3 / ChartOptions | JSON schema3 source, not production compiled | `src/Chart.Domain/ChartDocumentV3.cs:6` | Bar/line/area/pie options and categoryColors. Missing from explicit compile graphs: do not infer shipping behavior from file presence. |
| Semantic scene | RenderResult and geometry subtypes | JSON schema1 | `src/Chart.Renderers/SceneContracts.cs:7` | ChartId/modelRevision scoped nodes, geometry/paint, diagnostic references and verified renderHash. Broader DTO support than native profile. |
| Rendered native record | RenderedChartRecord / NativeBinding | XML rendered-chart:v1 | `src/Chart.Renderers/NativeRenderingCore.cs:65` | Canonical JSON plus sealed scene plus ordered node-token/shapeID hints. Zero-area nodes retain bindings with absent ShapeId. |
| Extracted source | IndexedPresentation / IndexedSlide / IndexedObject | In-memory extraction then relational storage | `src/Library.Core/PresentationSource.cs:16` | Package/native slide IDs separate from display positions. Object structural Key/ParentKey tolerates repeated native IDs; parent-space EMU preserved. |
| Search/registration | LibraryRegistration / SearchRequest/Page/SlideReference | In-memory local interfaces | `src/Library.Core/LibraryModels.cs:54` | Query modes/filters/pagination, source and slide fingerprints, source path/native slide identity, title/text/category/author/layout/master, dimensions/rank/time. |
| Library runtime | LibraryRuntimeProfile / LibrarySourceDefinition | Application/enterprise XML config | `src/Library.Core/LibraryConfiguration.cs:20` | Index and legacy locations, refresh/watch policy, roots with names/recursive/managed flags; Managed is not authorization. |
| Linked data | LinkedTable / DataSourceLink / ExcelRangeMapping | Versioned presentation/shape XML | `apps/powerpoint-vsto/DataSources.cs:254` | File-backed source identity/format/path/link timestamp and optional explicit worksheet/range/header/category/series mapping. No database tables for workbook data. |
| Agenda | ChapterAgendaState / ChapterRecord / AgendaView | Managed agenda:v2 XML | `apps/powerpoint-vsto/ChapterAgendaModel.cs:13` | Stable slide IDs, levels/scopes and caption baseline; derived entries carry numbering/highlight/layout. SlideIndex is position, never persistent identity authority. |
| Text element | TextShapeSpec | Text-shape XML ownership | `apps/powerpoint-vsto/TextShapeElements.cs:16` | Kind, edge-lock flags, native adjustment and AutoFit; native actual text/style stays in PowerPoint. |
| Process grid | ProcessGridSpec / ProcessGridCell | Process-grid:v1 presentation XML | `apps/powerpoint-vsto/ProcessFlowTools.cs:14` | Grid GUID+slideID+rows/columns/bounds/gap; cells have row/column/shapeID/active retained history. |
| Indicator | IndicatorSpec | Indicator:v1 shape XML | `apps/powerpoint-vsto/IndicatorElements.cs:11` | Harvey segments/filled or checkbox state, colors/size/rotation/custom symbol; native group ownership verified before edits. |
| Connector | ConnectorSpec / ConnectorEndpoint | Connector:v1 shape XML | `apps/powerpoint-vsto/UniversalConnectors.cs:14` | Endpoints site/point/free; native shape/site or chart series/point anchor, coordinates/rendered flag; route/arrows/dash/color/weight/autoroute. |
| Template projection | Template nodes/fields/palette/layout metadata | Template:v1 native group metadata | `apps/powerpoint-vsto/TemplateLibrary.cs:114` | Original source-slide library and generated editable-layout library are separate paths. Field text and palette/alignment controls do not make original licensed content independently authored. |

## Identity and ownership rules

- A chart GUID identifies a chart, semantic IDs identify domain children, node IDs identify one scoped projection and Office Shape.Id is a live host handle/hint. They are not interchangeable.
- Catalog presentation identity is independently generated; catalog SlideId is deterministic from that identity plus native slide ID. Positions can change without requiring a new slide identity. A uniquely hashed rename may retain presentation identity.
- The PIIDX preview is derivative data and can be regenerated. Canonical/native metadata is authoritative state for its own feature; no global single database covers every add-in object.
- Shared data DTOs have mutable CLR properties; boundary codecs clone/validate to prevent unvalidated aliasing. Most library APIs are internal, not network contracts.

## Units

Canonical and scene layouts use points. OOXML dimensions are converted from EMU using12700 units per point. Indexed object transforms intentionally stay raw parent-space EMU/rotation/flip metadata rather than pretending to be absolute coordinates. Native COM geometry uses Single precision; readback tolerances belong to that boundary, not arbitrary canonical rounding.

## Status and gaps

No application server database, cloud persistence, multi-tenant identity data model or general AI response executor is implemented by these core files. The SQLite schema below is a local library index. The v3 authored source is not wired into shipped compile graphs. Missing choices are represented by explicit options/rejection, not guessed defaults.

See [persistence design](persistence-design.md), [dictionary](data-dictionary.md), [class diagrams](class-diagrams.md) and [wire contracts](serialization-contracts.md).
