# Class and relationship diagrams

<!-- generated-toc -->
## Contents

- [Canonical v2 domain](#canonical-v2-domain)
- [Scene and editable host boundary](#scene-and-editable-host-boundary)
- [Catalog extraction and query objects](#catalog-extraction-and-query-objects)
- [Storage ERD](#storage-erd)
<!-- end-generated-toc -->

**VERIFIED structural reconstruction.** Diagrams show actual authored classes and relationships; compile status and runtime capability are called out separately. Field subsets keep diagrams readable; the dictionary contains complete opt-in wire fields and SQL columns.

## Canonical v2 domain

```mermaid
classDiagram
    class ChartDocument {
      int SchemaVersion
      string ChartId
      int Revision
      string ChartType
    }
    class EmbeddedDataset {
      string Id
    }
    class ChartCategory {
      string Id
      string Label
    }
    class ChartSeries {
      string Id
      string Name
    }
    class ChartAxis {
      string Id
      string Role
      string Scale
    }
    class ChartLayout {
      string Unit
      double X
      double Y
      double Width
      double Height
    }
    class ChartStyle {
      bool ShowLegend
    }
    class ChartCodec {
      Encode()
      Decode()
      Validate()
    }
    ChartDocument *-- EmbeddedDataset : dataset
    EmbeddedDataset *-- ChartCategory : categories
    ChartDocument *-- ChartSeries : series
    ChartDocument *-- ChartAxis : axes
    ChartDocument *-- ChartLayout : layout
    ChartDocument *-- ChartStyle : style
    ChartCodec ..> ChartDocument : validates
```

Evidence: `src/Chart.Domain/ChartDocument.cs:8`; `src/Chart.Domain/ChartCodec.cs:15`. ChartSeries.Values is an ordered list of nullable doubles with exactly the category count. Style.SeriesColors cardinality matches series. This v2 model is compiled into VSTO; authored ChartDocumentV3/ChartOptions are deliberately not conflated here because explicit production compile items omit them.

## Scene and editable host boundary

```mermaid
classDiagram
    class RenderResult {
      int SchemaVersion
      string ChartId
      int ModelRevision
      string RenderHash
    }
    class SceneNode {
      string Id
    }
    class SceneGeometry {
      string Type
    }
    class BoxGeometry {
      double X
      double Y
      double Width
      double Height
    }
    class TextGeometry {
      string Text
      string FontFamily
      double FontSize
      string Alignment
    }
    class PathGeometry
    class SectorGeometry {
      double Cx
      double Cy
      double Radius
      double StartAngle
      double Sweep
    }
    class ScenePaint {
      string Fill
      string Stroke
      double StrokeWidth
      double Opacity
    }
    class SceneCodec {
      Encode()
      Decode()
      ComputeHash()
    }
    class IEditableSceneSurface {
      <<interface>>
      VerifyAccess()
      Begin()
      CreateAnchor()
      CreateNode()
      Group()
      Persist()
      Rollback()
      Complete()
    }
    class NativeSceneRenderer {
      Insert()
    }
    class PowerPointSceneSurface {
      ReadVerified()
    }
    RenderResult *-- SceneNode : semanticScene
    SceneNode *-- SceneGeometry
    SceneNode *-- ScenePaint
    SceneGeometry <|-- BoxGeometry
    BoxGeometry <|-- TextGeometry
    SceneGeometry <|-- PathGeometry
    SceneGeometry <|-- SectorGeometry
    SceneCodec ..> RenderResult
    NativeSceneRenderer --> IEditableSceneSurface
    IEditableSceneSurface <|.. PowerPointSceneSurface
```

Evidence: `src/Chart.Renderers/SceneContracts.cs:7`; `src/Chart.Renderers/NativeRenderingCore.cs:136`; `src/Chart.Renderers/PowerPointSceneSurface.cs:15`. PathGeometry carries ordered points; sector angles/radius are bounded. The interface does not authorize arbitrary scene execution. ClusteredProjection validates a narrow canonical rectangle profile before Begin. Native adapter does not implement text/polygon/sector rendering despite their DTO presence.

## Catalog extraction and query objects

```mermaid
classDiagram
    class PiidxCatalog {
      string IndexPath
      Register()
      Refresh()
      Search()
      Slide()
      VerifySource()
      PutPreview()
      Preview()
      VerifyIntegrity()
    }
    class SqliteDatabase {
      Open()
      Execute()
      Query()
      ObserveCancellation()
      Dispose()
    }
    class PresentationSource {
      Discover()
      Fingerprint()
      Read()
    }
    class IndexedPresentation {
      string SourcePath
      string ContentHash
      double Width
      double Height
    }
    class IndexedSlide {
      uint NativeId
      int Position
      string PartUri
      string ContentFingerprint
    }
    class IndexedObject {
      string Key
      string ParentKey
      string NativeId
      string XmlFingerprint
    }
    class LibrarySearchRequest {
      string Query
      int Offset
      int Limit
    }
    class LibrarySearchPage {
      long Total
    }
    class LibrarySlideReference {
      string SlideId
      string SourcePath
      string SourceFingerprint
      string SlideFingerprint
    }
    PiidxCatalog ..> SqliteDatabase
    PiidxCatalog ..> PresentationSource
    PresentationSource ..> IndexedPresentation
    IndexedPresentation *-- IndexedSlide
    IndexedSlide *-- IndexedObject
    PiidxCatalog ..> LibrarySearchRequest
    PiidxCatalog ..> LibrarySearchPage
    LibrarySearchPage *-- LibrarySlideReference
```

Evidence: `src/Library.Core/PiidxCatalog.cs:14`; `src/Library.Core/PresentationSource.cs:17`; `src/Library.Core/LibraryModels.cs:55`. PiidxCatalog is one partial class across catalog/search/navigation/preview source files, not four independent databases. Extracted object transforms remain native parent-coordinate metadata; group parent references are structural strings rather than relational FKs.

## Storage ERD

See [persistence design](persistence-design.md) for actual SQL relationships. No application-server ERD is provided because the inspected implementation contains no such database. Native PowerPoint objects/tags/XML are host persistence, not fictitious relational tables.
