# Data-flow and trust-boundary diagrams

```mermaid
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

There is no verified data flow from presentation content to Azure OpenAI. Do not draw or document an authenticated AI request sequence until the gateway/client actually exists. Local configured-provider profiles are shown separately in [AI data flow](../06-ai-architecture/ai-data-flow.md).
