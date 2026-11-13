# Data dictionary

<!-- generated-toc -->
## Contents

- [JSON opt-in fields](#json-opt-in-fields)
- [PIIDX v1 SQL columns and constraints](#piidx-v1-sql-columns-and-constraints)
- [Library request and result contract](#library-request-and-result-contract)
- [Host metadata adjuncts](#host-metadata-adjuncts)
- [Data sensitivity](#data-sensitivity)
<!-- end-generated-toc -->

**VERIFIED:** inspected implementation, not a claim of current Office execution. The repository contains several independent versioned models; version numbers across these formats are not interchangeable.

## JSON opt-in fields

| Owner | Wire field | CLR type | Presence rule | Evidence |
| --- | --- | --- | --- | --- |
| ChartDocument | schemaVersion | int | Required.Always | `src/Chart.Domain/ChartDocument.cs:10` |
| ChartDocument | chartId | string | Required.Always | `src/Chart.Domain/ChartDocument.cs:11` |
| ChartDocument | revision | int | Required.Always | `src/Chart.Domain/ChartDocument.cs:12` |
| ChartDocument | chartType | string | Required.Always | `src/Chart.Domain/ChartDocument.cs:13` |
| ChartDocument | dataset | EmbeddedDataset | Required.Always | `src/Chart.Domain/ChartDocument.cs:14` |
| ChartDocument | series | List<ChartSeries> | Required.Always | `src/Chart.Domain/ChartDocument.cs:15` |
| ChartDocument | axes | List<ChartAxis> | Required.Always | `src/Chart.Domain/ChartDocument.cs:16` |
| ChartDocument | layout | ChartLayout | Required.Always | `src/Chart.Domain/ChartDocument.cs:17` |
| ChartDocument | style | ChartStyle | Required.Always | `src/Chart.Domain/ChartDocument.cs:18` |
| EmbeddedDataset | id | string | Required.Always | `src/Chart.Domain/ChartDocument.cs:23` |
| EmbeddedDataset | categories | List<ChartCategory> | Required.Always | `src/Chart.Domain/ChartDocument.cs:24` |
| ChartCategory | id | string | Required.Always | `src/Chart.Domain/ChartDocument.cs:29` |
| ChartCategory | label | string | Required.Always | `src/Chart.Domain/ChartDocument.cs:30` |
| ChartSeries | id | string | Required.Always | `src/Chart.Domain/ChartDocument.cs:35` |
| ChartSeries | name | string | Required.Always | `src/Chart.Domain/ChartDocument.cs:36` |
| ChartSeries | values | List<double?> | Required.Always | `src/Chart.Domain/ChartDocument.cs:37` |
| ChartAxis | id | string | Required.Always | `src/Chart.Domain/ChartDocument.cs:42` |
| ChartAxis | role | string | Required.Always | `src/Chart.Domain/ChartDocument.cs:43` |
| ChartAxis | scale | string | Required.Always | `src/Chart.Domain/ChartDocument.cs:44` |
| ChartAxis | numberFormat | string | Required.Always | `src/Chart.Domain/ChartDocument.cs:45` |
| ChartLayout | unit | string | Required.Always | `src/Chart.Domain/ChartDocument.cs:50` |
| ChartLayout | x | double | Required.Always | `src/Chart.Domain/ChartDocument.cs:51` |
| ChartLayout | y | double | Required.Always | `src/Chart.Domain/ChartDocument.cs:52` |
| ChartLayout | width | double | Required.Always | `src/Chart.Domain/ChartDocument.cs:53` |
| ChartLayout | height | double | Required.Always | `src/Chart.Domain/ChartDocument.cs:54` |
| ChartStyle | showLegend | bool | Required.Always | `src/Chart.Domain/ChartDocument.cs:59` |
| ChartStyle | seriesColors | List<string> | Required.Always | `src/Chart.Domain/ChartDocument.cs:60` |
| ChartDocumentV3 | schemaVersion | int | Required.Always | `src/Chart.Domain/ChartDocumentV3.cs:8` |
| ChartDocumentV3 | chartId | string | Required.Always | `src/Chart.Domain/ChartDocumentV3.cs:9` |
| ChartDocumentV3 | revision | int | Required.Always | `src/Chart.Domain/ChartDocumentV3.cs:10` |
| ChartDocumentV3 | chartType | string | Required.Always | `src/Chart.Domain/ChartDocumentV3.cs:11` |
| ChartDocumentV3 | dataset | EmbeddedDataset | Required.Always | `src/Chart.Domain/ChartDocumentV3.cs:12` |
| ChartDocumentV3 | series | List<ChartSeries> | Required.Always | `src/Chart.Domain/ChartDocumentV3.cs:13` |
| ChartDocumentV3 | axes | List<ChartAxis> | Required.Always | `src/Chart.Domain/ChartDocumentV3.cs:14` |
| ChartDocumentV3 | layout | ChartLayout | Required.Always | `src/Chart.Domain/ChartDocumentV3.cs:15` |
| ChartDocumentV3 | style | ChartStyleV3 | Required.Always | `src/Chart.Domain/ChartDocumentV3.cs:16` |
| ChartDocumentV3 | options | ChartOptions | Required.Always | `src/Chart.Domain/ChartDocumentV3.cs:17` |
| ChartStyleV3 | showLegend | bool | Required.Always | `src/Chart.Domain/ChartDocumentV3.cs:22` |
| ChartStyleV3 | seriesColors | List<string> | Required.Always | `src/Chart.Domain/ChartDocumentV3.cs:23` |
| ChartStyleV3 | categoryColors | List<string> | Required.Always | `src/Chart.Domain/ChartDocumentV3.cs:24` |
| ChartOptions | kind | string | Required.Always | `src/Chart.Domain/ChartDocumentV3.cs:28` |
| BarOptions | gapRatio | double | Required.Always | `src/Chart.Domain/ChartDocumentV3.cs:31` |
| LineOptions | gapPolicy | string | Required.Always | `src/Chart.Domain/ChartDocumentV3.cs:35` |
| LineOptions | includeZero | bool | Required.Always | `src/Chart.Domain/ChartDocumentV3.cs:36` |
| AreaOptions | gapPolicy | string | Required.Always | `src/Chart.Domain/ChartDocumentV3.cs:41` |
| AreaOptions | baseline | double | Required.Always | `src/Chart.Domain/ChartDocumentV3.cs:42` |
| PieOptions | startAngle | double | Required.Always | `src/Chart.Domain/ChartDocumentV3.cs:46` |
| RenderResult | schemaVersion | int | Required.Always | `src/Chart.Renderers/SceneContracts.cs:9` |
| RenderResult | chartId | string | Required.Always | `src/Chart.Renderers/SceneContracts.cs:10` |
| RenderResult | modelRevision | int | Required.Always | `src/Chart.Renderers/SceneContracts.cs:11` |
| RenderResult | unit | string | Required.Always | `src/Chart.Renderers/SceneContracts.cs:12` |
| RenderResult | plot | ScenePlot | Required.Always | `src/Chart.Renderers/SceneContracts.cs:13` |
| RenderResult | semanticScene | List<SceneNode> | Required.Always | `src/Chart.Renderers/SceneContracts.cs:14` |
| RenderResult | diagnostics | List<SceneDiagnostic> | Required.Always | `src/Chart.Renderers/SceneContracts.cs:15` |
| RenderResult | renderHash | string | Required.AllowNull | `src/Chart.Renderers/SceneContracts.cs:18` |
| ScenePlot | x | double | Required.Always | `src/Chart.Renderers/SceneContracts.cs:23` |
| ScenePlot | y | double | Required.Always | `src/Chart.Renderers/SceneContracts.cs:24` |
| ScenePlot | width | double | Required.Always | `src/Chart.Renderers/SceneContracts.cs:25` |
| ScenePlot | height | double | Required.Always | `src/Chart.Renderers/SceneContracts.cs:26` |
| ScenePlot | unit | string | Required.Always | `src/Chart.Renderers/SceneContracts.cs:27` |
| SceneNode | id | string | Required.Always | `src/Chart.Renderers/SceneContracts.cs:32` |
| SceneNode | geometry | SceneGeometry | Required.Always | `src/Chart.Renderers/SceneContracts.cs:33` |
| SceneNode | paint | ScenePaint | Required.Always | `src/Chart.Renderers/SceneContracts.cs:34` |
| ScenePaint | fill | string | Required.AllowNull | `src/Chart.Renderers/SceneContracts.cs:39` |
| ScenePaint | stroke | string | Required.AllowNull | `src/Chart.Renderers/SceneContracts.cs:40` |
| ScenePaint | strokeWidth | double | Required.Always | `src/Chart.Renderers/SceneContracts.cs:41` |
| ScenePaint | opacity | double | Required.Always | `src/Chart.Renderers/SceneContracts.cs:42` |
| SceneDiagnostic | code | string | Required.Always | `src/Chart.Renderers/SceneContracts.cs:47` |
| SceneDiagnostic | nodeId | string | Required.AllowNull | `src/Chart.Renderers/SceneContracts.cs:48` |
| SceneGeometry | type | string | Required.Always | `src/Chart.Renderers/SceneContracts.cs:52` |
| BoxGeometry | x | double | Required.Always | `src/Chart.Renderers/SceneContracts.cs:56` |
| BoxGeometry | y | double | Required.Always | `src/Chart.Renderers/SceneContracts.cs:57` |
| BoxGeometry | width | double | Required.Always | `src/Chart.Renderers/SceneContracts.cs:58` |
| BoxGeometry | height | double | Required.Always | `src/Chart.Renderers/SceneContracts.cs:59` |
| TextGeometry | text | string | Required.Always | `src/Chart.Renderers/SceneContracts.cs:64` |
| TextGeometry | fontFamily | string | Required.Always | `src/Chart.Renderers/SceneContracts.cs:65` |
| TextGeometry | fontSize | double | Required.Always | `src/Chart.Renderers/SceneContracts.cs:66` |
| TextGeometry | alignment | string | Required.Always | `src/Chart.Renderers/SceneContracts.cs:67` |
| PathGeometry | points | List<ScenePoint> | Required.Always | `src/Chart.Renderers/SceneContracts.cs:71` |
| ScenePoint | x | double | Required.Always | `src/Chart.Renderers/SceneContracts.cs:75` |
| ScenePoint | y | double | Required.Always | `src/Chart.Renderers/SceneContracts.cs:76` |
| SectorGeometry | cx | double | Required.Always | `src/Chart.Renderers/SceneContracts.cs:81` |
| SectorGeometry | cy | double | Required.Always | `src/Chart.Renderers/SceneContracts.cs:82` |
| SectorGeometry | radius | double | Required.Always | `src/Chart.Renderers/SceneContracts.cs:83` |
| SectorGeometry | startAngle | double | Required.Always | `src/Chart.Renderers/SceneContracts.cs:84` |
| SectorGeometry | sweep | double | Required.Always | `src/Chart.Renderers/SceneContracts.cs:85` |

## PIIDX v1 SQL columns and constraints

Nine ordinary tables plus one FTS5 virtual table are logical entities. SQLite creates FTS shadow tables; these are not invented application entities. Table-constraint rows preserve composite keys/uniqueness.

| Entity | Column | Declared type/constraints | Evidence |
| --- | --- | --- | --- |
| INDEX_METADATA | singleton | INTEGER PRIMARY KEY CHECK(singleton=1) | `src/Library.Core/PiidxSchema.cs:15` |
| INDEX_METADATA | format_id | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:15` |
| INDEX_METADATA | schema_version | INTEGER NOT NULL | `src/Library.Core/PiidxSchema.cs:15` |
| INDEX_METADATA | engine_version | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:15` |
| INDEX_METADATA | created_utc | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:15` |
| INDEX_METADATA | last_successful_update | TEXT | `src/Library.Core/PiidxSchema.cs:15` |
| INDEX_METADATA | library_id | TEXT | `src/Library.Core/PiidxSchema.cs:15` |
| INDEX_METADATA | source_fingerprint | TEXT | `src/Library.Core/PiidxSchema.cs:15` |
| INDEX_METADATA | integrity_digest | TEXT | `src/Library.Core/PiidxSchema.cs:15` |
| LIBRARIES | library_id | TEXT PRIMARY KEY NOT NULL | `src/Library.Core/PiidxSchema.cs:16` |
| LIBRARIES | root_path | TEXT NOT NULL UNIQUE | `src/Library.Core/PiidxSchema.cs:16` |
| LIBRARIES | name | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:16` |
| LIBRARIES | recursive | INTEGER NOT NULL CHECK(recursive IN(0,1)) | `src/Library.Core/PiidxSchema.cs:16` |
| LIBRARIES | registered_utc | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:16` |
| LIBRARIES | last_refresh_utc | TEXT | `src/Library.Core/PiidxSchema.cs:16` |
| LIBRARIES | enabled | INTEGER NOT NULL DEFAULT 1 CHECK(enabled IN(0,1)) | `src/Library.Core/PiidxSchema.cs:16` |
| PRESENTATIONS | presentation_id | TEXT PRIMARY KEY NOT NULL | `src/Library.Core/PiidxSchema.cs:17` |
| PRESENTATIONS | library_id | TEXT NOT NULL REFERENCES LIBRARIES(library_id) ON DELETE CASCADE | `src/Library.Core/PiidxSchema.cs:17` |
| PRESENTATIONS | source_path | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:17` |
| PRESENTATIONS | relative_path | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:17` |
| PRESENTATIONS | file_type | TEXT NOT NULL CHECK(file_type IN('.pptx','.potx')) | `src/Library.Core/PiidxSchema.cs:17` |
| PRESENTATIONS | file_size | INTEGER NOT NULL CHECK(file_size>=0) | `src/Library.Core/PiidxSchema.cs:17` |
| PRESENTATIONS | modified_utc | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:17` |
| PRESENTATIONS | content_hash | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:17` |
| PRESENTATIONS | index_status | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:17` |
| PRESENTATIONS | slide_count | INTEGER NOT NULL CHECK(slide_count>=0) | `src/Library.Core/PiidxSchema.cs:17` |
| PRESENTATIONS | [table constraint] | UNIQUE(library_id,relative_path) | `src/Library.Core/PiidxSchema.cs:17` |
| SLIDES | slide_id | TEXT PRIMARY KEY NOT NULL | `src/Library.Core/PiidxSchema.cs:18` |
| SLIDES | presentation_id | TEXT NOT NULL REFERENCES PRESENTATIONS(presentation_id) ON DELETE CASCADE | `src/Library.Core/PiidxSchema.cs:18` |
| SLIDES | position | INTEGER NOT NULL CHECK(position>0) | `src/Library.Core/PiidxSchema.cs:18` |
| SLIDES | native_slide_id | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:18` |
| SLIDES | part_uri | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:18` |
| SLIDES | title | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:18` |
| SLIDES | text | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:18` |
| SLIDES | keywords | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:18` |
| SLIDES | category | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:18` |
| SLIDES | author | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:18` |
| SLIDES | layout | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:18` |
| SLIDES | master | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:18` |
| SLIDES | width | REAL NOT NULL CHECK(width>0) | `src/Library.Core/PiidxSchema.cs:18` |
| SLIDES | height | REAL NOT NULL CHECK(height>0) | `src/Library.Core/PiidxSchema.cs:18` |
| SLIDES | content_fingerprint | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:18` |
| SLIDES | [table constraint] | UNIQUE(presentation_id,native_slide_id) | `src/Library.Core/PiidxSchema.cs:18` |
| SLIDES | [table constraint] | UNIQUE(presentation_id,position) | `src/Library.Core/PiidxSchema.cs:18` |
| SLIDE_OBJECTS | slide_id | TEXT NOT NULL REFERENCES SLIDES(slide_id) ON DELETE CASCADE | `src/Library.Core/PiidxSchema.cs:19` |
| SLIDE_OBJECTS | object_id | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:19` |
| SLIDE_OBJECTS | kind | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:19` |
| SLIDE_OBJECTS | name | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:19` |
| SLIDE_OBJECTS | text | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:19` |
| SLIDE_OBJECTS | bounds_json | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:19` |
| SLIDE_OBJECTS | metadata_json | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:19` |
| SLIDE_OBJECTS | [table constraint] | PRIMARY KEY(slide_id,object_id) | `src/Library.Core/PiidxSchema.cs:19` |
| SLIDE_TAGS | slide_id | TEXT NOT NULL REFERENCES SLIDES(slide_id) ON DELETE CASCADE | `src/Library.Core/PiidxSchema.cs:20` |
| SLIDE_TAGS | tag | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:20` |
| SLIDE_TAGS | normalized_tag | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:20` |
| SLIDE_TAGS | [table constraint] | PRIMARY KEY(slide_id,normalized_tag) | `src/Library.Core/PiidxSchema.cs:20` |
| SLIDE_PREVIEWS | slide_id | TEXT NOT NULL REFERENCES SLIDES(slide_id) ON DELETE CASCADE | `src/Library.Core/PiidxSchema.cs:21` |
| SLIDE_PREVIEWS | width | INTEGER NOT NULL CHECK(width>0 AND width<=4096) | `src/Library.Core/PiidxSchema.cs:21` |
| SLIDE_PREVIEWS | height | INTEGER NOT NULL CHECK(height>0 AND height<=4096) | `src/Library.Core/PiidxSchema.cs:21` |
| SLIDE_PREVIEWS | mime | TEXT NOT NULL CHECK(mime='image/png') | `src/Library.Core/PiidxSchema.cs:21` |
| SLIDE_PREVIEWS | image | BLOB NOT NULL | `src/Library.Core/PiidxSchema.cs:21` |
| SLIDE_PREVIEWS | content_hash | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:21` |
| SLIDE_PREVIEWS | source_fingerprint | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:21` |
| SLIDE_PREVIEWS | rendered_utc | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:21` |
| SLIDE_PREVIEWS | [table constraint] | PRIMARY KEY(slide_id,width,height) | `src/Library.Core/PiidxSchema.cs:21` |
| SLIDE_SEARCH | slide_id | UNINDEXED | `src/Library.Core/PiidxSchema.cs:22` |
| SLIDE_SEARCH | title | FTS text | `src/Library.Core/PiidxSchema.cs:22` |
| SLIDE_SEARCH | text | FTS text | `src/Library.Core/PiidxSchema.cs:22` |
| SLIDE_SEARCH | keywords | FTS text | `src/Library.Core/PiidxSchema.cs:22` |
| INDEX_JOBS | job_id | TEXT PRIMARY KEY NOT NULL | `src/Library.Core/PiidxSchema.cs:23` |
| INDEX_JOBS | library_id | TEXT REFERENCES LIBRARIES(library_id) ON DELETE SET NULL | `src/Library.Core/PiidxSchema.cs:23` |
| INDEX_JOBS | started_utc | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:23` |
| INDEX_JOBS | completed_utc | TEXT | `src/Library.Core/PiidxSchema.cs:23` |
| INDEX_JOBS | state | TEXT NOT NULL CHECK(state IN('running','completed','cancelled','failed','interrupted')) | `src/Library.Core/PiidxSchema.cs:23` |
| INDEX_JOBS | files_seen | INTEGER NOT NULL DEFAULT 0 | `src/Library.Core/PiidxSchema.cs:23` |
| INDEX_JOBS | files_updated | INTEGER NOT NULL DEFAULT 0 | `src/Library.Core/PiidxSchema.cs:23` |
| INDEX_JOBS | slides_indexed | INTEGER NOT NULL DEFAULT 0 | `src/Library.Core/PiidxSchema.cs:23` |
| INDEX_ERRORS | error_id | INTEGER PRIMARY KEY | `src/Library.Core/PiidxSchema.cs:24` |
| INDEX_ERRORS | job_id | TEXT REFERENCES INDEX_JOBS(job_id) ON DELETE CASCADE | `src/Library.Core/PiidxSchema.cs:24` |
| INDEX_ERRORS | source_path | TEXT | `src/Library.Core/PiidxSchema.cs:24` |
| INDEX_ERRORS | stage | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:24` |
| INDEX_ERRORS | error_code | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:24` |
| INDEX_ERRORS | message | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:24` |
| INDEX_ERRORS | recorded_utc | TEXT NOT NULL | `src/Library.Core/PiidxSchema.cs:24` |

## Library request and result contract

LibrarySearchRequest fields: Query (default empty), LibraryId, PresentationId, Category, CategoryRecursive, HideDuplicates, ModifiedSinceUtc, Aspect, Sort, Tag, Author, Layout, Mode(defaultAllWords), Offset and Limit(default100). Search bounds query<=4096, limit1..1000, offset0..10000000 and valid enum values. Each result supplies stable catalog IDs, source path/hash, slide hash/nativeID, title/text/category/author/layout/master, current position, width/height, rank and modifiedUTC. SearchPage has Total and read-only slide list. Search is parameterized lexical/metadata retrieval, not semantic AI search.

## Host metadata adjuncts

DataSourceLink stores ID/format/path/link time and optional ExcelRangeMapping; mapping stores Worksheet/Address, zero-based Category, distinct Series indices and original Headers. Range constraints: single A1 rectangle,2–50 rows including headers and2–15 columns within Excel limits; headers distinct/nonempty<=200 chars. Header or dimension drift stops refresh before chart mutation. See `apps/powerpoint-vsto/ExcelRangeMapping.cs:14` and `apps/powerpoint-vsto/DataSources.cs:254`.

Chapter records: SlideId/SlideIndex/Level/Scope/Title; caption baseline keyed by stable slideID. At most200 chapters, levels1..4 without hierarchy jumps, captions1..240 printable characters. Process grids:1..20rows,1..12columns,<=1000 retained cells, geometry fits slide, active cells form unique full rectangle and min48x28pt. Indicators: kind harvey/checkbox, segments1..100, filled0..segments, checkbox state0..2, size6..720pt, rotation[0,360), color0..0xFFFFFF and customsymbol<=8 chars. Connector endpoints whitelist site/point/free and center/top/right/bottom/left anchors; route straight/elbow/curve, weight0.25..12pt. These constraints belong to their codecs/planners, not PIIDX tables.

## Data sensitivity

Source paths, author/text/tag content, workbook-link paths and template content can be business-sensitive. Source/preview fingerprints are integrity consistency markers, not anonymization or encryption. Runtime readers use current-user filesystem permissions; no RBAC/tenant enforcement is implied. Examples in this export are synthetic and never contain credentials or real workbook values.
