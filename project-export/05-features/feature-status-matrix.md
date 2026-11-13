# Feature status matrix

Evidence baseline: current inspected checkout, HEAD `ffaec352cbdca666f69f3c402d334108e4115933` plus parent export hash inventory. **VERIFIED** means source implementation here; current native runtime acceptance is not implied. Source analysis levels and limitations are in `04-powerpoint-integration/vsto-source-review.csv`.

## Confirmed source subsets versus absent capabilities

| Area | Source status | Boundary / evidence |
|---|---|---|
| Native linked charts | Implemented subset | Column/bar/stacked/percent/line/profile/area/pie/doughnut/scatter/bubble/combo through native AddChart2; mapped refresh has a narrower allowlist. `apps/powerpoint-vsto/DemoTools.cs:165`, `apps/powerpoint-vsto/MappedLinkOperations.cs:14`. |
| Waterfall/Mekko/Gantt | Implemented local editable shapes | Grouped native shapes with linked table snapshot; not full reference-product semantic engine. `apps/powerpoint-vsto/DemoTools.cs:197`. |
| Linked scene renderer | Partial native shape projection | Rectangle columns; axes/legend omitted with diagnostics. `apps/powerpoint-vsto/RendererRibbonCommand.cs:44`. |
| Native tables | Implemented bounded subset | Linked headers/text/values, <=400 cells, datasheet editor; mapped refresh not implemented. `apps/powerpoint-vsto/DemoTools.cs:252`. |
| Text/process/connectors/indicators | Implemented managed subsets | Explicit owned metadata/geometry constraints; foreign/invalid edits fail preserved. See individual module designs. |
| Context properties | Implemented bounded source | Text/shape/table/chart/connector/indicator/grid/template/agenda/image/mixed classifications; native foreground/DPI/family acceptance requires desktop validation. `apps/powerpoint-vsto/ContextPanelModel.cs:10`. |
| Agendas | Implemented managed source | Schema2 chapters/hierarchy/contents/scopes/moves/native captions plus legacy schema1 derived title/section agenda. `apps/powerpoint-vsto/ChapterAgenda.cs:70`. |
| Designed library | Implemented generated specification | Deterministic catalog, editable fields, palettes/alignment, native optional builds/transitions; business uniqueness is not proven by count. `apps/powerpoint-vsto/TemplateCatalog.cs:129`. |
| Source-original library | Implemented local Office adapter | Indexed metadata plus actual native preview/keep-source insertion and strict source identity/size checks. `apps/powerpoint-vsto/TemplateLibrary.cs:462`. |
| AI / Azure / external services | Configuration only; live calls NOT IMPLEMENTED | Three nonsecret endpoint/auth-intent profiles; no client/credentials/token acquisition/connectivity test. `apps/powerpoint-vsto/ServiceConnections.cs:29`. |
| On-prem API/RBAC | NOT IMPLEMENTED as running service in this graph | A requirements/schema directory is not an executable deployed backend; see AI/API chapters. |
| SmartArt / Office JS add-in | NOT IMPLEMENTED in compiled graph | VSTO native Shapes/Charts/Tables; csproj references no Office.js web host. `apps/powerpoint-vsto/PowerPointIntelligence.csproj:26`. |
| Shared slide workbook | NOT IMPLEMENTED | Chart list opens separate native workbooks. `apps/powerpoint-vsto/ProductivityTools.cs:95`. |
| Automatic shape/chart recognition from pictures | NOT IMPLEMENTED | Scanner counts pictures/charts; it performs no OCR or chart reconstruction. DemoTools catalog scan entry. |
| External data-driven indicators | NOT IMPLEMENTED | Local direct SetValueSelected helper does not implement data linkage. `apps/powerpoint-vsto/IndicatorElements.cs:37`. |
| Full accessibility/brand certification | NOT IMPLEMENTED | Read-only bounded heuristics exclude multiple content classes. `apps/powerpoint-vsto/PresentationQuality.cs:49`. |

## Per-route ledger

| Feature | Status |
|---|---|
| CMD-DEMO-demo-tour | Implemented source; runtime unverified |
| CMD-DEMO-column | Implemented source; runtime unverified |
| CMD-DEMO-bar | Implemented source; runtime unverified |
| CMD-DEMO-stacked-column | Implemented source; runtime unverified |
| CMD-DEMO-percent-column | Implemented source; runtime unverified |
| CMD-DEMO-stacked-bar | Implemented source; runtime unverified |
| CMD-DEMO-percent-bar | Implemented source; runtime unverified |
| CMD-DEMO-line | Implemented source; runtime unverified |
| CMD-DEMO-profile | Implemented source; runtime unverified |
| CMD-DEMO-area | Implemented source; runtime unverified |
| CMD-DEMO-pie | Implemented source; runtime unverified |
| CMD-DEMO-doughnut | Implemented source; runtime unverified |
| CMD-DEMO-scatter | Implemented source; runtime unverified |
| CMD-DEMO-bubble | Implemented source; runtime unverified |
| CMD-DEMO-combo | Implemented source; runtime unverified |
| CMD-DEMO-waterfall | Implemented source; runtime unverified |
| CMD-DEMO-mekko | Implemented source; runtime unverified |
| CMD-DEMO-mekko-units | Implemented source; runtime unverified |
| CMD-DEMO-gantt | Implemented source; runtime unverified |
| CMD-DEMO-text | Implemented source; runtime unverified |
| CMD-DEMO-rounded | Implemented source; runtime unverified |
| CMD-DEMO-process | Implemented source; runtime unverified |
| CMD-DEMO-checkbox | Implemented source; runtime unverified |
| CMD-DEMO-harvey | Implemented source; runtime unverified |
| CMD-DEMO-table | Implemented source; runtime unverified |
| CMD-DEMO-agenda | Implemented source; runtime unverified |
| CMD-DEMO-image | Implemented source; runtime unverified |
| CMD-DEMO-symbol | Implemented source; runtime unverified |
| CMD-DEMO-auto-layout | Implemented source; runtime unverified |
| CMD-DEMO-paste-text | Implemented source; runtime unverified |
| CMD-DEMO-font | Implemented source; runtime unverified |
| CMD-DEMO-font-size | Implemented source; runtime unverified |
| CMD-DEMO-decimal | Implemented source; runtime unverified |
| CMD-DEMO-language | Implemented source; runtime unverified |
| CMD-DEMO-edit-data | Implemented source; runtime unverified |
| CMD-DEMO-legend | Implemented source; runtime unverified |
| CMD-DEMO-labels | Implemented source; runtime unverified |
| CMD-DEMO-data-table | Implemented source; runtime unverified |
| CMD-DEMO-chart-style | Implemented source; runtime unverified |
| CMD-DEMO-value-line | Implemented source; runtime unverified |
| CMD-DEMO-difference | Implemented source; runtime unverified |
| CMD-DEMO-cagr | Implemented source; runtime unverified |
| CMD-DEMO-scan | Implemented source; runtime unverified |
| CMD-DEMO-path | Implemented source; runtime unverified |
| CMD-DEMO-library | Implemented source; runtime unverified |
| CMD-DEMO-template | Implemented source; runtime unverified |
| CMD-DEMO-export-slide | Implemented source; runtime unverified |
| CMD-DEMO-theme | Implemented source; runtime unverified |
| CMD-DEMO-axes | Implemented source; runtime unverified |
| CMD-DEMO-plot-area | Implemented source; runtime unverified |
| CMD-DEMO-chart-title | Implemented source; runtime unverified |
| CMD-DEMO-duplicate-slide | Implemented source; runtime unverified |
| CMD-DEMO-save-slides | Implemented source; runtime unverified |
| CMD-DEMO-clean-comments | Implemented source; runtime unverified |
| CMD-DEMO-round-data | Implemented source; runtime unverified |
| CMD-DEMO-sheet-style | Implemented source; runtime unverified |
| CMD-DEMO-import-json | Implemented source; runtime unverified |
| CMD-DEMO-export-json | Implemented source; runtime unverified |
| CMD-DEMO-mini-toolbar | Implemented source; runtime unverified |
| CMD-DEMO-table-data | Implemented source; runtime unverified |
| CMD-DEMO-library-browser | Implemented source; runtime unverified |
| CMD-DEMO-agenda-layout | Implemented source; runtime unverified |
| CMD-DEMO-style-file | Implemented source; runtime unverified |
| CMD-DEMO-excel-cagr | Implemented source; runtime unverified |
| CMD-DEMO-link-excel | Implemented source; runtime unverified |
| CMD-DEMO-refresh-excel | Implemented source; runtime unverified |
| CMD-DEMO-manage-links | Implemented source; runtime unverified |
| CMD-DEMO-slide-workbook | Implemented source; runtime unverified |
| CMD-DEMO-left | Implemented source; runtime unverified |
| CMD-DEMO-center | Implemented source; runtime unverified |
| CMD-DEMO-right | Implemented source; runtime unverified |
| CMD-DEMO-top | Implemented source; runtime unverified |
| CMD-DEMO-middle | Implemented source; runtime unverified |
| CMD-DEMO-bottom | Implemented source; runtime unverified |
| CMD-DEMO-width | Implemented source; runtime unverified |
| CMD-DEMO-height | Implemented source; runtime unverified |
| CMD-DEMO-size | Implemented source; runtime unverified |
| CMD-DEMO-horizontal | Implemented source; runtime unverified |
| CMD-DEMO-vertical | Implemented source; runtime unverified |
| CMD-DEMO-swap | Implemented source; runtime unverified |
| CMD-DEMO-scale-up | Implemented source; runtime unverified |
| CMD-DEMO-scale-down | Implemented source; runtime unverified |
| CMD-DEMO-rotate | Implemented source; runtime unverified |
| CMD-DEMO-flip-horizontal | Implemented source; runtime unverified |
| CMD-DEMO-flip-vertical | Implemented source; runtime unverified |
| CMD-DEMO-duplicate | Implemented source; runtime unverified |
| CMD-DEMO-group | Implemented source; runtime unverified |
| CMD-DEMO-ungroup | Implemented source; runtime unverified |
| CMD-DEMO-front | Implemented source; runtime unverified |
| CMD-DEMO-back | Implemented source; runtime unverified |
| CMD-ELEMENT-pentagon | Implemented source; runtime unverified |
| CMD-ELEMENT-chevron | Implemented source; runtime unverified |
| CMD-ELEMENT-connector | Implemented source; runtime unverified |
| CMD-ELEMENT-edit | Implemented source; runtime unverified |
| CMD-ELEMENT-fit | Implemented source; runtime unverified |
| CMD-ELEMENT-convert-text | Implemented source; runtime unverified |
| CMD-ELEMENT-convert-pentagon | Implemented source; runtime unverified |
| CMD-ELEMENT-convert-chevron | Implemented source; runtime unverified |
| CMD-ELEMENT-convert-rounded | Implemented source; runtime unverified |
| CMD-ELEMENT-duplicate-left | Implemented source; runtime unverified |
| CMD-ELEMENT-duplicate-right | Implemented source; runtime unverified |
| CMD-ELEMENT-duplicate-up | Implemented source; runtime unverified |
| CMD-ELEMENT-duplicate-down | Implemented source; runtime unverified |
| CMD-ELEMENT-indicator-next | Implemented source; runtime unverified |
| CMD-ELEMENT-indicator-previous | Implemented source; runtime unverified |
| CMD-ELEMENT-indicator-keys | Implemented source; runtime unverified |
| CMD-ELEMENT-convert-harvey | Implemented source; runtime unverified |
| CMD-ELEMENT-convert-checkbox | Implemented source; runtime unverified |
| CMD-ELEMENT-grid-create | Implemented source; runtime unverified |
| CMD-ELEMENT-grid-row | Implemented source; runtime unverified |
| CMD-ELEMENT-grid-column | Implemented source; runtime unverified |
| CMD-ELEMENT-grid-remove-row | Implemented source; runtime unverified |
| CMD-ELEMENT-grid-remove-column | Implemented source; runtime unverified |
| CMD-ELEMENT-grid-left | Implemented source; runtime unverified |
| CMD-ELEMENT-grid-right | Implemented source; runtime unverified |
| CMD-ELEMENT-grid-stripe-rows | Implemented source; runtime unverified |
| CMD-ELEMENT-grid-stripe-columns | Implemented source; runtime unverified |
| CMD-ELEMENT-grid-resize | Implemented source; runtime unverified |
| CMD-ELEMENT-refresh | Implemented source; runtime unverified |
| CMD-TOOLS-symbol-picker | Implemented source; runtime unverified |
| CMD-TOOLS-decimal-safe | Implemented source; runtime unverified |
| CMD-TOOLS-raster-copy | Implemented source; runtime unverified |
| CMD-TOOLS-chart-workbooks | Implemented source; runtime unverified |
| CMD-TOOLS-data-links | Implemented source; runtime unverified |
| CMD-TOOLS-save-selected | Implemented source; runtime unverified |
| CMD-TOOLS-selected-style | Implemented source; runtime unverified |
| CMD-TOOLS-review | Implemented source; runtime unverified |
| CMD-TOOLS-help | Implemented source; runtime unverified |
| CMD-WORKFLOW-agenda-insert | Implemented source; runtime unverified |
| CMD-WORKFLOW-agenda-contents | Implemented source; runtime unverified |
| CMD-WORKFLOW-agenda-settings | Implemented source; runtime unverified |
| CMD-WORKFLOW-agenda-split | Implemented source; runtime unverified |
| CMD-WORKFLOW-agenda-refresh | Implemented source; runtime unverified |
| CMD-WORKFLOW-percent-area | Implemented source; runtime unverified |
| CMD-WORKFLOW-series-settings | Implemented source; runtime unverified |
| CMD-WORKFLOW-axis-settings | Implemented source; runtime unverified |
| CMD-WORKFLOW-match-scales | Implemented source; runtime unverified |
| CMD-WORKFLOW-annotation-refresh | Implemented source; runtime unverified |
| CMD-WORKFLOW-connections | Partial: nonsecret configuration only; no live service execution |
| CMD-WORKFLOW-relink-selected | Implemented source; runtime unverified |
| CMD-WORKFLOW-map-excel | Implemented source; runtime unverified |
| CMD-WORKFLOW-data-manager | Implemented source; runtime unverified |
| CMD-WORKFLOW-quality-review | Implemented source; runtime unverified |
| CMD-WORKFLOW-export-pdf | Implemented source; runtime unverified |
| CMD-WORKFLOW-text-left | Implemented source; runtime unverified |
| CMD-WORKFLOW-text-center | Implemented source; runtime unverified |
| CMD-WORKFLOW-text-right | Implemented source; runtime unverified |
| CMD-WORKFLOW-text-justify | Implemented source; runtime unverified |
| CMD-WORKFLOW-text-top | Implemented source; runtime unverified |
| CMD-WORKFLOW-text-middle | Implemented source; runtime unverified |
| CMD-WORKFLOW-text-bottom | Implemented source; runtime unverified |
| CMD-WORKFLOW-text-horizontal | Implemented source; runtime unverified |
| CMD-WORKFLOW-text-downward | Implemented source; runtime unverified |
| CMD-WORKFLOW-text-upward | Implemented source; runtime unverified |
| CMD-WORKFLOW-color-fill | Implemented source; runtime unverified |
| CMD-WORKFLOW-color-outline | Implemented source; runtime unverified |
| CMD-WORKFLOW-color-text | Implemented source; runtime unverified |
| CMD-SOURCE-excel | Implemented source; runtime unverified |
| CMD-SOURCE-csv | Implemented source; runtime unverified |
| CMD-SOURCE-json | Implemented source; runtime unverified |
| CMD-SOURCE-xml | Implemented source; runtime unverified |
| CMD-DIRECT-ppiTemplateLibrary | Implemented source; runtime unverified |
| CMD-DIRECT-ppiRenderShapes | Implemented source; runtime unverified |
| CMD-DIRECT-ppiInsert | Implemented source; runtime unverified |
| CMD-DIRECT-ppiContextProperties | Implemented source; runtime unverified |
| CMD-DIRECT-ppiLibraryLayout | Implemented source; runtime unverified |
| CMD-DIRECT-ppiTemplateAlignment | Implemented source; runtime unverified |
| CMD-DIRECT-ppiTemplatePalette | Implemented source; runtime unverified |
| CMD-DIRECT-ppiInspect | Implemented source; runtime unverified |
| CMD-DIRECT-ppiOpenWorkspace | Implemented source; runtime unverified |
| CMD-DIRECT-ppiDemo | Implemented source; runtime unverified |
| CMD-DIRECT-ppiCompatibility | Implemented source; runtime unverified |
| CONTEXT-FONT | Implemented source; runtime unverified |
| CONTEXT-IDENTITY | Implemented source; runtime unverified |
| CONTEXT-MIXED | Implemented source; runtime unverified |
| LIB-SEARCH | Implemented source; runtime unverified |
| LIB-REFRESH | Implemented source; runtime unverified |
| LIB-THUMBNAILS | Implemented source; runtime unverified |
| LIB-ORIGINAL | Implemented source; runtime unverified |
| LIB-DESTINATION | Implemented source; runtime unverified |
| LIB-DESIGNED | Implemented source; runtime unverified |
| LIB-PALETTES | Implemented source; runtime unverified |
| LIB-ALIGNMENT | Implemented source; runtime unverified |
| LIB-BUILDS | Implemented source; runtime unverified |
| LIB-OBJECTS | Implemented source; runtime unverified |
| LIB-LAYOUTLINK | Implemented source; runtime unverified |
| AGENDA-NATIVEEDIT | Implemented source; runtime unverified |
| AGENDA-MOVE | Implemented source; runtime unverified |
| S0-RECOVERY | Implemented source; runtime unverified |
| AUTOMATION-COM | Implemented source; runtime unverified |
