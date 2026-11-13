# Complete NexWave feature inventory

Evidence baseline: current inspected checkout, HEAD `ffaec352cbdca666f69f3c402d334108e4115933` plus parent export hash inventory. **VERIFIED** means source implementation here; current native runtime acceptance is not implied. Source analysis levels and limitations are in `04-powerpoint-integration/vsto-source-review.csv`.

## Counting and status rules

The catalog contains **190 explicit entries**: **172 executable UI command routes** and **18 cross-cutting capabilities**. Commands can be aliases to the same underlying operation; this is not 190 independent business engines, reference-product parity or current native runtime passes. The 90 DemoTools delegates include 68 statically declared entries and 22 expanded arrange modes. Routed entries additionally cover 29 element controls, 9 Tools commands, 29 workflow commands, 4 source formats and 11 direct callbacks. Generate Linked Report is a direct alias of demo-tour and is not double-counted.

FeatureInventory.cs separately contains 103 descriptive requirement/status records. They are preserved in descriptive-requirement-inventory.csv, with original wording, and never counted as implemented solely because their label exists. Source implementations, generated examples and future intentions can overlap those records.

Use [implementation details](feature-implementation-details.md) for every field and [CSV](feature-source-mapping.csv) for machine-readable traceability. Each feature has a stable documentation ID, distinct from native object IDs. This catalog covers VSTO user-facing routes/capabilities; core-library/headless developer APIs are also cataloged in the API/data chapters, not silently counted as extra ribbon features.

## Catalog

| Feature ID | Name | Kind / category | Source status | Evidence |
|---|---|---|---|---|
| CMD-DEMO-demo-tour | Generate linked report | command route / Start here | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:63` |
| CMD-DEMO-column | Clustered column | command route / Charts | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:64` |
| CMD-DEMO-bar | Clustered bar | command route / Charts | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:65` |
| CMD-DEMO-stacked-column | Stacked column | command route / Charts | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:66` |
| CMD-DEMO-percent-column | 100% stacked column | command route / Charts | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:67` |
| CMD-DEMO-stacked-bar | Stacked bar | command route / Charts | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:68` |
| CMD-DEMO-percent-bar | 100% stacked bar | command route / Charts | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:69` |
| CMD-DEMO-line | Line | command route / Charts | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:70` |
| CMD-DEMO-profile | Profile | command route / Charts | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:71` |
| CMD-DEMO-area | Area | command route / Charts | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:72` |
| CMD-DEMO-pie | Pie | command route / Charts | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:73` |
| CMD-DEMO-doughnut | Doughnut | command route / Charts | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:74` |
| CMD-DEMO-scatter | Scatter | command route / Charts | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:75` |
| CMD-DEMO-bubble | Bubble | command route / Charts | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:76` |
| CMD-DEMO-combo | Column and line | command route / Charts | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:77` |
| CMD-DEMO-waterfall | Waterfall diagram | command route / Charts | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:78` |
| CMD-DEMO-mekko | Mekko diagram (%) | command route / Charts | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:79` |
| CMD-DEMO-mekko-units | Mekko diagram (units) | command route / Charts | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:80` |
| CMD-DEMO-gantt | Gantt diagram | command route / Charts | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:81` |
| CMD-DEMO-text | Text box | command route / Elements | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:82` |
| CMD-DEMO-rounded | Rounded rectangle | command route / Elements | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:83` |
| CMD-DEMO-process | Process flow | command route / Elements | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:84` |
| CMD-DEMO-checkbox | Checkbox | command route / Elements | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:85` |
| CMD-DEMO-harvey | Harvey ball | command route / Elements | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:86` |
| CMD-DEMO-table | Table | command route / Elements | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:87` |
| CMD-DEMO-agenda | Agenda chapter | command route / Elements | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:88` |
| CMD-DEMO-image | Image from file | command route / Assets | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:89` |
| CMD-DEMO-symbol | Insert symbol | command route / Elements | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:90` |
| CMD-DEMO-auto-layout | Grid layout | command route / Arrange | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:96` |
| CMD-DEMO-paste-text | Paste lines into objects | command route / Text | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:97` |
| CMD-DEMO-font | Replace font | command route / Text | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:98` |
| CMD-DEMO-font-size | Resize font | command route / Text | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:99` |
| CMD-DEMO-decimal | Switch decimal separators | command route / Text | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:100` |
| CMD-DEMO-language | Proofing language | command route / Text | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:101` |
| CMD-DEMO-edit-data | Edit chart data | command route / Data | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:102` |
| CMD-DEMO-legend | Toggle legend | command route / Chart options | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:103` |
| CMD-DEMO-labels | Add value labels | command route / Chart options | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:104` |
| CMD-DEMO-data-table | Toggle data table | command route / Chart options | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:105` |
| CMD-DEMO-chart-style | Chart style | command route / Chart options | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:106` |
| CMD-DEMO-value-line | Value line annotation | command route / Chart options | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:107` |
| CMD-DEMO-difference | Difference arrow annotation | command route / Chart options | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:108` |
| CMD-DEMO-cagr | CAGR annotation | command route / Chart options | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:109` |
| CMD-DEMO-scan | Chart scanner | command route / Data | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:110` |
| CMD-DEMO-path | File path on slide | command route / Assets | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:111` |
| CMD-DEMO-library | Import slides | command route / Assets | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:112` |
| CMD-DEMO-template | Apply template | command route / Assets | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:113` |
| CMD-DEMO-export-slide | Export slide image | command route / Assets | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:114` |
| CMD-DEMO-theme | Apply local brand colors | command route / Assets | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:115` |
| CMD-DEMO-axes | Toggle value-axis title | command route / Chart options | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:116` |
| CMD-DEMO-plot-area | Style chart and plot area | command route / Chart options | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:117` |
| CMD-DEMO-chart-title | Chart title | command route / Chart options | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:118` |
| CMD-DEMO-duplicate-slide | Duplicate slide | command route / Slides | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:119` |
| CMD-DEMO-save-slides | Save selected slides | command route / Slides | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:120` |
| CMD-DEMO-clean-comments | Remove presentation comments | command route / Slides | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:121` |
| CMD-DEMO-round-data | Round chart workbook data | command route / Data | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:122` |
| CMD-DEMO-sheet-style | Style chart workbook header | command route / Data | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:123` |
| CMD-DEMO-import-json | Insert chart from JSON | command route / Automation | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:124` |
| CMD-DEMO-export-json | Export selected chart JSON | command route / Automation | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:125` |
| CMD-DEMO-mini-toolbar | Quick arrange toolbar | command route / Arrange | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:126` |
| CMD-DEMO-table-data | Edit table datasheet | command route / Data | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:127` |
| CMD-DEMO-library-browser | Browse local asset library | command route / Assets | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:128` |
| CMD-DEMO-agenda-layout | Agenda table of contents | command route / Assets | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:129` |
| CMD-DEMO-style-file | Apply style file | command route / Assets | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:130` |
| CMD-DEMO-excel-cagr | Insert Excel CAGR formula | command route / Data | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:131` |
| CMD-DEMO-link-excel | Chart from Excel file | command route / Data | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:132` |
| CMD-DEMO-refresh-excel | Refresh selected Excel link | command route / Data | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:133` |
| CMD-DEMO-manage-links | Manage local workbook links | command route / Data | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:134` |
| CMD-DEMO-slide-workbook | Current slide workbook list | command route / Data | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:135` |
| CMD-DEMO-left | Left | command route / Arrange | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:91` |
| CMD-DEMO-center | Center | command route / Arrange | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:91` |
| CMD-DEMO-right | Right | command route / Arrange | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:91` |
| CMD-DEMO-top | Top | command route / Arrange | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:91` |
| CMD-DEMO-middle | Middle | command route / Arrange | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:91` |
| CMD-DEMO-bottom | Bottom | command route / Arrange | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:91` |
| CMD-DEMO-width | Width | command route / Arrange | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:91` |
| CMD-DEMO-height | Height | command route / Arrange | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:91` |
| CMD-DEMO-size | Size | command route / Arrange | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:91` |
| CMD-DEMO-horizontal | Horizontal | command route / Arrange | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:91` |
| CMD-DEMO-vertical | Vertical | command route / Arrange | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:91` |
| CMD-DEMO-swap | Swap | command route / Arrange | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:91` |
| CMD-DEMO-scale-up | Scale Up | command route / Arrange | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:91` |
| CMD-DEMO-scale-down | Scale Down | command route / Arrange | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:91` |
| CMD-DEMO-rotate | Rotate | command route / Arrange | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:91` |
| CMD-DEMO-flip-horizontal | Flip Horizontal | command route / Arrange | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:91` |
| CMD-DEMO-flip-vertical | Flip Vertical | command route / Arrange | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:91` |
| CMD-DEMO-duplicate | Duplicate | command route / Arrange | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:91` |
| CMD-DEMO-group | Group | command route / Arrange | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:91` |
| CMD-DEMO-ungroup | Ungroup | command route / Arrange | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:91` |
| CMD-DEMO-front | Front | command route / Arrange | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:91` |
| CMD-DEMO-back | Back | command route / Arrange | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoTools.cs:91` |
| CMD-ELEMENT-pentagon | Right pentagon | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:59` |
| CMD-ELEMENT-chevron | Chevron | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:59` |
| CMD-ELEMENT-connector | Universal connector | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:59` |
| CMD-ELEMENT-edit | Edit selected element | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:60` |
| CMD-ELEMENT-fit | Fit text to unlocked edges | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:60` |
| CMD-ELEMENT-convert-text | Convert to Text Box | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:61` |
| CMD-ELEMENT-convert-pentagon | Convert to Right Pentagon | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:61` |
| CMD-ELEMENT-convert-chevron | Convert to Chevron | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:61` |
| CMD-ELEMENT-convert-rounded | Convert to Rounded Rectangle | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:61` |
| CMD-ELEMENT-duplicate-left | Duplicate left | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:62` |
| CMD-ELEMENT-duplicate-right | Duplicate right | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:62` |
| CMD-ELEMENT-duplicate-up | Duplicate above | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:62` |
| CMD-ELEMENT-duplicate-down | Duplicate below | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:62` |
| CMD-ELEMENT-indicator-next | Increase / next indicator state | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:63` |
| CMD-ELEMENT-indicator-previous | Decrease / previous state | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:63` |
| CMD-ELEMENT-indicator-keys | Indicator keyboard controls | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:63` |
| CMD-ELEMENT-convert-harvey | Convert indicator to Harvey ball | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:63` |
| CMD-ELEMENT-convert-checkbox | Convert indicator to Checkbox | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:63` |
| CMD-ELEMENT-grid-create | New process grid | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:64` |
| CMD-ELEMENT-grid-row | Add grid row | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:64` |
| CMD-ELEMENT-grid-column | Add grid column | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:64` |
| CMD-ELEMENT-grid-remove-row | Hide grid row | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:64` |
| CMD-ELEMENT-grid-remove-column | Hide grid column | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:64` |
| CMD-ELEMENT-grid-left | Move grid column left | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:64` |
| CMD-ELEMENT-grid-right | Move grid column right | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:64` |
| CMD-ELEMENT-grid-stripe-rows | Stripe grid rows | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:64` |
| CMD-ELEMENT-grid-stripe-columns | Stripe grid columns | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:64` |
| CMD-ELEMENT-grid-resize | Resize process grid | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:64` |
| CMD-ELEMENT-refresh | Refresh element geometry | command route / Element Controls | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:64` |
| CMD-TOOLS-symbol-picker | Insert symbol… | command route / Tools | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:52` |
| CMD-TOOLS-decimal-safe | Switch decimal symbols… | command route / Tools | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:52` |
| CMD-TOOLS-raster-copy | Create picture copy (retain original) | command route / Tools | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:52` |
| CMD-TOOLS-chart-workbooks | Chart workbooks… | command route / Tools | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:52` |
| CMD-TOOLS-data-links | Data links… | command route / Tools | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:52` |
| CMD-TOOLS-save-selected | Save selected slides… | command route / Tools | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:52` |
| CMD-TOOLS-selected-style | Apply selected-object style… | command route / Tools | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:52` |
| CMD-TOOLS-review | Review presentation… | command route / Tools | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:52` |
| CMD-TOOLS-help | Element tools help | command route / Tools | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:52` |
| CMD-WORKFLOW-agenda-insert | Chapter | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:21` |
| CMD-WORKFLOW-agenda-contents | Table of contents | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:21` |
| CMD-WORKFLOW-agenda-settings | Manage Agenda | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:21` |
| CMD-WORKFLOW-agenda-split | Split agenda here | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:21` |
| CMD-WORKFLOW-agenda-refresh | Update agenda | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:21` |
| CMD-WORKFLOW-percent-area | 100% stacked area | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:29` |
| CMD-WORKFLOW-series-settings | Series type and axis | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:29` |
| CMD-WORKFLOW-axis-settings | Value axis settings | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:29` |
| CMD-WORKFLOW-match-scales | Match value scales | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:29` |
| CMD-WORKFLOW-annotation-refresh | Refresh selected annotation | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:29` |
| CMD-WORKFLOW-connections | Configure services | command route / Workflow | Partial: nonsecret configuration only; no live service execution | `apps/powerpoint-vsto/DemoRibbon.cs:30` |
| CMD-WORKFLOW-relink-selected | Relink selected chart | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:35` |
| CMD-WORKFLOW-map-excel | Link Excel Range | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:35` |
| CMD-WORKFLOW-data-manager | Data Link Manager | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:35` |
| CMD-WORKFLOW-quality-review | Presentation review | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:41` |
| CMD-WORKFLOW-export-pdf | Export PDF | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:42` |
| CMD-WORKFLOW-text-left | Text left | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:37` |
| CMD-WORKFLOW-text-center | Text center | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:37` |
| CMD-WORKFLOW-text-right | Text right | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:37` |
| CMD-WORKFLOW-text-justify | Text justify | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:37` |
| CMD-WORKFLOW-text-top | Text top | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:37` |
| CMD-WORKFLOW-text-middle | Text middle | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:37` |
| CMD-WORKFLOW-text-bottom | Text bottom | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:37` |
| CMD-WORKFLOW-text-horizontal | Text horizontal | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:37` |
| CMD-WORKFLOW-text-downward | Text downward | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:37` |
| CMD-WORKFLOW-text-upward | Text upward | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:37` |
| CMD-WORKFLOW-color-fill | Choose fill color | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:40` |
| CMD-WORKFLOW-color-outline | Choose outline color | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:40` |
| CMD-WORKFLOW-color-text | Choose text color | command route / Workflow | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:40` |
| CMD-SOURCE-excel | Excel sheet | command route / Data Source | Implemented source; runtime unverified | `apps/powerpoint-vsto/DataSources.cs:267` |
| CMD-SOURCE-csv | CSV file | command route / Data Source | Implemented source; runtime unverified | `apps/powerpoint-vsto/DataSources.cs:268` |
| CMD-SOURCE-json | JSON file | command route / Data Source | Implemented source; runtime unverified | `apps/powerpoint-vsto/DataSources.cs:269` |
| CMD-SOURCE-xml | XML file | command route / Data Source | Implemented source; runtime unverified | `apps/powerpoint-vsto/DataSources.cs:270` |
| CMD-DIRECT-ppiTemplateLibrary | Library | command route / Direct ribbon | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:24` |
| CMD-DIRECT-ppiRenderShapes | Editable shape columns | command route / Direct ribbon | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:26` |
| CMD-DIRECT-ppiInsert | Linked column chart | command route / Direct ribbon | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:27` |
| CMD-DIRECT-ppiContextProperties | Properties | command route / Direct ribbon | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:28` |
| CMD-DIRECT-ppiLibraryLayout | Layout Links | command route / Direct ribbon | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:38` |
| CMD-DIRECT-ppiTemplateAlignment | Template Alignment | command route / Direct ribbon | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:39` |
| CMD-DIRECT-ppiTemplatePalette | Template Colors | command route / Direct ribbon | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:40` |
| CMD-DIRECT-ppiInspect | Check & Sync | command route / Direct ribbon | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:41` |
| CMD-DIRECT-ppiOpenWorkspace | Open Workspace | command route / Direct ribbon | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:44` |
| CMD-DIRECT-ppiDemo | Explore Demo Tools | command route / Direct ribbon | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:44` |
| CMD-DIRECT-ppiCompatibility | Office Compatibility | command route / Direct ribbon | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoRibbon.cs:44` |
| CONTEXT-FONT | Context font family, size, bold and italic | capability / Capability | Implemented source; runtime unverified | `apps/powerpoint-vsto/ContextPropertiesController.cs:59` |
| CONTEXT-IDENTITY | Stale-selection rejection and foreground behavior | capability / Capability | Implemented source; runtime unverified | `apps/powerpoint-vsto/ContextPropertiesController.cs:33` |
| CONTEXT-MIXED | Mixed group/table/native properties | capability / Capability | Implemented source; runtime unverified | `apps/powerpoint-vsto/ContextPanelModel.cs:62` |
| LIB-SEARCH | Indexed library/presentation/search/tag navigation | capability / Capability | Implemented source; runtime unverified | `apps/powerpoint-vsto/LocalLibraryWorkspace.cs:53` |
| LIB-REFRESH | Configured source refresh and cancellation | capability / Capability | Implemented source; runtime unverified | `apps/powerpoint-vsto/LocalLibraryWorkspace.cs:53` |
| LIB-THUMBNAILS | Native thumbnail gallery and larger preview | capability / Capability | Implemented source; runtime unverified | `apps/powerpoint-vsto/LocalLibraryGallery.cs:126` |
| LIB-ORIGINAL | Source-original slide/design insertion | capability / Capability | Implemented source; runtime unverified | `apps/powerpoint-vsto/TemplateLibrary.cs:489` |
| LIB-DESTINATION | Destination-design local slide insertion | capability / Capability | Implemented source; runtime unverified | `apps/powerpoint-vsto/LocalLibraryGallery.cs:136` |
| LIB-DESIGNED | Designed template catalog and editable drafts | capability / Capability | Implemented source; runtime unverified | `apps/powerpoint-vsto/TemplateCatalog.cs:1` |
| LIB-PALETTES | Twenty subtle palettes and default Executive Slate | capability / Capability | Implemented source; runtime unverified | `apps/powerpoint-vsto/TemplateCatalog.cs:63` |
| LIB-ALIGNMENT | Designed-template text alignment | capability / Capability | Implemented source; runtime unverified | `apps/powerpoint-vsto/TemplateLibrary.cs:133` |
| LIB-BUILDS | Native template builds and transitions | capability / Capability | Implemented source; runtime unverified | `apps/powerpoint-vsto/TemplateLibrary.cs:106` |
| LIB-OBJECTS | Static template group in current slide | capability / Capability | Implemented source; runtime unverified | `apps/powerpoint-vsto/TemplateLibrary.cs:66` |
| LIB-LAYOUTLINK | Live edge/offset layout links | capability / Capability | Implemented source; runtime unverified | `apps/powerpoint-vsto/TemplateLibrary.cs:411` |
| AGENDA-NATIVEEDIT | Native caption edits and hierarchy propagation | capability / Capability | Implemented source; runtime unverified | `apps/powerpoint-vsto/ChapterAgenda.cs:70` |
| AGENDA-MOVE | Chapter subtree/content movement | capability / Capability | Implemented source; runtime unverified | `apps/powerpoint-vsto/ChapterAgenda.cs:120` |
| S0-RECOVERY | S0 metadata edit capture and duplicate identity repair | capability / Capability | Implemented source; runtime unverified | `apps/powerpoint-vsto/PowerPointHost.cs:360` |
| AUTOMATION-COM | Bounded COM automation observation/insertion | capability / Capability | Implemented source; runtime unverified | `apps/powerpoint-vsto/DemoApi.cs:12` |
