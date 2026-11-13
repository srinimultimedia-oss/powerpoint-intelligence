# Feature implementation details

<!-- generated-toc -->
## Contents

- [Navigation](#navigation)
- [CMD-DEMO-demo-tour](#cmd-demo-demo-tour)
- [CMD-DEMO-column](#cmd-demo-column)
- [CMD-DEMO-bar](#cmd-demo-bar)
- [CMD-DEMO-stacked-column](#cmd-demo-stacked-column)
- [CMD-DEMO-percent-column](#cmd-demo-percent-column)
- [CMD-DEMO-stacked-bar](#cmd-demo-stacked-bar)
- [CMD-DEMO-percent-bar](#cmd-demo-percent-bar)
- [CMD-DEMO-line](#cmd-demo-line)
- [CMD-DEMO-profile](#cmd-demo-profile)
- [CMD-DEMO-area](#cmd-demo-area)
- [CMD-DEMO-pie](#cmd-demo-pie)
- [CMD-DEMO-doughnut](#cmd-demo-doughnut)
- [CMD-DEMO-scatter](#cmd-demo-scatter)
- [CMD-DEMO-bubble](#cmd-demo-bubble)
- [CMD-DEMO-combo](#cmd-demo-combo)
- [CMD-DEMO-waterfall](#cmd-demo-waterfall)
- [CMD-DEMO-mekko](#cmd-demo-mekko)
- [CMD-DEMO-mekko-units](#cmd-demo-mekko-units)
- [CMD-DEMO-gantt](#cmd-demo-gantt)
- [CMD-DEMO-text](#cmd-demo-text)
- [CMD-DEMO-rounded](#cmd-demo-rounded)
- [CMD-DEMO-process](#cmd-demo-process)
- [CMD-DEMO-checkbox](#cmd-demo-checkbox)
- [CMD-DEMO-harvey](#cmd-demo-harvey)
- [CMD-DEMO-table](#cmd-demo-table)
- [CMD-DEMO-agenda](#cmd-demo-agenda)
- [CMD-DEMO-image](#cmd-demo-image)
- [CMD-DEMO-symbol](#cmd-demo-symbol)
- [CMD-DEMO-auto-layout](#cmd-demo-auto-layout)
- [CMD-DEMO-paste-text](#cmd-demo-paste-text)
- [CMD-DEMO-font](#cmd-demo-font)
- [CMD-DEMO-font-size](#cmd-demo-font-size)
- [CMD-DEMO-decimal](#cmd-demo-decimal)
- [CMD-DEMO-language](#cmd-demo-language)
- [CMD-DEMO-edit-data](#cmd-demo-edit-data)
- [CMD-DEMO-legend](#cmd-demo-legend)
- [CMD-DEMO-labels](#cmd-demo-labels)
- [CMD-DEMO-data-table](#cmd-demo-data-table)
- [CMD-DEMO-chart-style](#cmd-demo-chart-style)
- [CMD-DEMO-value-line](#cmd-demo-value-line)
- [CMD-DEMO-difference](#cmd-demo-difference)
- [CMD-DEMO-cagr](#cmd-demo-cagr)
- [CMD-DEMO-scan](#cmd-demo-scan)
- [CMD-DEMO-path](#cmd-demo-path)
- [CMD-DEMO-library](#cmd-demo-library)
- [CMD-DEMO-template](#cmd-demo-template)
- [CMD-DEMO-export-slide](#cmd-demo-export-slide)
- [CMD-DEMO-theme](#cmd-demo-theme)
- [CMD-DEMO-axes](#cmd-demo-axes)
- [CMD-DEMO-plot-area](#cmd-demo-plot-area)
- [CMD-DEMO-chart-title](#cmd-demo-chart-title)
- [CMD-DEMO-duplicate-slide](#cmd-demo-duplicate-slide)
- [CMD-DEMO-save-slides](#cmd-demo-save-slides)
- [CMD-DEMO-clean-comments](#cmd-demo-clean-comments)
- [CMD-DEMO-round-data](#cmd-demo-round-data)
- [CMD-DEMO-sheet-style](#cmd-demo-sheet-style)
- [CMD-DEMO-import-json](#cmd-demo-import-json)
- [CMD-DEMO-export-json](#cmd-demo-export-json)
- [CMD-DEMO-mini-toolbar](#cmd-demo-mini-toolbar)
- [CMD-DEMO-table-data](#cmd-demo-table-data)
- [CMD-DEMO-library-browser](#cmd-demo-library-browser)
- [CMD-DEMO-agenda-layout](#cmd-demo-agenda-layout)
- [CMD-DEMO-style-file](#cmd-demo-style-file)
- [CMD-DEMO-excel-cagr](#cmd-demo-excel-cagr)
- [CMD-DEMO-link-excel](#cmd-demo-link-excel)
- [CMD-DEMO-refresh-excel](#cmd-demo-refresh-excel)
- [CMD-DEMO-manage-links](#cmd-demo-manage-links)
- [CMD-DEMO-slide-workbook](#cmd-demo-slide-workbook)
- [CMD-DEMO-left](#cmd-demo-left)
- [CMD-DEMO-center](#cmd-demo-center)
- [CMD-DEMO-right](#cmd-demo-right)
- [CMD-DEMO-top](#cmd-demo-top)
- [CMD-DEMO-middle](#cmd-demo-middle)
- [CMD-DEMO-bottom](#cmd-demo-bottom)
- [CMD-DEMO-width](#cmd-demo-width)
- [CMD-DEMO-height](#cmd-demo-height)
- [CMD-DEMO-size](#cmd-demo-size)
- [CMD-DEMO-horizontal](#cmd-demo-horizontal)
- [CMD-DEMO-vertical](#cmd-demo-vertical)
- [CMD-DEMO-swap](#cmd-demo-swap)
- [CMD-DEMO-scale-up](#cmd-demo-scale-up)
- [CMD-DEMO-scale-down](#cmd-demo-scale-down)
- [CMD-DEMO-rotate](#cmd-demo-rotate)
- [CMD-DEMO-flip-horizontal](#cmd-demo-flip-horizontal)
- [CMD-DEMO-flip-vertical](#cmd-demo-flip-vertical)
- [CMD-DEMO-duplicate](#cmd-demo-duplicate)
- [CMD-DEMO-group](#cmd-demo-group)
- [CMD-DEMO-ungroup](#cmd-demo-ungroup)
- [CMD-DEMO-front](#cmd-demo-front)
- [CMD-DEMO-back](#cmd-demo-back)
- [CMD-ELEMENT-pentagon](#cmd-element-pentagon)
- [CMD-ELEMENT-chevron](#cmd-element-chevron)
- [CMD-ELEMENT-connector](#cmd-element-connector)
- [CMD-ELEMENT-edit](#cmd-element-edit)
- [CMD-ELEMENT-fit](#cmd-element-fit)
- [CMD-ELEMENT-convert-text](#cmd-element-convert-text)
- [CMD-ELEMENT-convert-pentagon](#cmd-element-convert-pentagon)
- [CMD-ELEMENT-convert-chevron](#cmd-element-convert-chevron)
- [CMD-ELEMENT-convert-rounded](#cmd-element-convert-rounded)
- [CMD-ELEMENT-duplicate-left](#cmd-element-duplicate-left)
- [CMD-ELEMENT-duplicate-right](#cmd-element-duplicate-right)
- [CMD-ELEMENT-duplicate-up](#cmd-element-duplicate-up)
- [CMD-ELEMENT-duplicate-down](#cmd-element-duplicate-down)
- [CMD-ELEMENT-indicator-next](#cmd-element-indicator-next)
- [CMD-ELEMENT-indicator-previous](#cmd-element-indicator-previous)
- [CMD-ELEMENT-indicator-keys](#cmd-element-indicator-keys)
- [CMD-ELEMENT-convert-harvey](#cmd-element-convert-harvey)
- [CMD-ELEMENT-convert-checkbox](#cmd-element-convert-checkbox)
- [CMD-ELEMENT-grid-create](#cmd-element-grid-create)
- [CMD-ELEMENT-grid-row](#cmd-element-grid-row)
- [CMD-ELEMENT-grid-column](#cmd-element-grid-column)
- [CMD-ELEMENT-grid-remove-row](#cmd-element-grid-remove-row)
- [CMD-ELEMENT-grid-remove-column](#cmd-element-grid-remove-column)
- [CMD-ELEMENT-grid-left](#cmd-element-grid-left)
- [CMD-ELEMENT-grid-right](#cmd-element-grid-right)
- [CMD-ELEMENT-grid-stripe-rows](#cmd-element-grid-stripe-rows)
- [CMD-ELEMENT-grid-stripe-columns](#cmd-element-grid-stripe-columns)
- [CMD-ELEMENT-grid-resize](#cmd-element-grid-resize)
- [CMD-ELEMENT-refresh](#cmd-element-refresh)
- [CMD-TOOLS-symbol-picker](#cmd-tools-symbol-picker)
- [CMD-TOOLS-decimal-safe](#cmd-tools-decimal-safe)
- [CMD-TOOLS-raster-copy](#cmd-tools-raster-copy)
- [CMD-TOOLS-chart-workbooks](#cmd-tools-chart-workbooks)
- [CMD-TOOLS-data-links](#cmd-tools-data-links)
- [CMD-TOOLS-save-selected](#cmd-tools-save-selected)
- [CMD-TOOLS-selected-style](#cmd-tools-selected-style)
- [CMD-TOOLS-review](#cmd-tools-review)
- [CMD-TOOLS-help](#cmd-tools-help)
- [CMD-WORKFLOW-agenda-insert](#cmd-workflow-agenda-insert)
- [CMD-WORKFLOW-agenda-contents](#cmd-workflow-agenda-contents)
- [CMD-WORKFLOW-agenda-settings](#cmd-workflow-agenda-settings)
- [CMD-WORKFLOW-agenda-split](#cmd-workflow-agenda-split)
- [CMD-WORKFLOW-agenda-refresh](#cmd-workflow-agenda-refresh)
- [CMD-WORKFLOW-percent-area](#cmd-workflow-percent-area)
- [CMD-WORKFLOW-series-settings](#cmd-workflow-series-settings)
- [CMD-WORKFLOW-axis-settings](#cmd-workflow-axis-settings)
- [CMD-WORKFLOW-match-scales](#cmd-workflow-match-scales)
- [CMD-WORKFLOW-annotation-refresh](#cmd-workflow-annotation-refresh)
- [CMD-WORKFLOW-connections](#cmd-workflow-connections)
- [CMD-WORKFLOW-relink-selected](#cmd-workflow-relink-selected)
- [CMD-WORKFLOW-map-excel](#cmd-workflow-map-excel)
- [CMD-WORKFLOW-data-manager](#cmd-workflow-data-manager)
- [CMD-WORKFLOW-quality-review](#cmd-workflow-quality-review)
- [CMD-WORKFLOW-export-pdf](#cmd-workflow-export-pdf)
- [CMD-WORKFLOW-text-left](#cmd-workflow-text-left)
- [CMD-WORKFLOW-text-center](#cmd-workflow-text-center)
- [CMD-WORKFLOW-text-right](#cmd-workflow-text-right)
- [CMD-WORKFLOW-text-justify](#cmd-workflow-text-justify)
- [CMD-WORKFLOW-text-top](#cmd-workflow-text-top)
- [CMD-WORKFLOW-text-middle](#cmd-workflow-text-middle)
- [CMD-WORKFLOW-text-bottom](#cmd-workflow-text-bottom)
- [CMD-WORKFLOW-text-horizontal](#cmd-workflow-text-horizontal)
- [CMD-WORKFLOW-text-downward](#cmd-workflow-text-downward)
- [CMD-WORKFLOW-text-upward](#cmd-workflow-text-upward)
- [CMD-WORKFLOW-color-fill](#cmd-workflow-color-fill)
- [CMD-WORKFLOW-color-outline](#cmd-workflow-color-outline)
- [CMD-WORKFLOW-color-text](#cmd-workflow-color-text)
- [CMD-SOURCE-excel](#cmd-source-excel)
- [CMD-SOURCE-csv](#cmd-source-csv)
- [CMD-SOURCE-json](#cmd-source-json)
- [CMD-SOURCE-xml](#cmd-source-xml)
- [CMD-DIRECT-ppiTemplateLibrary](#cmd-direct-ppitemplatelibrary)
- [CMD-DIRECT-ppiRenderShapes](#cmd-direct-ppirendershapes)
- [CMD-DIRECT-ppiInsert](#cmd-direct-ppiinsert)
- [CMD-DIRECT-ppiContextProperties](#cmd-direct-ppicontextproperties)
- [CMD-DIRECT-ppiLibraryLayout](#cmd-direct-ppilibrarylayout)
- [CMD-DIRECT-ppiTemplateAlignment](#cmd-direct-ppitemplatealignment)
- [CMD-DIRECT-ppiTemplatePalette](#cmd-direct-ppitemplatepalette)
- [CMD-DIRECT-ppiInspect](#cmd-direct-ppiinspect)
- [CMD-DIRECT-ppiOpenWorkspace](#cmd-direct-ppiopenworkspace)
- [CMD-DIRECT-ppiDemo](#cmd-direct-ppidemo)
- [CMD-DIRECT-ppiCompatibility](#cmd-direct-ppicompatibility)
- [CONTEXT-FONT](#context-font)
- [CONTEXT-IDENTITY](#context-identity)
- [CONTEXT-MIXED](#context-mixed)
- [LIB-SEARCH](#lib-search)
- [LIB-REFRESH](#lib-refresh)
- [LIB-THUMBNAILS](#lib-thumbnails)
- [LIB-ORIGINAL](#lib-original)
- [LIB-DESTINATION](#lib-destination)
- [LIB-DESIGNED](#lib-designed)
- [LIB-PALETTES](#lib-palettes)
- [LIB-ALIGNMENT](#lib-alignment)
- [LIB-BUILDS](#lib-builds)
- [LIB-OBJECTS](#lib-objects)
- [LIB-LAYOUTLINK](#lib-layoutlink)
- [AGENDA-NATIVEEDIT](#agenda-nativeedit)
- [AGENDA-MOVE](#agenda-move)
- [S0-RECOVERY](#s0-recovery)
- [AUTOMATION-COM](#automation-com)
<!-- end-generated-toc -->

Evidence baseline: current inspected checkout, HEAD `ffaec352cbdca666f69f3c402d334108e4115933` plus parent export hash inventory. **VERIFIED** means source implementation here; current native runtime acceptance is not implied. Source analysis levels and limitations are in `04-powerpoint-integration/vsto-source-review.csv`.

## Navigation

- [CMD-DEMO-demo-tour — Generate linked report](#cmd-demo-demo-tour)
- [CMD-DEMO-column — Clustered column](#cmd-demo-column)
- [CMD-DEMO-bar — Clustered bar](#cmd-demo-bar)
- [CMD-DEMO-stacked-column — Stacked column](#cmd-demo-stacked-column)
- [CMD-DEMO-percent-column — 100% stacked column](#cmd-demo-percent-column)
- [CMD-DEMO-stacked-bar — Stacked bar](#cmd-demo-stacked-bar)
- [CMD-DEMO-percent-bar — 100% stacked bar](#cmd-demo-percent-bar)
- [CMD-DEMO-line — Line](#cmd-demo-line)
- [CMD-DEMO-profile — Profile](#cmd-demo-profile)
- [CMD-DEMO-area — Area](#cmd-demo-area)
- [CMD-DEMO-pie — Pie](#cmd-demo-pie)
- [CMD-DEMO-doughnut — Doughnut](#cmd-demo-doughnut)
- [CMD-DEMO-scatter — Scatter](#cmd-demo-scatter)
- [CMD-DEMO-bubble — Bubble](#cmd-demo-bubble)
- [CMD-DEMO-combo — Column and line](#cmd-demo-combo)
- [CMD-DEMO-waterfall — Waterfall diagram](#cmd-demo-waterfall)
- [CMD-DEMO-mekko — Mekko diagram (%)](#cmd-demo-mekko)
- [CMD-DEMO-mekko-units — Mekko diagram (units)](#cmd-demo-mekko-units)
- [CMD-DEMO-gantt — Gantt diagram](#cmd-demo-gantt)
- [CMD-DEMO-text — Text box](#cmd-demo-text)
- [CMD-DEMO-rounded — Rounded rectangle](#cmd-demo-rounded)
- [CMD-DEMO-process — Process flow](#cmd-demo-process)
- [CMD-DEMO-checkbox — Checkbox](#cmd-demo-checkbox)
- [CMD-DEMO-harvey — Harvey ball](#cmd-demo-harvey)
- [CMD-DEMO-table — Table](#cmd-demo-table)
- [CMD-DEMO-agenda — Agenda chapter](#cmd-demo-agenda)
- [CMD-DEMO-image — Image from file](#cmd-demo-image)
- [CMD-DEMO-symbol — Insert symbol](#cmd-demo-symbol)
- [CMD-DEMO-auto-layout — Grid layout](#cmd-demo-auto-layout)
- [CMD-DEMO-paste-text — Paste lines into objects](#cmd-demo-paste-text)
- [CMD-DEMO-font — Replace font](#cmd-demo-font)
- [CMD-DEMO-font-size — Resize font](#cmd-demo-font-size)
- [CMD-DEMO-decimal — Switch decimal separators](#cmd-demo-decimal)
- [CMD-DEMO-language — Proofing language](#cmd-demo-language)
- [CMD-DEMO-edit-data — Edit chart data](#cmd-demo-edit-data)
- [CMD-DEMO-legend — Toggle legend](#cmd-demo-legend)
- [CMD-DEMO-labels — Add value labels](#cmd-demo-labels)
- [CMD-DEMO-data-table — Toggle data table](#cmd-demo-data-table)
- [CMD-DEMO-chart-style — Chart style](#cmd-demo-chart-style)
- [CMD-DEMO-value-line — Value line annotation](#cmd-demo-value-line)
- [CMD-DEMO-difference — Difference arrow annotation](#cmd-demo-difference)
- [CMD-DEMO-cagr — CAGR annotation](#cmd-demo-cagr)
- [CMD-DEMO-scan — Chart scanner](#cmd-demo-scan)
- [CMD-DEMO-path — File path on slide](#cmd-demo-path)
- [CMD-DEMO-library — Import slides](#cmd-demo-library)
- [CMD-DEMO-template — Apply template](#cmd-demo-template)
- [CMD-DEMO-export-slide — Export slide image](#cmd-demo-export-slide)
- [CMD-DEMO-theme — Apply local brand colors](#cmd-demo-theme)
- [CMD-DEMO-axes — Toggle value-axis title](#cmd-demo-axes)
- [CMD-DEMO-plot-area — Style chart and plot area](#cmd-demo-plot-area)
- [CMD-DEMO-chart-title — Chart title](#cmd-demo-chart-title)
- [CMD-DEMO-duplicate-slide — Duplicate slide](#cmd-demo-duplicate-slide)
- [CMD-DEMO-save-slides — Save selected slides](#cmd-demo-save-slides)
- [CMD-DEMO-clean-comments — Remove presentation comments](#cmd-demo-clean-comments)
- [CMD-DEMO-round-data — Round chart workbook data](#cmd-demo-round-data)
- [CMD-DEMO-sheet-style — Style chart workbook header](#cmd-demo-sheet-style)
- [CMD-DEMO-import-json — Insert chart from JSON](#cmd-demo-import-json)
- [CMD-DEMO-export-json — Export selected chart JSON](#cmd-demo-export-json)
- [CMD-DEMO-mini-toolbar — Quick arrange toolbar](#cmd-demo-mini-toolbar)
- [CMD-DEMO-table-data — Edit table datasheet](#cmd-demo-table-data)
- [CMD-DEMO-library-browser — Browse local asset library](#cmd-demo-library-browser)
- [CMD-DEMO-agenda-layout — Agenda table of contents](#cmd-demo-agenda-layout)
- [CMD-DEMO-style-file — Apply style file](#cmd-demo-style-file)
- [CMD-DEMO-excel-cagr — Insert Excel CAGR formula](#cmd-demo-excel-cagr)
- [CMD-DEMO-link-excel — Chart from Excel file](#cmd-demo-link-excel)
- [CMD-DEMO-refresh-excel — Refresh selected Excel link](#cmd-demo-refresh-excel)
- [CMD-DEMO-manage-links — Manage local workbook links](#cmd-demo-manage-links)
- [CMD-DEMO-slide-workbook — Current slide workbook list](#cmd-demo-slide-workbook)
- [CMD-DEMO-left — Left](#cmd-demo-left)
- [CMD-DEMO-center — Center](#cmd-demo-center)
- [CMD-DEMO-right — Right](#cmd-demo-right)
- [CMD-DEMO-top — Top](#cmd-demo-top)
- [CMD-DEMO-middle — Middle](#cmd-demo-middle)
- [CMD-DEMO-bottom — Bottom](#cmd-demo-bottom)
- [CMD-DEMO-width — Width](#cmd-demo-width)
- [CMD-DEMO-height — Height](#cmd-demo-height)
- [CMD-DEMO-size — Size](#cmd-demo-size)
- [CMD-DEMO-horizontal — Horizontal](#cmd-demo-horizontal)
- [CMD-DEMO-vertical — Vertical](#cmd-demo-vertical)
- [CMD-DEMO-swap — Swap](#cmd-demo-swap)
- [CMD-DEMO-scale-up — Scale Up](#cmd-demo-scale-up)
- [CMD-DEMO-scale-down — Scale Down](#cmd-demo-scale-down)
- [CMD-DEMO-rotate — Rotate](#cmd-demo-rotate)
- [CMD-DEMO-flip-horizontal — Flip Horizontal](#cmd-demo-flip-horizontal)
- [CMD-DEMO-flip-vertical — Flip Vertical](#cmd-demo-flip-vertical)
- [CMD-DEMO-duplicate — Duplicate](#cmd-demo-duplicate)
- [CMD-DEMO-group — Group](#cmd-demo-group)
- [CMD-DEMO-ungroup — Ungroup](#cmd-demo-ungroup)
- [CMD-DEMO-front — Front](#cmd-demo-front)
- [CMD-DEMO-back — Back](#cmd-demo-back)
- [CMD-ELEMENT-pentagon — Right pentagon](#cmd-element-pentagon)
- [CMD-ELEMENT-chevron — Chevron](#cmd-element-chevron)
- [CMD-ELEMENT-connector — Universal connector](#cmd-element-connector)
- [CMD-ELEMENT-edit — Edit selected element](#cmd-element-edit)
- [CMD-ELEMENT-fit — Fit text to unlocked edges](#cmd-element-fit)
- [CMD-ELEMENT-convert-text — Convert to Text Box](#cmd-element-convert-text)
- [CMD-ELEMENT-convert-pentagon — Convert to Right Pentagon](#cmd-element-convert-pentagon)
- [CMD-ELEMENT-convert-chevron — Convert to Chevron](#cmd-element-convert-chevron)
- [CMD-ELEMENT-convert-rounded — Convert to Rounded Rectangle](#cmd-element-convert-rounded)
- [CMD-ELEMENT-duplicate-left — Duplicate left](#cmd-element-duplicate-left)
- [CMD-ELEMENT-duplicate-right — Duplicate right](#cmd-element-duplicate-right)
- [CMD-ELEMENT-duplicate-up — Duplicate above](#cmd-element-duplicate-up)
- [CMD-ELEMENT-duplicate-down — Duplicate below](#cmd-element-duplicate-down)
- [CMD-ELEMENT-indicator-next — Increase / next indicator state](#cmd-element-indicator-next)
- [CMD-ELEMENT-indicator-previous — Decrease / previous state](#cmd-element-indicator-previous)
- [CMD-ELEMENT-indicator-keys — Indicator keyboard controls](#cmd-element-indicator-keys)
- [CMD-ELEMENT-convert-harvey — Convert indicator to Harvey ball](#cmd-element-convert-harvey)
- [CMD-ELEMENT-convert-checkbox — Convert indicator to Checkbox](#cmd-element-convert-checkbox)
- [CMD-ELEMENT-grid-create — New process grid](#cmd-element-grid-create)
- [CMD-ELEMENT-grid-row — Add grid row](#cmd-element-grid-row)
- [CMD-ELEMENT-grid-column — Add grid column](#cmd-element-grid-column)
- [CMD-ELEMENT-grid-remove-row — Hide grid row](#cmd-element-grid-remove-row)
- [CMD-ELEMENT-grid-remove-column — Hide grid column](#cmd-element-grid-remove-column)
- [CMD-ELEMENT-grid-left — Move grid column left](#cmd-element-grid-left)
- [CMD-ELEMENT-grid-right — Move grid column right](#cmd-element-grid-right)
- [CMD-ELEMENT-grid-stripe-rows — Stripe grid rows](#cmd-element-grid-stripe-rows)
- [CMD-ELEMENT-grid-stripe-columns — Stripe grid columns](#cmd-element-grid-stripe-columns)
- [CMD-ELEMENT-grid-resize — Resize process grid](#cmd-element-grid-resize)
- [CMD-ELEMENT-refresh — Refresh element geometry](#cmd-element-refresh)
- [CMD-TOOLS-symbol-picker — Insert symbol…](#cmd-tools-symbol-picker)
- [CMD-TOOLS-decimal-safe — Switch decimal symbols…](#cmd-tools-decimal-safe)
- [CMD-TOOLS-raster-copy — Create picture copy (retain original)](#cmd-tools-raster-copy)
- [CMD-TOOLS-chart-workbooks — Chart workbooks…](#cmd-tools-chart-workbooks)
- [CMD-TOOLS-data-links — Data links…](#cmd-tools-data-links)
- [CMD-TOOLS-save-selected — Save selected slides…](#cmd-tools-save-selected)
- [CMD-TOOLS-selected-style — Apply selected-object style…](#cmd-tools-selected-style)
- [CMD-TOOLS-review — Review presentation…](#cmd-tools-review)
- [CMD-TOOLS-help — Element tools help](#cmd-tools-help)
- [CMD-WORKFLOW-agenda-insert — Chapter](#cmd-workflow-agenda-insert)
- [CMD-WORKFLOW-agenda-contents — Table of contents](#cmd-workflow-agenda-contents)
- [CMD-WORKFLOW-agenda-settings — Manage Agenda](#cmd-workflow-agenda-settings)
- [CMD-WORKFLOW-agenda-split — Split agenda here](#cmd-workflow-agenda-split)
- [CMD-WORKFLOW-agenda-refresh — Update agenda](#cmd-workflow-agenda-refresh)
- [CMD-WORKFLOW-percent-area — 100% stacked area](#cmd-workflow-percent-area)
- [CMD-WORKFLOW-series-settings — Series type and axis](#cmd-workflow-series-settings)
- [CMD-WORKFLOW-axis-settings — Value axis settings](#cmd-workflow-axis-settings)
- [CMD-WORKFLOW-match-scales — Match value scales](#cmd-workflow-match-scales)
- [CMD-WORKFLOW-annotation-refresh — Refresh selected annotation](#cmd-workflow-annotation-refresh)
- [CMD-WORKFLOW-connections — Configure services](#cmd-workflow-connections)
- [CMD-WORKFLOW-relink-selected — Relink selected chart](#cmd-workflow-relink-selected)
- [CMD-WORKFLOW-map-excel — Link Excel Range](#cmd-workflow-map-excel)
- [CMD-WORKFLOW-data-manager — Data Link Manager](#cmd-workflow-data-manager)
- [CMD-WORKFLOW-quality-review — Presentation review](#cmd-workflow-quality-review)
- [CMD-WORKFLOW-export-pdf — Export PDF](#cmd-workflow-export-pdf)
- [CMD-WORKFLOW-text-left — Text left](#cmd-workflow-text-left)
- [CMD-WORKFLOW-text-center — Text center](#cmd-workflow-text-center)
- [CMD-WORKFLOW-text-right — Text right](#cmd-workflow-text-right)
- [CMD-WORKFLOW-text-justify — Text justify](#cmd-workflow-text-justify)
- [CMD-WORKFLOW-text-top — Text top](#cmd-workflow-text-top)
- [CMD-WORKFLOW-text-middle — Text middle](#cmd-workflow-text-middle)
- [CMD-WORKFLOW-text-bottom — Text bottom](#cmd-workflow-text-bottom)
- [CMD-WORKFLOW-text-horizontal — Text horizontal](#cmd-workflow-text-horizontal)
- [CMD-WORKFLOW-text-downward — Text downward](#cmd-workflow-text-downward)
- [CMD-WORKFLOW-text-upward — Text upward](#cmd-workflow-text-upward)
- [CMD-WORKFLOW-color-fill — Choose fill color](#cmd-workflow-color-fill)
- [CMD-WORKFLOW-color-outline — Choose outline color](#cmd-workflow-color-outline)
- [CMD-WORKFLOW-color-text — Choose text color](#cmd-workflow-color-text)
- [CMD-SOURCE-excel — Excel sheet](#cmd-source-excel)
- [CMD-SOURCE-csv — CSV file](#cmd-source-csv)
- [CMD-SOURCE-json — JSON file](#cmd-source-json)
- [CMD-SOURCE-xml — XML file](#cmd-source-xml)
- [CMD-DIRECT-ppiTemplateLibrary — Library](#cmd-direct-ppitemplatelibrary)
- [CMD-DIRECT-ppiRenderShapes — Editable shape columns](#cmd-direct-ppirendershapes)
- [CMD-DIRECT-ppiInsert — Linked column chart](#cmd-direct-ppiinsert)
- [CMD-DIRECT-ppiContextProperties — Properties](#cmd-direct-ppicontextproperties)
- [CMD-DIRECT-ppiLibraryLayout — Layout Links](#cmd-direct-ppilibrarylayout)
- [CMD-DIRECT-ppiTemplateAlignment — Template Alignment](#cmd-direct-ppitemplatealignment)
- [CMD-DIRECT-ppiTemplatePalette — Template Colors](#cmd-direct-ppitemplatepalette)
- [CMD-DIRECT-ppiInspect — Check & Sync](#cmd-direct-ppiinspect)
- [CMD-DIRECT-ppiOpenWorkspace — Open Workspace](#cmd-direct-ppiopenworkspace)
- [CMD-DIRECT-ppiDemo — Explore Demo Tools](#cmd-direct-ppidemo)
- [CMD-DIRECT-ppiCompatibility — Office Compatibility](#cmd-direct-ppicompatibility)
- [CONTEXT-FONT — Context font family, size, bold and italic](#context-font)
- [CONTEXT-IDENTITY — Stale-selection rejection and foreground behavior](#context-identity)
- [CONTEXT-MIXED — Mixed group/table/native properties](#context-mixed)
- [LIB-SEARCH — Indexed library/presentation/search/tag navigation](#lib-search)
- [LIB-REFRESH — Configured source refresh and cancellation](#lib-refresh)
- [LIB-THUMBNAILS — Native thumbnail gallery and larger preview](#lib-thumbnails)
- [LIB-ORIGINAL — Source-original slide/design insertion](#lib-original)
- [LIB-DESTINATION — Destination-design local slide insertion](#lib-destination)
- [LIB-DESIGNED — Designed template catalog and editable drafts](#lib-designed)
- [LIB-PALETTES — Twenty subtle palettes and default Executive Slate](#lib-palettes)
- [LIB-ALIGNMENT — Designed-template text alignment](#lib-alignment)
- [LIB-BUILDS — Native template builds and transitions](#lib-builds)
- [LIB-OBJECTS — Static template group in current slide](#lib-objects)
- [LIB-LAYOUTLINK — Live edge/offset layout links](#lib-layoutlink)
- [AGENDA-NATIVEEDIT — Native caption edits and hierarchy propagation](#agenda-nativeedit)
- [AGENDA-MOVE — Chapter subtree/content movement](#agenda-move)
- [S0-RECOVERY — S0 metadata edit capture and duplicate identity repair](#s0-recovery)
- [AUTOMATION-COM — Bounded COM automation observation/insertion](#automation-com)

## CMD-DEMO-demo-tour

### Generate linked report

**Business purpose:** Add chart report slides using a linked source in the current presentation..

**UI entry:** NexWave > Start here > Generate linked report.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; CreateTour.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Add chart report slides using a linked source in the current presentation..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; CreateTour.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:63.

## CMD-DEMO-column

### Clustered column

**Business purpose:** Native chart from the linked data source; edit values through Edit Data..

**UI entry:** NexWave > Charts > Clustered column.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DataSources.cs.

**Responsible classes:** DemoTools; InsertChart.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Native chart from the linked data source; edit values through Edit Data..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; InsertChart.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/MappedExcelOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:64.

## CMD-DEMO-bar

### Clustered bar

**Business purpose:** Native horizontal comparison chart from the linked data source..

**UI entry:** NexWave > Charts > Clustered bar.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DataSources.cs.

**Responsible classes:** DemoTools; InsertChart.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Native horizontal comparison chart from the linked data source..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; InsertChart.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/MappedExcelOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:65.

## CMD-DEMO-stacked-column

### Stacked column

**Business purpose:** Native stacked column chart from the linked data source..

**UI entry:** NexWave > Charts > Stacked column.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DataSources.cs.

**Responsible classes:** DemoTools; InsertChart.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Native stacked column chart from the linked data source..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; InsertChart.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/MappedExcelOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:66.

## CMD-DEMO-percent-column

### 100% stacked column

**Business purpose:** Native normalized stacked column chart from the linked data source..

**UI entry:** NexWave > Charts > 100% stacked column.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DataSources.cs.

**Responsible classes:** DemoTools; InsertChart.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Native normalized stacked column chart from the linked data source..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; InsertChart.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/MappedExcelOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:67.

## CMD-DEMO-stacked-bar

### Stacked bar

**Business purpose:** Native stacked bar chart from the linked data source..

**UI entry:** NexWave > Charts > Stacked bar.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DataSources.cs.

**Responsible classes:** DemoTools; InsertChart.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Native stacked bar chart from the linked data source..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; InsertChart.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/MappedExcelOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:68.

## CMD-DEMO-percent-bar

### 100% stacked bar

**Business purpose:** Native normalized stacked bar chart from the linked data source..

**UI entry:** NexWave > Charts > 100% stacked bar.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DataSources.cs.

**Responsible classes:** DemoTools; InsertChart.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Native normalized stacked bar chart from the linked data source..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; InsertChart.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/MappedExcelOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:69.

## CMD-DEMO-line

### Line

**Business purpose:** Native line chart from the linked data source..

**UI entry:** NexWave > Charts > Line.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DataSources.cs.

**Responsible classes:** DemoTools; InsertChart.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Native line chart from the linked data source..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; InsertChart.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/MappedExcelOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:70.

## CMD-DEMO-profile

### Profile

**Business purpose:** Native line comparison with series plotted in rows; not a rotated value-axis renderer..

**UI entry:** NexWave > Charts > Profile.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DataSources.cs.

**Responsible classes:** DemoTools; InsertChart.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Native line comparison with series plotted in rows; not a rotated value-axis renderer..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; InsertChart.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/MappedExcelOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:71.

## CMD-DEMO-area

### Area

**Business purpose:** Native stacked area chart from the linked data source..

**UI entry:** NexWave > Charts > Area.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DataSources.cs.

**Responsible classes:** DemoTools; InsertChart.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Native stacked area chart from the linked data source..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; InsertChart.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/MappedExcelOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:72.

## CMD-DEMO-pie

### Pie

**Business purpose:** Native pie chart from the first linked numeric series..

**UI entry:** NexWave > Charts > Pie.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DataSources.cs.

**Responsible classes:** DemoTools; InsertChart.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Native pie chart from the first linked numeric series..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; InsertChart.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/MappedExcelOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:73.

## CMD-DEMO-doughnut

### Doughnut

**Business purpose:** Native doughnut chart from the first linked numeric series..

**UI entry:** NexWave > Charts > Doughnut.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DataSources.cs.

**Responsible classes:** DemoTools; InsertChart.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Native doughnut chart from the first linked numeric series..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; InsertChart.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/MappedExcelOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:74.

## CMD-DEMO-scatter

### Scatter

**Business purpose:** Native numeric XY chart from the linked data source..

**UI entry:** NexWave > Charts > Scatter.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DataSources.cs.

**Responsible classes:** DemoTools; InsertChart.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Native numeric XY chart from the linked data source..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; InsertChart.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/MappedExcelOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:75.

## CMD-DEMO-bubble

### Bubble

**Business purpose:** Native chart from linked X/Y and Size columns..

**UI entry:** NexWave > Charts > Bubble.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DataSources.cs.

**Responsible classes:** DemoTools; InsertChart.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Native chart from linked X/Y and Size columns..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; InsertChart.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/MappedExcelOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:76.

## CMD-DEMO-combo

### Column and line

**Business purpose:** Native column chart with its second series rendered as a line..

**UI entry:** NexWave > Charts > Column and line.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DataSources.cs.

**Responsible classes:** DemoTools; InsertCombo.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Native column chart with its second series rendered as a line..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; InsertCombo.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/MappedExcelOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:77.

## CMD-DEMO-waterfall

### Waterfall diagram

**Business purpose:** Linked signed movements from zero rendered as editable shapes..

**UI entry:** NexWave > Charts > Waterfall diagram.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DataSources.cs.

**Responsible classes:** DemoTools; InsertWaterfall.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Linked signed movements from zero rendered as editable shapes..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; InsertWaterfall.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/MappedExcelOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:78.

## CMD-DEMO-mekko

### Mekko diagram (%)

**Business purpose:** Linked nonnegative values: category totals set widths; series shares set percent heights..

**UI entry:** NexWave > Charts > Mekko diagram (%).

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DataSources.cs.

**Responsible classes:** DemoTools; InsertMekko.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Linked nonnegative values: category totals set widths; series shares set percent heights..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; InsertMekko.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/MappedExcelOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:79.

## CMD-DEMO-mekko-units

### Mekko diagram (units)

**Business purpose:** Linked nonnegative category totals set widths and heights..

**UI entry:** NexWave > Charts > Mekko diagram (units).

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DataSources.cs.

**Responsible classes:** DemoTools; InsertMekko.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Linked nonnegative category totals set widths and heights..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; InsertMekko.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/MappedExcelOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:80.

## CMD-DEMO-gantt

### Gantt diagram

**Business purpose:** Linked Task, numeric Start and positive Duration columns..

**UI entry:** NexWave > Charts > Gantt diagram.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DataSources.cs.

**Responsible classes:** DemoTools; InsertGantt.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Linked Task, numeric Start and positive Duration columns..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; InsertGantt.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/MappedExcelOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:81.

## CMD-DEMO-text

### Text box

**Business purpose:** Insert a native editable text box..

**UI entry:** NexWave > Elements > Text box.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/TextShapeElements.cs.

**Responsible classes:** DemoTools; TextShapeElements.Insert.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Insert a native editable text box..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; TextShapeElements.Insert.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:82.

## CMD-DEMO-rounded

### Rounded rectangle

**Business purpose:** Insert an editable rounded rectangle..

**UI entry:** NexWave > Elements > Rounded rectangle.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/TextShapeElements.cs.

**Responsible classes:** DemoTools; TextShapeElements.Insert.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Insert an editable rounded rectangle..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; TextShapeElements.Insert.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:83.

## CMD-DEMO-process

### Process flow

**Business purpose:** Insert an editable two-row process grid with managed cell controls..

**UI entry:** NexWave > Elements > Process flow.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/ProcessFlowTools.cs.

**Responsible classes:** DemoTools; ProcessFlowTools.Create.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Insert an editable two-row process grid with managed cell controls..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; ProcessFlowTools.Create.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:84.

## CMD-DEMO-checkbox

### Checkbox

**Business purpose:** Insert an editable vector checkbox with empty, check and cross states..

**UI entry:** NexWave > Elements > Checkbox.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/IndicatorElements.cs.

**Responsible classes:** DemoTools; IndicatorElements.Insert.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Insert an editable vector checkbox with empty, check and cross states..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; IndicatorElements.Insert.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:85.

## CMD-DEMO-harvey

### Harvey ball

**Business purpose:** Insert an editable segmented vector indicator; use Element Controls to change its fill..

**UI entry:** NexWave > Elements > Harvey ball.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/IndicatorElements.cs.

**Responsible classes:** DemoTools; IndicatorElements.Insert.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Insert an editable segmented vector indicator; use Element Controls to change its fill..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; IndicatorElements.Insert.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:86.

## CMD-DEMO-table

### Table

**Business purpose:** Insert a native editable table from linked headers/text/values, limited to 400 cells..

**UI entry:** NexWave > Elements > Table.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DataSources.cs.

**Responsible classes:** DemoTools; InsertTable.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Insert a native editable table from linked headers/text/values, limited to 400 cells..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; InsertTable.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:87.

## CMD-DEMO-agenda

### Agenda chapter

**Business purpose:** Insert a native editable chapter slide and synchronize managed chapter agendas..

**UI entry:** NexWave > Elements > Agenda chapter.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/ChapterAgenda.cs.

**Responsible classes:** DemoTools; Agenda.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Insert a native editable chapter slide and synchronize managed chapter agendas..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; Agenda.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:88.

## CMD-DEMO-image

### Image from file

**Business purpose:** Insert a local image chosen by you..

**UI entry:** NexWave > Assets > Image from file.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; InsertImage.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Insert a local image chosen by you..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; InsertImage.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:89.

## CMD-DEMO-symbol

### Insert symbol

**Business purpose:** Choose text or a Unicode symbol to insert..

**UI entry:** NexWave > Elements > Insert symbol.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; native delegate.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Choose text or a Unicode symbol to insert..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; native delegate.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:90.

## CMD-DEMO-auto-layout

### Grid layout

**Business purpose:** Arrange selected objects using their dimensions and slide bounds..

**UI entry:** NexWave > Arrange > Grid layout.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; Grid.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Arrange selected objects using their dimensions and slide bounds..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; Grid.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:96.

## CMD-DEMO-paste-text

### Paste lines into objects

**Business purpose:** Enter one line per selected text object; validates the count before editing..

**UI entry:** NexWave > Text > Paste lines into objects.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; PasteText.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Enter one line per selected text object; validates the count before editing..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; PasteText.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:97.

## CMD-DEMO-font

### Replace font

**Business purpose:** Apply a chosen font family to selected text objects..

**UI entry:** NexWave > Text > Replace font.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; Font.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Apply a chosen font family to selected text objects..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; Font.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:98.

## CMD-DEMO-font-size

### Resize font

**Business purpose:** Apply a chosen point size to selected text objects..

**UI entry:** NexWave > Text > Resize font.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; Font.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Apply a chosen point size to selected text objects..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; Font.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:99.

## CMD-DEMO-decimal

### Switch decimal separators

**Business purpose:** Swap dots and commas only between digits in selected text; chart data is unchanged..

**UI entry:** NexWave > Text > Switch decimal separators.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; SwitchDecimals.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Swap dots and commas only between digits in selected text; chart data is unchanged..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; SwitchDecimals.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:100.

## CMD-DEMO-language

### Proofing language

**Business purpose:** Apply an Office language ID to selected text (for example 1033 English or 1031 German)..

**UI entry:** NexWave > Text > Proofing language.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; Language.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Apply an Office language ID to selected text (for example 1033 English or 1031 German)..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; Language.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:101.

## CMD-DEMO-edit-data

### Edit chart data

**Business purpose:** Open the selected native chart's embedded Excel workbook..

**UI entry:** NexWave > Data > Edit chart data.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; Chart.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Open the selected native chart's embedded Excel workbook..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; Chart.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:102.

## CMD-DEMO-legend

### Toggle legend

**Business purpose:** Show or hide the selected native chart's legend..

**UI entry:** NexWave > Chart options > Toggle legend.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; native delegate.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Show or hide the selected native chart's legend..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; native delegate.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:103.

## CMD-DEMO-labels

### Add value labels

**Business purpose:** Apply native data labels to the selected chart..

**UI entry:** NexWave > Chart options > Add value labels.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; Chart.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Apply native data labels to the selected chart..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; Chart.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:104.

## CMD-DEMO-data-table

### Toggle data table

**Business purpose:** Show or hide a native chart data table where supported..

**UI entry:** NexWave > Chart options > Toggle data table.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; native delegate.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Show or hide a native chart data table where supported..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; native delegate.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:105.

## CMD-DEMO-chart-style

### Chart style

**Business purpose:** Apply a numbered native Office chart style (1–48)..

**UI entry:** NexWave > Chart options > Chart style.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; ChartStyle.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Apply a numbered native Office chart style (1–48)..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; ChartStyle.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:106.

## CMD-DEMO-value-line

### Value line annotation

**Business purpose:** Insert a native editable reference line mapped to the selected chart axis; explicit annotation refresh supported..

**UI entry:** NexWave > Chart options > Value line annotation.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/ChartAnnotations.cs.

**Responsible classes:** DemoTools; ValueLine.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Insert a native editable reference line mapped to the selected chart axis; explicit annotation refresh supported..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; ValueLine.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:107.

## CMD-DEMO-difference

### Difference arrow annotation

**Business purpose:** Calculate a difference between native series points and insert an editable annotation; explicit refresh supported..

**UI entry:** NexWave > Chart options > Difference arrow annotation.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/ChartAnnotations.cs.

**Responsible classes:** DemoTools; Difference.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Calculate a difference between native series points and insert an editable annotation; explicit refresh supported..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; Difference.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:108.

## CMD-DEMO-cagr

### CAGR annotation

**Business purpose:** Calculate CAGR from entered start/end/years and insert editable text..

**UI entry:** NexWave > Chart options > CAGR annotation.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; Cagr.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Calculate CAGR from entered start/end/years and insert editable text..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; Cagr.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:109.

## CMD-DEMO-scan

### Chart scanner

**Business purpose:** Count native charts and pictures in the current presentation; no chart recognition from pixels..

**UI entry:** NexWave > Data > Chart scanner.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; Scan.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Count native charts and pictures in the current presentation; no chart recognition from pixels..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; Scan.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:110.

## CMD-DEMO-path

### File path on slide

**Business purpose:** Insert the current presentation's path as editable text..

**UI entry:** NexWave > Assets > File path on slide.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; Text.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Insert the current presentation's path as editable text..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; Text.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:111.

## CMD-DEMO-library

### Import slides

**Business purpose:** Choose a local PowerPoint file and insert its slides after the active slide..

**UI entry:** NexWave > Assets > Import slides.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; ImportSlides.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Choose a local PowerPoint file and insert its slides after the active slide..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; ImportSlides.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:112.

## CMD-DEMO-template

### Apply template

**Business purpose:** Choose and apply a local .potx template to this presentation..

**UI entry:** NexWave > Assets > Apply template.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; ApplyTemplate.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Choose and apply a local .potx template to this presentation..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; ApplyTemplate.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:113.

## CMD-DEMO-export-slide

### Export slide image

**Business purpose:** Save the current slide to a PNG file chosen by you..

**UI entry:** NexWave > Assets > Export slide image.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; ExportSlide.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Save the current slide to a PNG file chosen by you..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; ExportSlide.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:114.

## CMD-DEMO-theme

### Apply local brand colors

**Business purpose:** Apply an original navy/teal palette to selected shapes and Segoe UI to their text. Does not change the master theme..

**UI entry:** NexWave > Assets > Apply local brand colors.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; Brand.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Apply an original navy/teal palette to selected shapes and Segoe UI to their text. Does not change the master theme..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; Brand.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:115.

## CMD-DEMO-axes

### Toggle value-axis title

**Business purpose:** Set a native value-axis title where the chart supports axes..

**UI entry:** NexWave > Chart options > Toggle value-axis title.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; AxisTitle.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Set a native value-axis title where the chart supports axes..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; AxisTitle.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:116.

## CMD-DEMO-plot-area

### Style chart and plot area

**Business purpose:** Apply a white chart area and pale plot area to the selected native chart..

**UI entry:** NexWave > Chart options > Style chart and plot area.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; native delegate.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Apply a white chart area and pale plot area to the selected native chart..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; native delegate.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:117.

## CMD-DEMO-chart-title

### Chart title

**Business purpose:** Set the selected native chart title..

**UI entry:** NexWave > Chart options > Chart title.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; native delegate.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Set the selected native chart title..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; native delegate.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:118.

## CMD-DEMO-duplicate-slide

### Duplicate slide

**Business purpose:** Duplicate the current slide using PowerPoint's native operation..

**UI entry:** NexWave > Slides > Duplicate slide.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; Slide.Duplicate.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Duplicate the current slide using PowerPoint's native operation..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; Slide.Duplicate.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:119.

## CMD-DEMO-save-slides

### Save selected slides

**Business purpose:** Save selected native slides to a new PowerPoint file; no email is sent..

**UI entry:** NexWave > Slides > Save selected slides.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; SaveSlides.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Save selected native slides to a new PowerPoint file; no email is sent..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; SaveSlides.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:120.

## CMD-DEMO-clean-comments

### Remove presentation comments

**Business purpose:** Explicitly confirm removal of legacy comments from this presentation. Modern threaded comments are not handled..

**UI entry:** NexWave > Slides > Remove presentation comments.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; CleanComments.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Explicitly confirm removal of legacy comments from this presentation. Modern threaded comments are not handled..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; CleanComments.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:121.

## CMD-DEMO-round-data

### Round chart workbook data

**Business purpose:** Round numeric constants in the selected chart's embedded workbook, preserving formulas. Maximum 1,000 used cells..

**UI entry:** NexWave > Data > Round chart workbook data.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DataSources.cs.

**Responsible classes:** DemoTools; RoundData.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Round numeric constants in the selected chart's embedded workbook, preserving formulas. Maximum 1,000 used cells..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; RoundData.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:122.

## CMD-DEMO-sheet-style

### Style chart workbook header

**Business purpose:** Apply a bold shaded header to the selected chart's embedded workbook..

**UI entry:** NexWave > Data > Style chart workbook header.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DataSources.cs.

**Responsible classes:** DemoTools; StyleSheet.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Apply a bold shaded header to the selected chart's embedded workbook..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; StyleSheet.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:123.

## CMD-DEMO-import-json

### Insert chart from JSON

**Business purpose:** Choose a bounded, validated S0 ChartDocument JSON file. Inserts a managed clustered-column chart; not an arbitrary COM command interface..

**UI entry:** NexWave > Automation > Insert chart from JSON.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; ImportJson.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Choose a bounded, validated S0 ChartDocument JSON file. Inserts a managed clustered-column chart; not an arbitrary COM command interface..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; ImportJson.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:124.

## CMD-DEMO-export-json

### Export selected chart JSON

**Business purpose:** Export validated metadata of a managed S0 chart to a JSON file. Other demo charts do not carry this model..

**UI entry:** NexWave > Automation > Export selected chart JSON.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; ExportJson.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Export validated metadata of a managed S0 chart to a JSON file. Other demo charts do not carry this model..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; ExportJson.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:125.

## CMD-DEMO-mini-toolbar

### Quick arrange toolbar

**Business purpose:** Open a floating toolbar for selected objects. Select slide objects, then use alignment or size actions..

**UI entry:** NexWave > Arrange > Quick arrange toolbar.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; MiniToolbar.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Open a floating toolbar for selected objects. Select slide objects, then use alignment or size actions..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; MiniToolbar.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:126.

## CMD-DEMO-table-data

### Edit table datasheet

**Business purpose:** Edit a selected native table in a grid, then Apply. Limited to 400 cells..

**UI entry:** NexWave > Data > Edit table datasheet.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; TableData.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Edit a selected native table in a grid, then Apply. Limited to 400 cells..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; TableData.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:127.

## CMD-DEMO-library-browser

### Browse local asset library

**Business purpose:** Browse a folder for local images, presentations and templates. No network library service or deletion is provided..

**UI entry:** NexWave > Assets > Browse local asset library.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; Library.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Browse a folder for local images, presentations and templates. No network library service or deletion is provided..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; Library.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:128.

## CMD-DEMO-agenda-layout

### Agenda table of contents

**Business purpose:** Insert a synchronized table of contents with chapter hyperlinks and no current-chapter highlight..

**UI entry:** NexWave > Assets > Agenda table of contents.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/ChapterAgenda.cs.

**Responsible classes:** DemoTools; AgendaCurrent.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Insert a synchronized table of contents with chapter hyperlinks and no current-chapter highlight..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; AgendaCurrent.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:129.

## CMD-DEMO-style-file

### Apply style file

**Business purpose:** Choose an original JSON style with FontName, FontSize, FillRgb and TextRgb; applies to selected text shapes..

**UI entry:** NexWave > Assets > Apply style file.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; StyleFile.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Choose an original JSON style with FontName, FontSize, FillRgb and TextRgb; applies to selected text shapes..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; StyleFile.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:130.

## CMD-DEMO-excel-cagr

### Insert Excel CAGR formula

**Business purpose:** Write a CAGR formula to a chosen cell in the selected chart workbook; enter positive inputs and a single-cell address..

**UI entry:** NexWave > Data > Insert Excel CAGR formula.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; ExcelCagr.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Write a CAGR formula to a chosen cell in the selected chart workbook; enter positive inputs and a single-cell address..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; ExcelCagr.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:131.

## CMD-DEMO-link-excel

### Chart from Excel file

**Business purpose:** Insert a native column chart from the first worksheet of a chosen local .xlsx. Manual refresh; maximum 50 rows by 15 columns. Source path is stored in the presentation..

**UI entry:** NexWave > Data > Chart from Excel file.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoLinks.cs.

**Responsible classes:** DemoTools; native delegate.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Insert a native column chart from the first worksheet of a chosen local .xlsx. Manual refresh; maximum 50 rows by 15 columns. Source path is stored in the presentation..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; native delegate.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:132.

## CMD-DEMO-refresh-excel

### Refresh selected Excel link

**Business purpose:** Read the saved file path again. Missing sources fail without automatically rebinding. Does not refresh in the background..

**UI entry:** NexWave > Data > Refresh selected Excel link.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoLinks.cs.

**Responsible classes:** DemoTools; DemoLinks.Refresh.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Read the saved file path again. Missing sources fail without automatically rebinding. Does not refresh in the background..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; DemoLinks.Refresh.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:133.

## CMD-DEMO-manage-links

### Manage local workbook links

**Business purpose:** Count workbook links and show the manual-refresh behavior..

**UI entry:** NexWave > Data > Manage local workbook links.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoLinks.cs.

**Responsible classes:** DemoTools; DemoLinks.Describe.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then Count workbook links and show the manual-refresh behavior..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; DemoLinks.Describe.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:134.

## CMD-DEMO-slide-workbook

### Current slide workbook list

**Business purpose:** List charts on this slide; select each chart to edit its embedded workbook. No unified shared slide workbook is created..

**UI entry:** NexWave > Data > Current slide workbook list.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools; Notify.

**Execution workflow:** ThisAddIn.RunDemo resolves ID; DemoTools.Execute starts Undo; delegate performs selection/source preflight then List charts on this slide; select each chart to edit its embedded workbook. No unified shared slide workbook is created..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools; Notify.

**Data structures:** LinkedTable/DataSourceLink for linked charts; native selection or input dialog for local tools.

**Failure behavior:** Selection/input preflight plus ThisAddIn.Run exception containment; native family-specific limitations apply..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/DemoContractTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:135.

## CMD-DEMO-left

### Left

**Business purpose:** Arrange selected native objects using left semantics..

**UI entry:** NexWave > Layout > Arrange > left.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools.Arrange.

**Execution workflow:** Resolve current shapes; group/ungroup, swap exactly two, distribute at least three, or apply selected native property relative to first shape..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools.Arrange.

**Data structures:** Native ShapeRange; coordinates/size/ZOrder.

**Failure behavior:** Select native objects first. Swap exactly2; distribution>=3. Some simple mutations rely on native Undo rather than compensating rollback..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:91.

## CMD-DEMO-center

### Center

**Business purpose:** Arrange selected native objects using center semantics..

**UI entry:** NexWave > Layout > Arrange > center.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools.Arrange.

**Execution workflow:** Resolve current shapes; group/ungroup, swap exactly two, distribute at least three, or apply selected native property relative to first shape..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools.Arrange.

**Data structures:** Native ShapeRange; coordinates/size/ZOrder.

**Failure behavior:** Select native objects first. Swap exactly2; distribution>=3. Some simple mutations rely on native Undo rather than compensating rollback..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:91.

## CMD-DEMO-right

### Right

**Business purpose:** Arrange selected native objects using right semantics..

**UI entry:** NexWave > Layout > Arrange > right.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools.Arrange.

**Execution workflow:** Resolve current shapes; group/ungroup, swap exactly two, distribute at least three, or apply selected native property relative to first shape..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools.Arrange.

**Data structures:** Native ShapeRange; coordinates/size/ZOrder.

**Failure behavior:** Select native objects first. Swap exactly2; distribution>=3. Some simple mutations rely on native Undo rather than compensating rollback..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:91.

## CMD-DEMO-top

### Top

**Business purpose:** Arrange selected native objects using top semantics..

**UI entry:** NexWave > Layout > Arrange > top.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools.Arrange.

**Execution workflow:** Resolve current shapes; group/ungroup, swap exactly two, distribute at least three, or apply selected native property relative to first shape..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools.Arrange.

**Data structures:** Native ShapeRange; coordinates/size/ZOrder.

**Failure behavior:** Select native objects first. Swap exactly2; distribution>=3. Some simple mutations rely on native Undo rather than compensating rollback..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:91.

## CMD-DEMO-middle

### Middle

**Business purpose:** Arrange selected native objects using middle semantics..

**UI entry:** NexWave > Layout > Arrange > middle.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools.Arrange.

**Execution workflow:** Resolve current shapes; group/ungroup, swap exactly two, distribute at least three, or apply selected native property relative to first shape..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools.Arrange.

**Data structures:** Native ShapeRange; coordinates/size/ZOrder.

**Failure behavior:** Select native objects first. Swap exactly2; distribution>=3. Some simple mutations rely on native Undo rather than compensating rollback..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:91.

## CMD-DEMO-bottom

### Bottom

**Business purpose:** Arrange selected native objects using bottom semantics..

**UI entry:** NexWave > Layout > Arrange > bottom.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools.Arrange.

**Execution workflow:** Resolve current shapes; group/ungroup, swap exactly two, distribute at least three, or apply selected native property relative to first shape..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools.Arrange.

**Data structures:** Native ShapeRange; coordinates/size/ZOrder.

**Failure behavior:** Select native objects first. Swap exactly2; distribution>=3. Some simple mutations rely on native Undo rather than compensating rollback..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:91.

## CMD-DEMO-width

### Width

**Business purpose:** Arrange selected native objects using width semantics..

**UI entry:** NexWave > Layout > Arrange > width.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools.Arrange.

**Execution workflow:** Resolve current shapes; group/ungroup, swap exactly two, distribute at least three, or apply selected native property relative to first shape..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools.Arrange.

**Data structures:** Native ShapeRange; coordinates/size/ZOrder.

**Failure behavior:** Select native objects first. Swap exactly2; distribution>=3. Some simple mutations rely on native Undo rather than compensating rollback..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:91.

## CMD-DEMO-height

### Height

**Business purpose:** Arrange selected native objects using height semantics..

**UI entry:** NexWave > Layout > Arrange > height.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools.Arrange.

**Execution workflow:** Resolve current shapes; group/ungroup, swap exactly two, distribute at least three, or apply selected native property relative to first shape..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools.Arrange.

**Data structures:** Native ShapeRange; coordinates/size/ZOrder.

**Failure behavior:** Select native objects first. Swap exactly2; distribution>=3. Some simple mutations rely on native Undo rather than compensating rollback..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:91.

## CMD-DEMO-size

### Size

**Business purpose:** Arrange selected native objects using size semantics..

**UI entry:** NexWave > Layout > Arrange > size.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools.Arrange.

**Execution workflow:** Resolve current shapes; group/ungroup, swap exactly two, distribute at least three, or apply selected native property relative to first shape..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools.Arrange.

**Data structures:** Native ShapeRange; coordinates/size/ZOrder.

**Failure behavior:** Select native objects first. Swap exactly2; distribution>=3. Some simple mutations rely on native Undo rather than compensating rollback..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:91.

## CMD-DEMO-horizontal

### Horizontal

**Business purpose:** Arrange selected native objects using horizontal semantics..

**UI entry:** NexWave > Layout > Arrange > horizontal.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools.Arrange.

**Execution workflow:** Resolve current shapes; group/ungroup, swap exactly two, distribute at least three, or apply selected native property relative to first shape..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools.Arrange.

**Data structures:** Native ShapeRange; coordinates/size/ZOrder.

**Failure behavior:** Select native objects first. Swap exactly2; distribution>=3. Some simple mutations rely on native Undo rather than compensating rollback..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:91.

## CMD-DEMO-vertical

### Vertical

**Business purpose:** Arrange selected native objects using vertical semantics..

**UI entry:** NexWave > Layout > Arrange > vertical.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools.Arrange.

**Execution workflow:** Resolve current shapes; group/ungroup, swap exactly two, distribute at least three, or apply selected native property relative to first shape..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools.Arrange.

**Data structures:** Native ShapeRange; coordinates/size/ZOrder.

**Failure behavior:** Select native objects first. Swap exactly2; distribution>=3. Some simple mutations rely on native Undo rather than compensating rollback..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:91.

## CMD-DEMO-swap

### Swap

**Business purpose:** Arrange selected native objects using swap semantics..

**UI entry:** NexWave > Layout > Arrange > swap.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools.Arrange.

**Execution workflow:** Resolve current shapes; group/ungroup, swap exactly two, distribute at least three, or apply selected native property relative to first shape..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools.Arrange.

**Data structures:** Native ShapeRange; coordinates/size/ZOrder.

**Failure behavior:** Select native objects first. Swap exactly2; distribution>=3. Some simple mutations rely on native Undo rather than compensating rollback..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:91.

## CMD-DEMO-scale-up

### Scale Up

**Business purpose:** Arrange selected native objects using scale-up semantics..

**UI entry:** NexWave > Layout > Arrange > scale-up.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools.Arrange.

**Execution workflow:** Resolve current shapes; group/ungroup, swap exactly two, distribute at least three, or apply selected native property relative to first shape..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools.Arrange.

**Data structures:** Native ShapeRange; coordinates/size/ZOrder.

**Failure behavior:** Select native objects first. Swap exactly2; distribution>=3. Some simple mutations rely on native Undo rather than compensating rollback..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:91.

## CMD-DEMO-scale-down

### Scale Down

**Business purpose:** Arrange selected native objects using scale-down semantics..

**UI entry:** NexWave > Layout > Arrange > scale-down.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools.Arrange.

**Execution workflow:** Resolve current shapes; group/ungroup, swap exactly two, distribute at least three, or apply selected native property relative to first shape..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools.Arrange.

**Data structures:** Native ShapeRange; coordinates/size/ZOrder.

**Failure behavior:** Select native objects first. Swap exactly2; distribution>=3. Some simple mutations rely on native Undo rather than compensating rollback..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:91.

## CMD-DEMO-rotate

### Rotate

**Business purpose:** Arrange selected native objects using rotate semantics..

**UI entry:** NexWave > Layout > Arrange > rotate.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools.Arrange.

**Execution workflow:** Resolve current shapes; group/ungroup, swap exactly two, distribute at least three, or apply selected native property relative to first shape..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools.Arrange.

**Data structures:** Native ShapeRange; coordinates/size/ZOrder.

**Failure behavior:** Select native objects first. Swap exactly2; distribution>=3. Some simple mutations rely on native Undo rather than compensating rollback..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:91.

## CMD-DEMO-flip-horizontal

### Flip Horizontal

**Business purpose:** Arrange selected native objects using flip-horizontal semantics..

**UI entry:** NexWave > Layout > Arrange > flip-horizontal.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools.Arrange.

**Execution workflow:** Resolve current shapes; group/ungroup, swap exactly two, distribute at least three, or apply selected native property relative to first shape..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools.Arrange.

**Data structures:** Native ShapeRange; coordinates/size/ZOrder.

**Failure behavior:** Select native objects first. Swap exactly2; distribution>=3. Some simple mutations rely on native Undo rather than compensating rollback..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:91.

## CMD-DEMO-flip-vertical

### Flip Vertical

**Business purpose:** Arrange selected native objects using flip-vertical semantics..

**UI entry:** NexWave > Layout > Arrange > flip-vertical.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools.Arrange.

**Execution workflow:** Resolve current shapes; group/ungroup, swap exactly two, distribute at least three, or apply selected native property relative to first shape..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools.Arrange.

**Data structures:** Native ShapeRange; coordinates/size/ZOrder.

**Failure behavior:** Select native objects first. Swap exactly2; distribution>=3. Some simple mutations rely on native Undo rather than compensating rollback..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:91.

## CMD-DEMO-duplicate

### Duplicate

**Business purpose:** Arrange selected native objects using duplicate semantics..

**UI entry:** NexWave > Layout > Arrange > duplicate.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools.Arrange.

**Execution workflow:** Resolve current shapes; group/ungroup, swap exactly two, distribute at least three, or apply selected native property relative to first shape..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools.Arrange.

**Data structures:** Native ShapeRange; coordinates/size/ZOrder.

**Failure behavior:** Select native objects first. Swap exactly2; distribution>=3. Some simple mutations rely on native Undo rather than compensating rollback..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:91.

## CMD-DEMO-group

### Group

**Business purpose:** Arrange selected native objects using group semantics..

**UI entry:** NexWave > Layout > Arrange > group.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools.Arrange.

**Execution workflow:** Resolve current shapes; group/ungroup, swap exactly two, distribute at least three, or apply selected native property relative to first shape..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools.Arrange.

**Data structures:** Native ShapeRange; coordinates/size/ZOrder.

**Failure behavior:** Select native objects first. Swap exactly2; distribution>=3. Some simple mutations rely on native Undo rather than compensating rollback..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:91.

## CMD-DEMO-ungroup

### Ungroup

**Business purpose:** Arrange selected native objects using ungroup semantics..

**UI entry:** NexWave > Layout > Arrange > ungroup.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools.Arrange.

**Execution workflow:** Resolve current shapes; group/ungroup, swap exactly two, distribute at least three, or apply selected native property relative to first shape..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools.Arrange.

**Data structures:** Native ShapeRange; coordinates/size/ZOrder.

**Failure behavior:** Select native objects first. Swap exactly2; distribution>=3. Some simple mutations rely on native Undo rather than compensating rollback..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:91.

## CMD-DEMO-front

### Front

**Business purpose:** Arrange selected native objects using front semantics..

**UI entry:** NexWave > Layout > Arrange > front.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools.Arrange.

**Execution workflow:** Resolve current shapes; group/ungroup, swap exactly two, distribute at least three, or apply selected native property relative to first shape..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools.Arrange.

**Data structures:** Native ShapeRange; coordinates/size/ZOrder.

**Failure behavior:** Select native objects first. Swap exactly2; distribution>=3. Some simple mutations rely on native Undo rather than compensating rollback..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:91.

## CMD-DEMO-back

### Back

**Business purpose:** Arrange selected native objects using back semantics..

**UI entry:** NexWave > Layout > Arrange > back.

**Implementation files:** apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools.Arrange.

**Execution workflow:** Resolve current shapes; group/ungroup, swap exactly two, distribute at least three, or apply selected native property relative to first shape..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools.Arrange.

**Data structures:** Native ShapeRange; coordinates/size/ZOrder.

**Failure behavior:** Select native objects first. Swap exactly2; distribution>=3. Some simple mutations rely on native Undo rather than compensating rollback..

**Relevant tests (source only):** tests/s0/DemoScenario.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoTools.cs:91.

## CMD-ELEMENT-pentagon

### Right pentagon

**Business purpose:** Invoke the supported Right pentagon local operation..

**UI entry:** NexWave > Elements > Element Controls > Right pentagon.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ElementCommands.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:59.

## CMD-ELEMENT-chevron

### Chevron

**Business purpose:** Invoke the supported Chevron local operation..

**UI entry:** NexWave > Elements > Element Controls > Chevron.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ElementCommands.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:59.

## CMD-ELEMENT-connector

### Universal connector

**Business purpose:** Invoke the supported Universal connector local operation..

**UI entry:** NexWave > Elements > Element Controls > Universal connector.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/UniversalConnectors.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:59.

## CMD-ELEMENT-edit

### Edit selected element

**Business purpose:** Invoke the supported Edit selected element local operation..

**UI entry:** NexWave > Elements > Element Controls > Edit selected element.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ElementCommands.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:60.

## CMD-ELEMENT-fit

### Fit text to unlocked edges

**Business purpose:** Invoke the supported Fit text to unlocked edges local operation..

**UI entry:** NexWave > Elements > Element Controls > Fit text to unlocked edges.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ElementCommands.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:60.

## CMD-ELEMENT-convert-text

### Convert to Text Box

**Business purpose:** Invoke the supported Convert to Text Box local operation..

**UI entry:** NexWave > Elements > Element Controls > Convert to Text Box.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ElementCommands.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:61.

## CMD-ELEMENT-convert-pentagon

### Convert to Right Pentagon

**Business purpose:** Invoke the supported Convert to Right Pentagon local operation..

**UI entry:** NexWave > Elements > Element Controls > Convert to Right Pentagon.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ElementCommands.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:61.

## CMD-ELEMENT-convert-chevron

### Convert to Chevron

**Business purpose:** Invoke the supported Convert to Chevron local operation..

**UI entry:** NexWave > Elements > Element Controls > Convert to Chevron.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ElementCommands.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:61.

## CMD-ELEMENT-convert-rounded

### Convert to Rounded Rectangle

**Business purpose:** Invoke the supported Convert to Rounded Rectangle local operation..

**UI entry:** NexWave > Elements > Element Controls > Convert to Rounded Rectangle.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ElementCommands.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:61.

## CMD-ELEMENT-duplicate-left

### Duplicate left

**Business purpose:** Invoke the supported Duplicate left local operation..

**UI entry:** NexWave > Elements > Element Controls > Duplicate left.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ElementCommands.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:62.

## CMD-ELEMENT-duplicate-right

### Duplicate right

**Business purpose:** Invoke the supported Duplicate right local operation..

**UI entry:** NexWave > Elements > Element Controls > Duplicate right.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ElementCommands.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:62.

## CMD-ELEMENT-duplicate-up

### Duplicate above

**Business purpose:** Invoke the supported Duplicate above local operation..

**UI entry:** NexWave > Elements > Element Controls > Duplicate above.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ElementCommands.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:62.

## CMD-ELEMENT-duplicate-down

### Duplicate below

**Business purpose:** Invoke the supported Duplicate below local operation..

**UI entry:** NexWave > Elements > Element Controls > Duplicate below.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ElementCommands.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:62.

## CMD-ELEMENT-indicator-next

### Increase / next indicator state

**Business purpose:** Invoke the supported Increase / next indicator state local operation..

**UI entry:** NexWave > Elements > Element Controls > Increase / next indicator state.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/IndicatorElements.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:63.

## CMD-ELEMENT-indicator-previous

### Decrease / previous state

**Business purpose:** Invoke the supported Decrease / previous state local operation..

**UI entry:** NexWave > Elements > Element Controls > Decrease / previous state.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/IndicatorElements.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:63.

## CMD-ELEMENT-indicator-keys

### Indicator keyboard controls

**Business purpose:** Invoke the supported Indicator keyboard controls local operation..

**UI entry:** NexWave > Elements > Element Controls > Indicator keyboard controls.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/IndicatorElements.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:63.

## CMD-ELEMENT-convert-harvey

### Convert indicator to Harvey ball

**Business purpose:** Invoke the supported Convert indicator to Harvey ball local operation..

**UI entry:** NexWave > Elements > Element Controls > Convert indicator to Harvey ball.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/IndicatorElements.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:63.

## CMD-ELEMENT-convert-checkbox

### Convert indicator to Checkbox

**Business purpose:** Invoke the supported Convert indicator to Checkbox local operation..

**UI entry:** NexWave > Elements > Element Controls > Convert indicator to Checkbox.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/IndicatorElements.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:63.

## CMD-ELEMENT-grid-create

### New process grid

**Business purpose:** Invoke the supported New process grid local operation..

**UI entry:** NexWave > Elements > Element Controls > New process grid.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ProcessFlowTools.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:64.

## CMD-ELEMENT-grid-row

### Add grid row

**Business purpose:** Invoke the supported Add grid row local operation..

**UI entry:** NexWave > Elements > Element Controls > Add grid row.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ProcessFlowTools.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:64.

## CMD-ELEMENT-grid-column

### Add grid column

**Business purpose:** Invoke the supported Add grid column local operation..

**UI entry:** NexWave > Elements > Element Controls > Add grid column.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ProcessFlowTools.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:64.

## CMD-ELEMENT-grid-remove-row

### Hide grid row

**Business purpose:** Invoke the supported Hide grid row local operation..

**UI entry:** NexWave > Elements > Element Controls > Hide grid row.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ProcessFlowTools.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:64.

## CMD-ELEMENT-grid-remove-column

### Hide grid column

**Business purpose:** Invoke the supported Hide grid column local operation..

**UI entry:** NexWave > Elements > Element Controls > Hide grid column.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ProcessFlowTools.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:64.

## CMD-ELEMENT-grid-left

### Move grid column left

**Business purpose:** Invoke the supported Move grid column left local operation..

**UI entry:** NexWave > Elements > Element Controls > Move grid column left.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ProcessFlowTools.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:64.

## CMD-ELEMENT-grid-right

### Move grid column right

**Business purpose:** Invoke the supported Move grid column right local operation..

**UI entry:** NexWave > Elements > Element Controls > Move grid column right.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ProcessFlowTools.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:64.

## CMD-ELEMENT-grid-stripe-rows

### Stripe grid rows

**Business purpose:** Invoke the supported Stripe grid rows local operation..

**UI entry:** NexWave > Elements > Element Controls > Stripe grid rows.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ProcessFlowTools.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:64.

## CMD-ELEMENT-grid-stripe-columns

### Stripe grid columns

**Business purpose:** Invoke the supported Stripe grid columns local operation..

**UI entry:** NexWave > Elements > Element Controls > Stripe grid columns.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ProcessFlowTools.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:64.

## CMD-ELEMENT-grid-resize

### Resize process grid

**Business purpose:** Invoke the supported Resize process grid local operation..

**UI entry:** NexWave > Elements > Element Controls > Resize process grid.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ProcessFlowTools.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:64.

## CMD-ELEMENT-refresh

### Refresh element geometry

**Business purpose:** Invoke the supported Refresh element geometry local operation..

**UI entry:** NexWave > Elements > Element Controls > Refresh element geometry.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ElementCommands.cs.

**Responsible classes:** ElementCommands.Run.

**Execution workflow:** Ribbon tag reaches ElementCommands.Run; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ElementCommands.Run.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/TextShapeElementTests.cs:1; tests/s0/ProcessFlowTests.cs:1; tests/s0/IndicatorElementTests.cs:1; tests/s0/UniversalConnectorTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:64.

## CMD-TOOLS-symbol-picker

### Insert symbol…

**Business purpose:** Invoke the supported Insert symbol… local operation..

**UI entry:** NexWave > Elements > Tools > Insert symbol….

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ProductivityTools.cs.

**Responsible classes:** ThisAddIn.RunProductivity.

**Execution workflow:** Ribbon tag reaches ThisAddIn.RunProductivity; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ThisAddIn.RunProductivity.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/ProductivityToolTests.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:52.

## CMD-TOOLS-decimal-safe

### Switch decimal symbols…

**Business purpose:** Invoke the supported Switch decimal symbols… local operation..

**UI entry:** NexWave > Elements > Tools > Switch decimal symbols….

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ProductivityTools.cs.

**Responsible classes:** ThisAddIn.RunProductivity.

**Execution workflow:** Ribbon tag reaches ThisAddIn.RunProductivity; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ThisAddIn.RunProductivity.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/ProductivityToolTests.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:52.

## CMD-TOOLS-raster-copy

### Create picture copy (retain original)

**Business purpose:** Invoke the supported Create picture copy (retain original) local operation..

**UI entry:** NexWave > Elements > Tools > Create picture copy (retain original).

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ProductivityTools.cs.

**Responsible classes:** ThisAddIn.RunProductivity.

**Execution workflow:** Ribbon tag reaches ThisAddIn.RunProductivity; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ThisAddIn.RunProductivity.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/ProductivityToolTests.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:52.

## CMD-TOOLS-chart-workbooks

### Chart workbooks…

**Business purpose:** Invoke the supported Chart workbooks… local operation..

**UI entry:** NexWave > Elements > Tools > Chart workbooks….

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ProductivityTools.cs.

**Responsible classes:** ThisAddIn.RunProductivity.

**Execution workflow:** Ribbon tag reaches ThisAddIn.RunProductivity; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ThisAddIn.RunProductivity.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/ProductivityToolTests.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:52.

## CMD-TOOLS-data-links

### Data links…

**Business purpose:** Invoke the supported Data links… local operation..

**UI entry:** NexWave > Elements > Tools > Data links….

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ProductivityTools.cs.

**Responsible classes:** ThisAddIn.RunProductivity.

**Execution workflow:** Ribbon tag reaches ThisAddIn.RunProductivity; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ThisAddIn.RunProductivity.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/ProductivityToolTests.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:52.

## CMD-TOOLS-save-selected

### Save selected slides…

**Business purpose:** Invoke the supported Save selected slides… local operation..

**UI entry:** NexWave > Elements > Tools > Save selected slides….

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ProductivityTools.cs.

**Responsible classes:** ThisAddIn.RunProductivity.

**Execution workflow:** Ribbon tag reaches ThisAddIn.RunProductivity; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ThisAddIn.RunProductivity.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/ProductivityToolTests.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:52.

## CMD-TOOLS-selected-style

### Apply selected-object style…

**Business purpose:** Invoke the supported Apply selected-object style… local operation..

**UI entry:** NexWave > Elements > Tools > Apply selected-object style….

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ProductivityTools.cs.

**Responsible classes:** ThisAddIn.RunProductivity.

**Execution workflow:** Ribbon tag reaches ThisAddIn.RunProductivity; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ThisAddIn.RunProductivity.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/ProductivityToolTests.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:52.

## CMD-TOOLS-review

### Review presentation…

**Business purpose:** Invoke the supported Review presentation… local operation..

**UI entry:** NexWave > Elements > Tools > Review presentation….

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ProductivityTools.cs.

**Responsible classes:** ThisAddIn.RunProductivity.

**Execution workflow:** Ribbon tag reaches ThisAddIn.RunProductivity; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ThisAddIn.RunProductivity.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/ProductivityToolTests.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:52.

## CMD-TOOLS-help

### Element tools help

**Business purpose:** Invoke the supported Element tools help local operation..

**UI entry:** NexWave > Elements > Tools > Element tools help.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ProductivityTools.cs.

**Responsible classes:** ThisAddIn.RunProductivity.

**Execution workflow:** Ribbon tag reaches ThisAddIn.RunProductivity; validate live selection and owned element, then call its typed adapter; grid removal asks confirmation and retains inactive history..

**Dependencies:** PowerPoint/VSTO/WinForms; ThisAddIn.RunProductivity.

**Data structures:** Managed element spec/XML or selected native objects.

**Failure behavior:** Enabled/ownership rules restrict incompatible selection; malformed or stale metadata preserved. Help/aliases are UI routing, not separate data engines..

**Relevant tests (source only):** tests/s0/ProductivityToolTests.cs:1; tests/s0/ProductivityToolOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:52.

## CMD-WORKFLOW-agenda-insert

### Chapter

**Business purpose:** Apply Chapter within the selected presentation..

**UI entry:** NexWave ribbon workflow > Chapter.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ChapterAgenda.cs.

**Responsible classes:** ChapterAgenda.InsertChapter.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes ChapterAgenda.InsertChapter..

**Dependencies:** PowerPoint/VSTO/WinForms; ChapterAgenda.InsertChapter.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:21.

## CMD-WORKFLOW-agenda-contents

### Table of contents

**Business purpose:** Apply Table of contents within the selected presentation..

**UI entry:** NexWave ribbon workflow > Table of contents.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ChapterAgenda.cs.

**Responsible classes:** ChapterAgenda.InsertContents.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes ChapterAgenda.InsertContents..

**Dependencies:** PowerPoint/VSTO/WinForms; ChapterAgenda.InsertContents.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:21.

## CMD-WORKFLOW-agenda-settings

### Manage Agenda

**Business purpose:** Apply Manage Agenda within the selected presentation..

**UI entry:** NexWave ribbon workflow > Manage Agenda.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ChapterAgenda.cs.

**Responsible classes:** ChapterAgenda.ShowManager.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes ChapterAgenda.ShowManager..

**Dependencies:** PowerPoint/VSTO/WinForms; ChapterAgenda.ShowManager.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:21.

## CMD-WORKFLOW-agenda-split

### Split agenda here

**Business purpose:** Apply Split agenda here within the selected presentation..

**UI entry:** NexWave ribbon workflow > Split agenda here.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ChapterAgenda.cs.

**Responsible classes:** ChapterAgenda.SplitHere.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes ChapterAgenda.SplitHere..

**Dependencies:** PowerPoint/VSTO/WinForms; ChapterAgenda.SplitHere.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:21.

## CMD-WORKFLOW-agenda-refresh

### Update agenda

**Business purpose:** Apply Update agenda within the selected presentation..

**UI entry:** NexWave ribbon workflow > Update agenda.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ChapterAgenda.cs.

**Responsible classes:** ChapterAgenda.RefreshAll / PresentationAgenda.RefreshSelected.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes ChapterAgenda.RefreshAll / PresentationAgenda.RefreshSelected..

**Dependencies:** PowerPoint/VSTO/WinForms; ChapterAgenda.RefreshAll / PresentationAgenda.RefreshSelected.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:21.

## CMD-WORKFLOW-percent-area

### 100% stacked area

**Business purpose:** Apply 100% stacked area within the selected presentation..

**UI entry:** NexWave ribbon workflow > 100% stacked area.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/DemoTools.cs.

**Responsible classes:** DemoTools.InsertChart.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes DemoTools.InsertChart..

**Dependencies:** PowerPoint/VSTO/WinForms; DemoTools.InsertChart.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:29.

## CMD-WORKFLOW-series-settings

### Series type and axis

**Business purpose:** Apply Series type and axis within the selected presentation..

**UI entry:** NexWave ribbon workflow > Series type and axis.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/NativeSeriesControls.cs.

**Responsible classes:** NativeSeriesControls.Show.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes NativeSeriesControls.Show..

**Dependencies:** PowerPoint/VSTO/WinForms; NativeSeriesControls.Show.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:29.

## CMD-WORKFLOW-axis-settings

### Value axis settings

**Business purpose:** Apply Value axis settings within the selected presentation..

**UI entry:** NexWave ribbon workflow > Value axis settings.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/NativeChartControls.cs.

**Responsible classes:** NativeChartControls.ShowAxis.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes NativeChartControls.ShowAxis..

**Dependencies:** PowerPoint/VSTO/WinForms; NativeChartControls.ShowAxis.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:29.

## CMD-WORKFLOW-match-scales

### Match value scales

**Business purpose:** Apply Match value scales within the selected presentation..

**UI entry:** NexWave ribbon workflow > Match value scales.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/NativeChartControls.cs.

**Responsible classes:** NativeChartControls.MatchScales.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes NativeChartControls.MatchScales..

**Dependencies:** PowerPoint/VSTO/WinForms; NativeChartControls.MatchScales.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:29.

## CMD-WORKFLOW-annotation-refresh

### Refresh selected annotation

**Business purpose:** Apply Refresh selected annotation within the selected presentation..

**UI entry:** NexWave ribbon workflow > Refresh selected annotation.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ChartAnnotations.cs.

**Responsible classes:** ChartAnnotations.RefreshSelected.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes ChartAnnotations.RefreshSelected..

**Dependencies:** PowerPoint/VSTO/WinForms; ChartAnnotations.RefreshSelected.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:29.

## CMD-WORKFLOW-connections

### Configure services

**Business purpose:** Apply Configure services within the selected presentation..

**UI entry:** NexWave ribbon workflow > Configure services.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ServiceConnections.cs.

**Responsible classes:** ServiceConnections.ShowConfiguration.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes ServiceConnections.ShowConfiguration..

**Dependencies:** PowerPoint/VSTO/WinForms; ServiceConnections.ShowConfiguration.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Partial: nonsecret configuration only; no live service execution.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:30.

## CMD-WORKFLOW-relink-selected

### Relink selected chart

**Business purpose:** Apply Relink selected chart within the selected presentation..

**UI entry:** NexWave ribbon workflow > Relink selected chart.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/MappedLinkOperations.cs.

**Responsible classes:** MappedLinkOperations.RelinkSelected.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes MappedLinkOperations.RelinkSelected..

**Dependencies:** PowerPoint/VSTO/WinForms; MappedLinkOperations.RelinkSelected.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:35.

## CMD-WORKFLOW-map-excel

### Link Excel Range

**Business purpose:** Apply Link Excel Range within the selected presentation..

**UI entry:** NexWave ribbon workflow > Link Excel Range.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/ExcelRangeMapping.cs.

**Responsible classes:** MappedExcelData.Choose.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes MappedExcelData.Choose..

**Dependencies:** PowerPoint/VSTO/WinForms; MappedExcelData.Choose.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:35.

## CMD-WORKFLOW-data-manager

### Data Link Manager

**Business purpose:** Apply Data Link Manager within the selected presentation..

**UI entry:** NexWave ribbon workflow > Data Link Manager.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/DataLinkManager.cs.

**Responsible classes:** DataLinkManager.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes DataLinkManager..

**Dependencies:** PowerPoint/VSTO/WinForms; DataLinkManager.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:35.

## CMD-WORKFLOW-quality-review

### Presentation review

**Business purpose:** Apply Presentation review within the selected presentation..

**UI entry:** NexWave ribbon workflow > Presentation review.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/PresentationQuality.cs.

**Responsible classes:** PresentationQuality.ShowInspection.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes PresentationQuality.ShowInspection..

**Dependencies:** PowerPoint/VSTO/WinForms; PresentationQuality.ShowInspection.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:41.

## CMD-WORKFLOW-export-pdf

### Export PDF

**Business purpose:** Apply Export PDF within the selected presentation..

**UI entry:** NexWave ribbon workflow > Export PDF.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/NativeSeriesControls.cs.

**Responsible classes:** NativePdfExport.Show.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes NativePdfExport.Show..

**Dependencies:** PowerPoint/VSTO/WinForms; NativePdfExport.Show.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:42.

## CMD-WORKFLOW-text-left

### Text left

**Business purpose:** Apply Text left within the selected presentation..

**UI entry:** NexWave ribbon workflow > Text left.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/NativeFormatting.cs.

**Responsible classes:** NativeFormatting.Text.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes NativeFormatting.Text..

**Dependencies:** PowerPoint/VSTO/WinForms; NativeFormatting.Text.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:37.

## CMD-WORKFLOW-text-center

### Text center

**Business purpose:** Apply Text center within the selected presentation..

**UI entry:** NexWave ribbon workflow > Text center.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/NativeFormatting.cs.

**Responsible classes:** NativeFormatting.Text.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes NativeFormatting.Text..

**Dependencies:** PowerPoint/VSTO/WinForms; NativeFormatting.Text.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:37.

## CMD-WORKFLOW-text-right

### Text right

**Business purpose:** Apply Text right within the selected presentation..

**UI entry:** NexWave ribbon workflow > Text right.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/NativeFormatting.cs.

**Responsible classes:** NativeFormatting.Text.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes NativeFormatting.Text..

**Dependencies:** PowerPoint/VSTO/WinForms; NativeFormatting.Text.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:37.

## CMD-WORKFLOW-text-justify

### Text justify

**Business purpose:** Apply Text justify within the selected presentation..

**UI entry:** NexWave ribbon workflow > Text justify.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/NativeFormatting.cs.

**Responsible classes:** NativeFormatting.Text.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes NativeFormatting.Text..

**Dependencies:** PowerPoint/VSTO/WinForms; NativeFormatting.Text.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:37.

## CMD-WORKFLOW-text-top

### Text top

**Business purpose:** Apply Text top within the selected presentation..

**UI entry:** NexWave ribbon workflow > Text top.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/NativeFormatting.cs.

**Responsible classes:** NativeFormatting.Text.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes NativeFormatting.Text..

**Dependencies:** PowerPoint/VSTO/WinForms; NativeFormatting.Text.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:37.

## CMD-WORKFLOW-text-middle

### Text middle

**Business purpose:** Apply Text middle within the selected presentation..

**UI entry:** NexWave ribbon workflow > Text middle.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/NativeFormatting.cs.

**Responsible classes:** NativeFormatting.Text.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes NativeFormatting.Text..

**Dependencies:** PowerPoint/VSTO/WinForms; NativeFormatting.Text.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:37.

## CMD-WORKFLOW-text-bottom

### Text bottom

**Business purpose:** Apply Text bottom within the selected presentation..

**UI entry:** NexWave ribbon workflow > Text bottom.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/NativeFormatting.cs.

**Responsible classes:** NativeFormatting.Text.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes NativeFormatting.Text..

**Dependencies:** PowerPoint/VSTO/WinForms; NativeFormatting.Text.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:37.

## CMD-WORKFLOW-text-horizontal

### Text horizontal

**Business purpose:** Apply Text horizontal within the selected presentation..

**UI entry:** NexWave ribbon workflow > Text horizontal.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/NativeFormatting.cs.

**Responsible classes:** NativeFormatting.Text.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes NativeFormatting.Text..

**Dependencies:** PowerPoint/VSTO/WinForms; NativeFormatting.Text.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:37.

## CMD-WORKFLOW-text-downward

### Text downward

**Business purpose:** Apply Text downward within the selected presentation..

**UI entry:** NexWave ribbon workflow > Text downward.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/NativeFormatting.cs.

**Responsible classes:** NativeFormatting.Text.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes NativeFormatting.Text..

**Dependencies:** PowerPoint/VSTO/WinForms; NativeFormatting.Text.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:37.

## CMD-WORKFLOW-text-upward

### Text upward

**Business purpose:** Apply Text upward within the selected presentation..

**UI entry:** NexWave ribbon workflow > Text upward.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/NativeFormatting.cs.

**Responsible classes:** NativeFormatting.Text.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes NativeFormatting.Text..

**Dependencies:** PowerPoint/VSTO/WinForms; NativeFormatting.Text.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:37.

## CMD-WORKFLOW-color-fill

### Choose fill color

**Business purpose:** Apply Choose fill color within the selected presentation..

**UI entry:** NexWave ribbon workflow > Choose fill color.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/NativeFormatting.cs.

**Responsible classes:** NativeFormatting.Color.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes NativeFormatting.Color..

**Dependencies:** PowerPoint/VSTO/WinForms; NativeFormatting.Color.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:40.

## CMD-WORKFLOW-color-outline

### Choose outline color

**Business purpose:** Apply Choose outline color within the selected presentation..

**UI entry:** NexWave ribbon workflow > Choose outline color.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/NativeFormatting.cs.

**Responsible classes:** NativeFormatting.Color.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes NativeFormatting.Color..

**Dependencies:** PowerPoint/VSTO/WinForms; NativeFormatting.Color.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:40.

## CMD-WORKFLOW-color-text

### Choose text color

**Business purpose:** Apply Choose text color within the selected presentation..

**UI entry:** NexWave ribbon workflow > Choose text color.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/NativeFormatting.cs.

**Responsible classes:** NativeFormatting.Color.

**Execution workflow:** WorkflowEnabled gates active document/selection; RunWorkflow enters the shared mutation boundary (Undo is adapter-specific) and invokes NativeFormatting.Color..

**Dependencies:** PowerPoint/VSTO/WinForms; NativeFormatting.Color.

**Data structures:** Typed module DTOs and current native selection.

**Failure behavior:** Specific module capability/preflight applies; service configuration does not test/authenticate/connect to endpoints..

**Relevant tests (source only):** tests/s0/WorkflowOfficeTests.cs:1; tests/s0/NativeFormattingTests.cs:1; tests/s0/ChapterAgendaTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:40.

## CMD-SOURCE-excel

### Excel sheet

**Business purpose:** Choose and persist a .xlsx/.xls/.xlsm/.xlsb source reference..

**UI entry:** NexWave > Data Source > Excel sheet.

**Implementation files:** apps/powerpoint-vsto/DataSources.cs; apps/powerpoint-vsto/ThisAddIn.cs.

**Responsible classes:** DataSources.ChooseAndLink; DataSources.Save.

**Execution workflow:** Open extension-filtered picker; validate selected path/format; save or reuse presentation registry entry; selection does not parse data yet..

**Dependencies:** PowerPoint/VSTO/WinForms; DataSources.ChooseAndLink; DataSources.Save.

**Data structures:** DataSourceLink v1/v2.

**Failure behavior:** Existing path, supported extension/canonical path and bounded metadata required; Cancel leaves document unchanged..

**Relevant tests (source only):** tests/s0/DataSourceTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DataSources.cs:267.

## CMD-SOURCE-csv

### CSV file

**Business purpose:** Choose and persist a .csv source reference..

**UI entry:** NexWave > Data Source > CSV file.

**Implementation files:** apps/powerpoint-vsto/DataSources.cs; apps/powerpoint-vsto/ThisAddIn.cs.

**Responsible classes:** DataSources.ChooseAndLink; DataSources.Save.

**Execution workflow:** Open extension-filtered picker; validate selected path/format; save or reuse presentation registry entry; selection does not parse data yet..

**Dependencies:** PowerPoint/VSTO/WinForms; DataSources.ChooseAndLink; DataSources.Save.

**Data structures:** DataSourceLink v1/v2.

**Failure behavior:** Existing path, supported extension/canonical path and bounded metadata required; Cancel leaves document unchanged..

**Relevant tests (source only):** tests/s0/DataSourceTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DataSources.cs:268.

## CMD-SOURCE-json

### JSON file

**Business purpose:** Choose and persist a .json source reference..

**UI entry:** NexWave > Data Source > JSON file.

**Implementation files:** apps/powerpoint-vsto/DataSources.cs; apps/powerpoint-vsto/ThisAddIn.cs.

**Responsible classes:** DataSources.ChooseAndLink; DataSources.Save.

**Execution workflow:** Open extension-filtered picker; validate selected path/format; save or reuse presentation registry entry; selection does not parse data yet..

**Dependencies:** PowerPoint/VSTO/WinForms; DataSources.ChooseAndLink; DataSources.Save.

**Data structures:** DataSourceLink v1/v2.

**Failure behavior:** Existing path, supported extension/canonical path and bounded metadata required; Cancel leaves document unchanged..

**Relevant tests (source only):** tests/s0/DataSourceTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DataSources.cs:269.

## CMD-SOURCE-xml

### XML file

**Business purpose:** Choose and persist a .xml source reference..

**UI entry:** NexWave > Data Source > XML file.

**Implementation files:** apps/powerpoint-vsto/DataSources.cs; apps/powerpoint-vsto/ThisAddIn.cs.

**Responsible classes:** DataSources.ChooseAndLink; DataSources.Save.

**Execution workflow:** Open extension-filtered picker; validate selected path/format; save or reuse presentation registry entry; selection does not parse data yet..

**Dependencies:** PowerPoint/VSTO/WinForms; DataSources.ChooseAndLink; DataSources.Save.

**Data structures:** DataSourceLink v1/v2.

**Failure behavior:** Existing path, supported extension/canonical path and bounded metadata required; Cancel leaves document unchanged..

**Relevant tests (source only):** tests/s0/DataSourceTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DataSources.cs:270.

## CMD-DIRECT-ppiTemplateLibrary

### Library

**Business purpose:** Open or apply Library for the current presentation..

**UI entry:** NexWave > Library.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/DemoRibbon.cs.

**Responsible classes:** IntelligenceRibbon.ShowTemplateLibrary.

**Execution workflow:** Office invokes ShowTemplateLibrary; ThisAddIn manages compatibility/current window/mutation; delegates to the corresponding module..

**Dependencies:** PowerPoint/VSTO/WinForms; IntelligenceRibbon.ShowTemplateLibrary.

**Data structures:** Current document/selection; source or template input where required.

**Failure behavior:** Compatibility/current presentation/selection and module validation; properties/library views have dedicated ownership/cleanup..

**Relevant tests (source only):** tests/s0/DemoContractTests.cs:1; tests/s3/RendererRibbonTests.cs:1; tests/compatibility/OfficeCompatibilityTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:24.

## CMD-DIRECT-ppiRenderShapes

### Editable shape columns

**Business purpose:** Open or apply Editable shape columns for the current presentation..

**UI entry:** NexWave > Editable shape columns.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/DemoRibbon.cs.

**Responsible classes:** IntelligenceRibbon.InsertRenderedShapes.

**Execution workflow:** Office invokes InsertRenderedShapes; ThisAddIn manages compatibility/current window/mutation; delegates to the corresponding module..

**Dependencies:** PowerPoint/VSTO/WinForms; IntelligenceRibbon.InsertRenderedShapes.

**Data structures:** Current document/selection; source or template input where required.

**Failure behavior:** Compatibility/current presentation/selection and module validation; properties/library views have dedicated ownership/cleanup..

**Relevant tests (source only):** tests/s0/DemoContractTests.cs:1; tests/s3/RendererRibbonTests.cs:1; tests/compatibility/OfficeCompatibilityTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:26.

## CMD-DIRECT-ppiInsert

### Linked column chart

**Business purpose:** Open or apply Linked column chart for the current presentation..

**UI entry:** NexWave > Linked column chart.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/DemoRibbon.cs.

**Responsible classes:** IntelligenceRibbon.InsertSample.

**Execution workflow:** Office invokes InsertSample; ThisAddIn manages compatibility/current window/mutation; delegates to the corresponding module..

**Dependencies:** PowerPoint/VSTO/WinForms; IntelligenceRibbon.InsertSample.

**Data structures:** Current document/selection; source or template input where required.

**Failure behavior:** Compatibility/current presentation/selection and module validation; properties/library views have dedicated ownership/cleanup..

**Relevant tests (source only):** tests/s0/DemoContractTests.cs:1; tests/s3/RendererRibbonTests.cs:1; tests/compatibility/OfficeCompatibilityTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:27.

## CMD-DIRECT-ppiContextProperties

### Properties

**Business purpose:** Open or apply Properties for the current presentation..

**UI entry:** NexWave > Properties.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/DemoRibbon.cs.

**Responsible classes:** IntelligenceRibbon.ShowContextProperties.

**Execution workflow:** Office invokes ShowContextProperties; ThisAddIn manages compatibility/current window/mutation; delegates to the corresponding module..

**Dependencies:** PowerPoint/VSTO/WinForms; IntelligenceRibbon.ShowContextProperties.

**Data structures:** Current document/selection; source or template input where required.

**Failure behavior:** Compatibility/current presentation/selection and module validation; properties/library views have dedicated ownership/cleanup..

**Relevant tests (source only):** tests/s0/DemoContractTests.cs:1; tests/s3/RendererRibbonTests.cs:1; tests/compatibility/OfficeCompatibilityTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:28.

## CMD-DIRECT-ppiLibraryLayout

### Layout Links

**Business purpose:** Open or apply Layout Links for the current presentation..

**UI entry:** NexWave > Layout Links.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/DemoRibbon.cs.

**Responsible classes:** IntelligenceRibbon.ShowLibraryLayout.

**Execution workflow:** Office invokes ShowLibraryLayout; ThisAddIn manages compatibility/current window/mutation; delegates to the corresponding module..

**Dependencies:** PowerPoint/VSTO/WinForms; IntelligenceRibbon.ShowLibraryLayout.

**Data structures:** Current document/selection; source or template input where required.

**Failure behavior:** Compatibility/current presentation/selection and module validation; properties/library views have dedicated ownership/cleanup..

**Relevant tests (source only):** tests/s0/DemoContractTests.cs:1; tests/s3/RendererRibbonTests.cs:1; tests/compatibility/OfficeCompatibilityTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:38.

## CMD-DIRECT-ppiTemplateAlignment

### Template Alignment

**Business purpose:** Open or apply Template Alignment for the current presentation..

**UI entry:** NexWave > Template Alignment.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/DemoRibbon.cs.

**Responsible classes:** IntelligenceRibbon.ChangeTemplateAlignment.

**Execution workflow:** Office invokes ChangeTemplateAlignment; ThisAddIn manages compatibility/current window/mutation; delegates to the corresponding module..

**Dependencies:** PowerPoint/VSTO/WinForms; IntelligenceRibbon.ChangeTemplateAlignment.

**Data structures:** Current document/selection; source or template input where required.

**Failure behavior:** Compatibility/current presentation/selection and module validation; properties/library views have dedicated ownership/cleanup..

**Relevant tests (source only):** tests/s0/DemoContractTests.cs:1; tests/s3/RendererRibbonTests.cs:1; tests/compatibility/OfficeCompatibilityTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:39.

## CMD-DIRECT-ppiTemplatePalette

### Template Colors

**Business purpose:** Open or apply Template Colors for the current presentation..

**UI entry:** NexWave > Template Colors.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/DemoRibbon.cs.

**Responsible classes:** IntelligenceRibbon.ChangeTemplatePalette.

**Execution workflow:** Office invokes ChangeTemplatePalette; ThisAddIn manages compatibility/current window/mutation; delegates to the corresponding module..

**Dependencies:** PowerPoint/VSTO/WinForms; IntelligenceRibbon.ChangeTemplatePalette.

**Data structures:** Current document/selection; source or template input where required.

**Failure behavior:** Compatibility/current presentation/selection and module validation; properties/library views have dedicated ownership/cleanup..

**Relevant tests (source only):** tests/s0/DemoContractTests.cs:1; tests/s3/RendererRibbonTests.cs:1; tests/compatibility/OfficeCompatibilityTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:40.

## CMD-DIRECT-ppiInspect

### Check & Sync

**Business purpose:** Open or apply Check & Sync for the current presentation..

**UI entry:** NexWave > Check & Sync.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/DemoRibbon.cs.

**Responsible classes:** IntelligenceRibbon.InspectChart.

**Execution workflow:** Office invokes InspectChart; ThisAddIn manages compatibility/current window/mutation; delegates to the corresponding module..

**Dependencies:** PowerPoint/VSTO/WinForms; IntelligenceRibbon.InspectChart.

**Data structures:** Current document/selection; source or template input where required.

**Failure behavior:** Compatibility/current presentation/selection and module validation; properties/library views have dedicated ownership/cleanup..

**Relevant tests (source only):** tests/s0/DemoContractTests.cs:1; tests/s3/RendererRibbonTests.cs:1; tests/compatibility/OfficeCompatibilityTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:41.

## CMD-DIRECT-ppiOpenWorkspace

### Open Workspace

**Business purpose:** Open or apply Open Workspace for the current presentation..

**UI entry:** NexWave > Open Workspace.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/DemoRibbon.cs.

**Responsible classes:** IntelligenceRibbon.ShowWorkspace.

**Execution workflow:** Office invokes ShowWorkspace; ThisAddIn manages compatibility/current window/mutation; delegates to the corresponding module..

**Dependencies:** PowerPoint/VSTO/WinForms; IntelligenceRibbon.ShowWorkspace.

**Data structures:** Current document/selection; source or template input where required.

**Failure behavior:** Compatibility/current presentation/selection and module validation; properties/library views have dedicated ownership/cleanup..

**Relevant tests (source only):** tests/s0/DemoContractTests.cs:1; tests/s3/RendererRibbonTests.cs:1; tests/compatibility/OfficeCompatibilityTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:44.

## CMD-DIRECT-ppiDemo

### Explore Demo Tools

**Business purpose:** Open or apply Explore Demo Tools for the current presentation..

**UI entry:** NexWave > Explore Demo Tools.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/DemoRibbon.cs.

**Responsible classes:** IntelligenceRibbon.ShowDemo.

**Execution workflow:** Office invokes ShowDemo; ThisAddIn manages compatibility/current window/mutation; delegates to the corresponding module..

**Dependencies:** PowerPoint/VSTO/WinForms; IntelligenceRibbon.ShowDemo.

**Data structures:** Current document/selection; source or template input where required.

**Failure behavior:** Compatibility/current presentation/selection and module validation; properties/library views have dedicated ownership/cleanup..

**Relevant tests (source only):** tests/s0/DemoContractTests.cs:1; tests/s3/RendererRibbonTests.cs:1; tests/compatibility/OfficeCompatibilityTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:44.

## CMD-DIRECT-ppiCompatibility

### Office Compatibility

**Business purpose:** Open or apply Office Compatibility for the current presentation..

**UI entry:** NexWave > Office Compatibility.

**Implementation files:** apps/powerpoint-vsto/ThisAddIn.cs; apps/powerpoint-vsto/DemoRibbon.cs.

**Responsible classes:** IntelligenceRibbon.ShowCompatibility.

**Execution workflow:** Office invokes ShowCompatibility; ThisAddIn manages compatibility/current window/mutation; delegates to the corresponding module..

**Dependencies:** PowerPoint/VSTO/WinForms; IntelligenceRibbon.ShowCompatibility.

**Data structures:** Current document/selection; source or template input where required.

**Failure behavior:** Compatibility/current presentation/selection and module validation; properties/library views have dedicated ownership/cleanup..

**Relevant tests (source only):** tests/s0/DemoContractTests.cs:1; tests/s3/RendererRibbonTests.cs:1; tests/compatibility/OfficeCompatibilityTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoRibbon.cs:44.

## CONTEXT-FONT

### Context font family, size, bold and italic

**Business purpose:** Per-property mixed state; preserve unrequested rich text and selected-character scope..

**UI entry:** Selection-aware properties > Font.

**Implementation files:** apps/powerpoint-vsto/ContextPropertiesController.cs.

**Responsible classes:** ContextFormatting; ContextChartFormatting.

**Execution workflow:** Per-property mixed state; preserve unrequested rich text and selected-character scope..

**Dependencies:** PowerPoint/VSTO/WinForms; ContextFormatting; ContextChartFormatting.

**Data structures:** Typed local DTOs/native selection and metadata.

**Failure behavior:** Bounds, ownership, validation and capability restrictions detailed in corresponding module design..

**Relevant tests (source only):** tests/s0/ContextFormattingTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/ContextPropertiesController.cs:59.

## CONTEXT-IDENTITY

### Stale-selection rejection and foreground behavior

**Business purpose:** Recheck live document/window/shape/text/table-cell identity; automatic panel hides when host loses foreground and does not force activation..

**UI entry:** Automatic property box.

**Implementation files:** apps/powerpoint-vsto/ContextPropertiesController.cs.

**Responsible classes:** ContextSelectionBridge; ContextPanelModel.

**Execution workflow:** Recheck live document/window/shape/text/table-cell identity; automatic panel hides when host loses foreground and does not force activation..

**Dependencies:** PowerPoint/VSTO/WinForms; ContextSelectionBridge; ContextPanelModel.

**Data structures:** Typed local DTOs/native selection and metadata.

**Failure behavior:** Bounds, ownership, validation and capability restrictions detailed in corresponding module design..

**Relevant tests (source only):** tests/s0/ContextPropertiesOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/ContextPropertiesController.cs:33.

## CONTEXT-MIXED

### Mixed group/table/native properties

**Business purpose:** Intersect shared capabilities and retain Empty/Uniform/Mixed values; table cells/characters are bounded..

**UI entry:** Properties.

**Implementation files:** apps/powerpoint-vsto/ContextPanelModel.cs.

**Responsible classes:** ContextPanelModel; ContextSelectionBridge.

**Execution workflow:** Intersect shared capabilities and retain Empty/Uniform/Mixed values; table cells/characters are bounded..

**Dependencies:** PowerPoint/VSTO/WinForms; ContextPanelModel; ContextSelectionBridge.

**Data structures:** Typed local DTOs/native selection and metadata.

**Failure behavior:** Bounds, ownership, validation and capability restrictions detailed in corresponding module design..

**Relevant tests (source only):** tests/s0/ContextPanelModelTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/ContextPanelModel.cs:62.

## LIB-SEARCH

### Indexed library/presentation/search/tag navigation

**Business purpose:** Select library/presentation, query mode and exact tag; update cancellable catalog-backed page..

**UI entry:** Library.

**Implementation files:** apps/powerpoint-vsto/LocalLibraryWorkspace.cs.

**Responsible classes:** PiidxCatalog; LocalLibraryForm.

**Execution workflow:** Select library/presentation, query mode and exact tag; update cancellable catalog-backed page..

**Dependencies:** PowerPoint/VSTO/WinForms; PiidxCatalog; LocalLibraryForm.

**Data structures:** Typed local DTOs/native selection and metadata.

**Failure behavior:** Bounds, ownership, validation and capability restrictions detailed in corresponding module design..

**Relevant tests (source only):** tests/library/PiidxGalleryTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/LocalLibraryWorkspace.cs:53.

## LIB-REFRESH

### Configured source refresh and cancellation

**Business purpose:** Background source index updates and watcher hints; cancellation/stale view revisions prevent applying outdated results..

**UI entry:** Library > Refresh/Rebuild/Cancel indexing.

**Implementation files:** apps/powerpoint-vsto/LocalLibraryWorkspace.cs.

**Responsible classes:** LibraryRefreshMonitor; PiidxCatalog.

**Execution workflow:** Background source index updates and watcher hints; cancellation/stale view revisions prevent applying outdated results..

**Dependencies:** PowerPoint/VSTO/WinForms; LibraryRefreshMonitor; PiidxCatalog.

**Data structures:** Typed local DTOs/native selection and metadata.

**Failure behavior:** Bounds, ownership, validation and capability restrictions detailed in corresponding module design..

**Relevant tests (source only):** tests/library/piidx/PiidxWorkspaceTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/LocalLibraryWorkspace.cs:53.

## LIB-THUMBNAILS

### Native thumbnail gallery and larger preview

**Business purpose:** One-at-a-time owned Office PNG preview with placeholders/error state, selected larger preview and bounded96thumbnail memory cache..

**UI entry:** Library gallery.

**Implementation files:** apps/powerpoint-vsto/LocalLibraryGallery.cs.

**Responsible classes:** SourceTemplateSession; LocalLibraryForm.

**Execution workflow:** One-at-a-time owned Office PNG preview with placeholders/error state, selected larger preview and bounded96thumbnail memory cache..

**Dependencies:** PowerPoint/VSTO/WinForms; SourceTemplateSession; LocalLibraryForm.

**Data structures:** Typed local DTOs/native selection and metadata.

**Failure behavior:** Bounds, ownership, validation and capability restrictions detailed in corresponding module design..

**Relevant tests (source only):** tests/library/GalleryTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/LocalLibraryGallery.cs:126.

## LIB-ORIGINAL

### Source-original slide/design insertion

**Business purpose:** Validate source fingerprint/native ID/size, insert one source slide, clone Design and match CustomLayout/background..

**UI entry:** Library > Insert selected slide.

**Implementation files:** apps/powerpoint-vsto/TemplateLibrary.cs.

**Responsible classes:** SourceTemplateSession.

**Execution workflow:** Validate source fingerprint/native ID/size, insert one source slide, clone Design and match CustomLayout/background..

**Dependencies:** PowerPoint/VSTO/WinForms; SourceTemplateSession.

**Data structures:** Typed local DTOs/native selection and metadata.

**Failure behavior:** Bounds, ownership, validation and capability restrictions detailed in corresponding module design..

**Relevant tests (source only):** tests/library/TemplateFidelityTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/TemplateLibrary.cs:489.

## LIB-DESTINATION

### Destination-design local slide insertion

**Business purpose:** Optional local source slide insertion applies active destination Design. Original-preservation mode is a separate operation..

**UI entry:** Library > Use destination design.

**Implementation files:** apps/powerpoint-vsto/LocalLibraryGallery.cs.

**Responsible classes:** LocalLibraryForm.

**Execution workflow:** Optional local source slide insertion applies active destination Design. Original-preservation mode is a separate operation..

**Dependencies:** PowerPoint/VSTO/WinForms; LocalLibraryForm.

**Data structures:** Typed local DTOs/native selection and metadata.

**Failure behavior:** Bounds, ownership, validation and capability restrictions detailed in corresponding module design..

**Relevant tests (source only):** tests/library/LibraryBrowserTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/LocalLibraryGallery.cs:136.

## LIB-DESIGNED

### Designed template catalog and editable drafts

**Business purpose:** Choose deterministic layout; draft fields by key; native insertion remains editable rather than flattening preview image..

**UI entry:** Library > Designed templates.

**Implementation files:** apps/powerpoint-vsto/TemplateCatalog.cs.

**Responsible classes:** TemplateCatalog; LibraryDrafts.

**Execution workflow:** Choose deterministic layout; draft fields by key; native insertion remains editable rather than flattening preview image..

**Dependencies:** PowerPoint/VSTO/WinForms; TemplateCatalog; LibraryDrafts.

**Data structures:** Typed local DTOs/native selection and metadata.

**Failure behavior:** Bounds, ownership, validation and capability restrictions detailed in corresponding module design..

**Relevant tests (source only):** tests/library/TemplateLibraryTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/TemplateCatalog.cs:1.

## LIB-PALETTES

### Twenty subtle palettes and default Executive Slate

**Business purpose:** Choose palette before insertion; color shortcut targets inserted tagged template instance/slide after insertion..

**UI entry:** Designed templates > Palette.

**Implementation files:** apps/powerpoint-vsto/TemplateCatalog.cs.

**Responsible classes:** TemplatePalettes; TemplateInserter.

**Execution workflow:** Choose palette before insertion; color shortcut targets inserted tagged template instance/slide after insertion..

**Dependencies:** PowerPoint/VSTO/WinForms; TemplatePalettes; TemplateInserter.

**Data structures:** Typed local DTOs/native selection and metadata.

**Failure behavior:** Bounds, ownership, validation and capability restrictions detailed in corresponding module design..

**Relevant tests (source only):** tests/library/TemplateLibraryTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/TemplateCatalog.cs:63.

## LIB-ALIGNMENT

### Designed-template text alignment

**Business purpose:** Default Center/Middle with independent horizontal/vertical control; selected text/group or full designed slide updates native paragraphs/anchors..

**UI entry:** Template Alignment / designed-template picker.

**Implementation files:** apps/powerpoint-vsto/TemplateLibrary.cs.

**Responsible classes:** TemplateAlignment; TemplateInserter.

**Execution workflow:** Default Center/Middle with independent horizontal/vertical control; selected text/group or full designed slide updates native paragraphs/anchors..

**Dependencies:** PowerPoint/VSTO/WinForms; TemplateAlignment; TemplateInserter.

**Data structures:** Typed local DTOs/native selection and metadata.

**Failure behavior:** Bounds, ownership, validation and capability restrictions detailed in corresponding module design..

**Relevant tests (source only):** tests/library/TemplateLibraryTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/TemplateLibrary.cs:133.

## LIB-BUILDS

### Native template builds and transitions

**Business purpose:** Stage shapes into click-triggered Fade/Wipe/Fly-in.35s builds; new-slide None/Fade/Wipe transition and click-only advance..

**UI entry:** Designed templates > build/transition options.

**Implementation files:** apps/powerpoint-vsto/TemplateLibrary.cs.

**Responsible classes:** TemplateInserter.

**Execution workflow:** Stage shapes into click-triggered Fade/Wipe/Fly-in.35s builds; new-slide None/Fade/Wipe transition and click-only advance..

**Dependencies:** PowerPoint/VSTO/WinForms; TemplateInserter.

**Data structures:** Typed local DTOs/native selection and metadata.

**Failure behavior:** Bounds, ownership, validation and capability restrictions detailed in corresponding module design..

**Relevant tests (source only):** tests/library/TemplateLibraryTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/TemplateLibrary.cs:106.

## LIB-OBJECTS

### Static template group in current slide

**Business purpose:** Insert tagged editable static group in existing slide; staged animation mode explicitly rejected..

**UI entry:** Designed templates > object-only.

**Implementation files:** apps/powerpoint-vsto/TemplateLibrary.cs.

**Responsible classes:** TemplateInserter.

**Execution workflow:** Insert tagged editable static group in existing slide; staged animation mode explicitly rejected..

**Dependencies:** PowerPoint/VSTO/WinForms; TemplateInserter.

**Data structures:** Typed local DTOs/native selection and metadata.

**Failure behavior:** Bounds, ownership, validation and capability restrictions detailed in corresponding module design..

**Relevant tests (source only):** tests/library/TemplateLibraryTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/TemplateLibrary.cs:66.

## LIB-LAYOUTLINK

### Live edge/offset layout links

**Business purpose:** Store native leader/follower edge/offset rule and reflow from750ms timer; pausing occurs on invalid dependencies..

**UI entry:** Layout Links floating controls.

**Implementation files:** apps/powerpoint-vsto/TemplateLibrary.cs.

**Responsible classes:** LayoutRule; LibraryLayoutLinks.

**Execution workflow:** Store native leader/follower edge/offset rule and reflow from750ms timer; pausing occurs on invalid dependencies..

**Dependencies:** PowerPoint/VSTO/WinForms; LayoutRule; LibraryLayoutLinks.

**Data structures:** Typed local DTOs/native selection and metadata.

**Failure behavior:** Bounds, ownership, validation and capability restrictions detailed in corresponding module design..

**Relevant tests (source only):** tests/library/LibraryBrowserTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/TemplateLibrary.cs:411.

## AGENDA-NATIVEEDIT

### Native caption edits and hierarchy propagation

**Business purpose:** Resolve caption edits against canonical baseline; preserve formatting and update views/page links; conflict stops mutation..

**UI entry:** Agenda native fields and Manage Agenda.

**Implementation files:** apps/powerpoint-vsto/ChapterAgenda.cs.

**Responsible classes:** ChapterAgendaPlanner; ChapterAgenda.

**Execution workflow:** Resolve caption edits against canonical baseline; preserve formatting and update views/page links; conflict stops mutation..

**Dependencies:** PowerPoint/VSTO/WinForms; ChapterAgendaPlanner; ChapterAgenda.

**Data structures:** Typed local DTOs/native selection and metadata.

**Failure behavior:** Bounds, ownership, validation and capability restrictions detailed in corresponding module design..

**Relevant tests (source only):** tests/s0/ChapterAgendaOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/ChapterAgenda.cs:70.

## AGENDA-MOVE

### Chapter subtree/content movement

**Business purpose:** Plan and move contiguous chapter subtree/content slides within scope; compensate on failure..

**UI entry:** Manage Agenda > Move chapter.

**Implementation files:** apps/powerpoint-vsto/ChapterAgenda.cs.

**Responsible classes:** ChapterAgendaPlanner.

**Execution workflow:** Plan and move contiguous chapter subtree/content slides within scope; compensate on failure..

**Dependencies:** PowerPoint/VSTO/WinForms; ChapterAgendaPlanner.

**Data structures:** Typed local DTOs/native selection and metadata.

**Failure behavior:** Bounds, ownership, validation and capability restrictions detailed in corresponding module design..

**Relevant tests (source only):** tests/s0/ChapterAgendaOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/ChapterAgenda.cs:120.

## S0-RECOVERY

### S0 metadata edit capture and duplicate identity repair

**Business purpose:** Capture supported native series values/labels; revision increments; duplicates receive fresh GUID; damaged objects preserved..

**UI entry:** Check & Sync and active-deck BeforeSave.

**Implementation files:** apps/powerpoint-vsto/PowerPointHost.cs.

**Responsible classes:** PowerPointHost; MetadataCodec.

**Execution workflow:** Capture supported native series values/labels; revision increments; duplicates receive fresh GUID; damaged objects preserved..

**Dependencies:** PowerPoint/VSTO/WinForms; PowerPointHost; MetadataCodec.

**Data structures:** Typed local DTOs/native selection and metadata.

**Failure behavior:** Bounds, ownership, validation and capability restrictions detailed in corresponding module design..

**Relevant tests (source only):** tests/s0/HostIntegrationTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/PowerPointHost.cs:360.

## AUTOMATION-COM

### Bounded COM automation observation/insertion

**Business purpose:** Expose build/ribbon state and bounded validated S0 JSON insertion with UI-STA enforcement; not HTTP or arbitrary function execution..

**UI entry:** COMAddIn.Object automation.

**Implementation files:** apps/powerpoint-vsto/DemoApi.cs.

**Responsible classes:** IDemoAutomation; DemoAutomation.

**Execution workflow:** Expose build/ribbon state and bounded validated S0 JSON insertion with UI-STA enforcement; not HTTP or arbitrary function execution..

**Dependencies:** PowerPoint/VSTO/WinForms; IDemoAutomation; DemoAutomation.

**Data structures:** Typed local DTOs/native selection and metadata.

**Failure behavior:** Bounds, ownership, validation and capability restrictions detailed in corresponding module design..

**Relevant tests (source only):** tests/s0/ContextPropertiesOfficeTests.cs:1.

**Status:** Implemented source; runtime unverified.

**Source evidence:** apps/powerpoint-vsto/DemoApi.cs:12.
