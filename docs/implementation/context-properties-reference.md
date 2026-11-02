# Contextual properties reference and acceptance

Implementation status and executed-check evidence: [floating contextual properties](context-properties.md). Focused automated desktop acceptance passed on 29 October 2026; manual visuals and remaining matrix coverage stay open.

Study date: 2026-10-28. Scope: independent public UI behavior and local source inventory, not vendor implementation. AGENTS.md prohibits vendor binary/protected-resource inspection. REA native tooling in this session is bound to Mach-O, not Windows UI; parent uses permitted computer-use observation. No desktop actions or product-code edits were performed by this documentation task.

## Evidence ledger

| Evidence | Established fact | Boundary |
|---|---|---|
| User-supplied screenshot, temporary image identifier beginning 58a960, as described by parent | Floating properties shows font/size, mixed font color, horizontal left, vertical top, three orientation choices, mixed fill and type-conversion controls. | Screenshot is a static state, not proof of automatic appearance, event timing, action success, every element family or mixed-selection semantics. User attachment: codex-clipboard-58a960e8-f74f-49bd-bc25-d235d0e6372c.png. |
| [Official element formatting manual](https://www.think-cell.com/en/resources/manual/format-and-style-elements) | Selecting one or multiple elements opens a mini toolbar with selection-specific options. Text controls include font/style/size/color, left/center/right, top/middle/bottom and horizontal/90-degree directions; built-in PowerPoint text formatting also applies. | Public behavior synopsis; no copied design/assets or claim to replicate internal selection machinery. Per-family settings require their own evidence. |
| [Official linked-element formatting manual](https://www.think-cell.com/en/resources/manual/format-and-style-charts) | Linked elements have specific formatting controls, including matching supported formatting to their data source and chart-specific schemes/features. | This does not authorize a generic fill/font operation against arbitrary chart subobjects or establish live-data support in our plugin. |
| Local ContextFormatting.cs | Source contains bounded native font collection, mixed-state merging and individual-character font snapshots; table selected cells and grouped text are handled by explicit rules. | Source-level capability, not proof the new automatic floating UI correctly invokes it. |
| Local NativeFormatting.cs, NativeChartControls.cs, NativeSeriesControls.cs and element modules | Existing commands provide candidates for dispatch. | A button label or available API is insufficient; enablement, target resolution and native results need executed acceptance. |

No screenshot proves automatic show/hide, viewport anchoring, DPI behavior, keyboard focus, multiple-window ownership, modal interaction, stale-selection safety or save/reopen behavior. Those remain native acceptance gates.

## Capability-based dispatch matrix

The approximately 90 tool commands are not 90 object types. Insertion, selection commands and non-element utilities must be distinguished. The following covers local families found in DemoTools catalog and current modules; it defines intended dispatch boundaries, not completed UI acceptance.

| Local family / source | Common properties when actually supported | Family-specific dispatch | Do not infer |
|---|---|---|---|
| Native text boxes/placeholders and freeform text | Font, size, bold/italic, text color, paragraph alignment, anchor, orientation; shape fill/outline where supported | TextShapeElements settings only for managed members | Every placeholder is safe to convert; selected substring means whole shape. |
| Managed pentagon, chevron, rounded rectangle; TextShapeElements | Text/shape common subset | Type conversion, edge-lock/fit policy, geometry controls exposed by that module | Native AutoSize equals a full constraint solver. |
| Process grids; ProcessFlowTools | Selected cell text/shape subset | Row/column operations, duplicate, striping and managed reflow | Outer group should receive every text property or be converted wholesale. |
| Harvey balls and checkboxes; IndicatorElements | Fill/outline only when mapped to canonical state | Segment/completion, symbol state, size and type conversion through indicator API | Generic group fill or text-font changes safely preserve rendered state. |
| Universal connectors; UniversalConnectors | Line color, weight and arrows where supported | Endpoints, routing, attachment and refresh | Shape fill/font applies; every chart interior feature is an available endpoint. |
| Native column/bar, stacked/100%, line/profile, area, pie/doughnut, scatter/bubble/combo | Explicit chart-supported subset only | Edit Data, legend/labels/table, native chart/axis/series dialogs | Chart is a TextFrame; every chart kind has axes, labels or the same series settings. |
| Native-renderer grouped shapes; RendererRibbonCommand/host | Native leaf formatting only under a declared policy | Managed renderer refresh/metadata editor if actually implemented | Generic group mutation updates canonical chart data or survives regeneration. |
| Grouped waterfall/Mekko/Gantt diagrams | Safe native selected-leaf subset | Source-specific regeneration only when wired | Arbitrary shape edits update chart semantics; native chart APIs apply. |
| Chart annotations; ChartAnnotations | Specific line/text components | Refresh selected annotation from its managed source | Change label changes underlying numeric source or axis binding. |
| Native tables; table-data / ContextFormatting | Selected cells, or entire table only under explicit selection policy | Existing bounded table editor; per-cell font/alignment/fill | A table group or chart data table is interchangeable with native table cells. |
| Agenda chapters/contents; ChapterAgenda + legacy PresentationAgenda | Managed field-aware styling | Chapter hierarchy/rename, scope/style class, synchronization | Generic text replacement automatically renames a chapter in canonical records. |
| Library originals, generated library templates; TemplateLibrary | Native selected-leaf properties | Palette/alignment/layout links only for recognized ownership | All imported vendor or user shapes carry our metadata or can be converted. |
| Pictures, icons, raster copies, images from file | Geometry and supported image appearance only | Appropriate image dialog if available | Editable text exists inside a bitmap; fill color recolors pixels. |
| Ordinary shapes/groups, symbols, path text | Intersection of native leaf capabilities | Safe group/ungroup/arrange only where valid | Every group is a managed indicator or diagram; every glyph has indicator state. |
| Linked chart/source selections; DataLinkManager/MappedExcelData | Properties of actual selected object | Source mapping/relink/refresh only for recognized link records | Existence of a linked source makes any element a supported data consumer. |
| Utilities: quality inspector, PDF/export/import, service configuration, source chooser, workbook manager | No standalone slide-object properties | Keep their dedicated dialogs/commands | Every command should appear as a propertybox type; vendor admin/licensing belongs here. |

## Automatic contextual UI acceptance checklist

1. **Insertion and selection:** every supported insertion path selects the created object and displays the applicable panel without an extra ribbon click. Existing object selection, keyboard selection, group-child selection, chart/table selection and slide changes recalculate capabilities. No selection hides the panel. Unsupported selection hides it or shows a precise read-only explanation.
2. **Editing scope:** text-caret/substring editing preserves caret and applies only to the intended range. Whole-object actions must not silently replace text selection. The panel must not steal typing, interfere with PowerPoint context menus or cancel native double-click behavior for unrelated objects/add-ins.
3. **Mixed selections:** show an explicit mixed/unknown state, never a fabricated default. Leave unchosen mixed fields unchanged. Apply only a validated capability intersection or clearly refuse unsupported mixtures before mutation. Test table cells, groups, indicators mixed with ordinary shapes and different chart families.
4. **Freshness and ownership:** resolve document/window/slide/object identity again before every action. Reject closed, deleted, copied or changed selections and selection switches while color/font dialogs are open. Avoid cached COM references that silently target a previous deck. Document whether intentionally pinned inspectors follow or retain selection.
5. **Window lifecycle:** separate PowerPoint windows/documents must own their panel state. Hide on deactivation/minimize/slide-show/modal dialogs; restore on valid return and survive ribbon/dialog activation. Do not float above unrelated applications. Dispose timers/events at shutdown and avoid retaining closed presentation references.
6. **Position:** use active window/screen coordinates and a clamped work area, not raw slide points. Verify 100/125/150/200% DPI, multi-monitor negative coordinates, zoom, scrolling, ribbon/sidebar changes, restored/maximized windows and narrow viewports. Keep the selected content and native controls accessible.
7. **Mutation:** validate all targets before changes; one undo group per intentional action; test rollback and partial-failure reporting. Preserve unrelated styles, metadata, layout bindings and canonical state. Failed reads/actions must not loop dialogs or spam logs. Do not log document text.
8. **Functional verification:** execute each visible control against each enabled family and reopen the file. Check canonical-aware type changes, chart controls, indicator states, connector attachments, agenda styling and palette persistence independently. A panel screenshot, source compilation or pure capability test alone cannot establish this.

## Implementation recommendations and current risks

Use a typed capability snapshot keyed to the current window, slide and selection generation. Build common controls from the supported intersection, then add specific actions through existing validated modules. Dispatch actions through the parent's normal busy/error/Undo path; suppress reentrant event updates while mutations or modal editors run. Coalesce selection events with bounded UI-thread work rather than scanning the entire presentation on a timer.

Highest-risk gaps are applying native text/font collectors to chart/indicator internals, flattening mixed text runs, losing canonical ownership on type conversion, changing stale selections after a dialog, and panel focus/DPI behavior. Keep unsupported controls disabled or omitted with a reason. The user requested all local families, so acceptance requires the matrix above to be resolved row by row; no complete parity claim is made here.
