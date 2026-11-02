# Original ribbon and workspace experience

Public reference reviewed on 22 October 2026:
https://www.think-cell.com/en/resources/manual/think-cell-elements
https://www.think-cell.com/en/resources/manual/column-line-area-charts
https://www.think-cell.com/en/resources/kb/0119

The owner's exact ribbon URL was inaccessible via the browsing tool. The public
manuals describe an Elements insertion menu, chart families, selection-based
editing, and commands enabled in editable slide contexts. These inform interaction
requirements only. No proprietary code, artwork, icons or binary behavior was used.

## Implemented first increment

Intelligence tab separates Create, Review and Workspace. Elements offers the
implemented native clustered-column sample, explicitly labeled as sample data.
Check & Sync captures native edits and reports damaged metadata. Open Workspace
opens a pane with a short introduction, insert/review/compatibility actions and
selection/status details. Controls have descriptive tooltips and accessible names;
the pane uses system-scaled text. No unimplemented chart is offered as a command.

The task pane is a proposed usability improvement: discoverable actions and
status in one place, without feasibility-stage labels. Improved usability is
not established without desktop visual review and user evaluation.

## Broader implementation sequence

| Area | Original interaction direction | Status / gate |
|---|---|---|
| Element discovery | Searchable gallery grouped by comparison, trend, composition and planning | Future; after S0 runtime acceptance |
| Chart families | Native clustered/stacked bars, line/area, combination, waterfall, scatter, pie, then more complex families | Only clustered-column S0 sample currently implemented |
| Data editing | Selected-chart data grid with explicit apply and validation | Future typed command implementation |
| Labels / formatting | Contextual inspector for labels, axes, legend and style | Future; preserve native edits |
| Arrangement | Selection-aware alignment, resize and distribution | Future; reversible operations and undo tests |
| Data links | Clear source, freshness and refresh controls | Blocked on approved link semantics |
| Recovery | Plain-language issue details, preserve editable objects | Current S0 diagnostics; richer guidance future |

This document does not authorize moving past the owner's S0 stability gate.
Full feature parity with the referenced product is not implemented or implied.
Ribbon structural checks and compilation do not prove runtime loading, layout,
accessibility, Office theme compatibility or high-DPI appearance. Verify in the
installed add-in, including keyboard navigation, small pane widths and no-document
states. Rebuild/install to receive source changes; -SkipBuild uses old artifacts.

## Data Source menu increment

The current ribbon also exposes Data Source > Excel sheet / CSV file / JSON file / XML file, each with a format-specific native picker and original icon. Selected paths are saved as versioned presentation references; see data-source-links.md for scope, tests and runtime gates. The existing native-chart and renderer acceptance gates are unchanged. Source marker v5 requires a rebuilt installed package.
