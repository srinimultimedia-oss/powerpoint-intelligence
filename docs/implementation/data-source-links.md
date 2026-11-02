# Data Source ribbon menu

## v17 explicit mapping update

AD-017 approves explicit worksheet/range/columns and confirmed relinking. The **Link Excel Range** workflow selects an existing workbook, worksheet, bounded A1 rectangle (2�50 rows including headers, 2�15 columns), category column and data columns. It stores a version-2 source record with the expected headers; values are read from the explicitly selected range. The existing four file-reference choices remain available. Version-1 references retain their historical behavior.

**Data Link Manager** inventories source metadata and object associations, reports available/missing/remote/damaged states, locates objects and permits explicit compatible native-chart refresh. A file timestamp is not a successful-refresh baseline. Network paths are not automatically probed. XY, bubble, single-series projections, mixed/secondary-axis charts and older charts without orientation metadata are refused by mapped refresh; tables/diagrams have no mapped refresh adapter yet.

**Relink selected chart** selects a new mapping and asks for confirmation before replacing that object's reference. It leaves chart values unchanged until explicit refresh. Missing sources and header/dimension mismatches stop refresh; no implicit path search/rebinding occurs. Old registry entries remain visible, and uncertain metadata removal is reported with the replacement retained for inspection. No batch or automatic refresh is implemented.

The namespace remains `urn:powerpoint-intelligence:data-source:v1`; the XML `version` attribute distinguishes file-only v1 and mapped Excel v2 records. The codec rejects missing/extra/invalid range metadata. Source headers/paths are retained in presentation metadata as necessary for explicit mapping; source contents are not logged.

The sections below describe historical increments and their dated evidence; they do not override the v17 limits above. Native v17 Excel access, refresh, relinking and persistence await desktop validation.


The Intelligence ribbon now includes a large **Data Source** menu in the Data group:

| Option | File picker extensions |
|---|---|
| Excel sheet | .xlsx, .xls, .xlsm, .xlsb |
| CSV file | .csv |
| JSON file | .json |
| XML file | .xml |

Each choice opens the native Windows file browser filtered to its format, requires an existing file and allows one selection. Cancel creates no link. An open presentation and an allowed Office compatibility profile are required. The active presentation is captured before the dialog so switching windows cannot redirect the selected source to another presentation.

The selected source is linked to the presentation as a file reference: stable GUID, format, normalized full path and UTC linking time. Each record is a versioned CustomXMLPart in `urn:powerpoint-intelligence:data-source:v1`. Re-selecting the same format/path returns the existing link. Save the presentation to retain its references. The task pane confirms the linked format and filename. Missing files retain their saved reference when metadata is read; selecting a new link still requires an existing matching file.

Linking alone retains a file reference. Chart commands now read that linked file and embed the resulting values in native chart workbooks or versioned editable-diagram metadata. Full source paths are retained in the presentation; source contents are never written to diagnostic logs. Excel source reading uses a separate hidden, read-only Excel instance, disables macros and external-link updates, reads the first worksheet and closes only its owned instance. Automatic refresh and arbitrary sheet/range/field mapping remain unimplemented. See the linked-report behavior below.

The v1 codec rejects unknown versions/formats/fields, ambiguous path elements, invalid GUIDs, relative/device paths, DTD/entities and oversized metadata. Existing unrelated XML parts and slide objects are preserved. A provisional guard allows 256 source references per presentation. No production Excel refresh or moved-file policy is chosen by this UI increment.

## Implementation

- apps/powerpoint-vsto/DataSources.cs: immutable option catalog, picker, reference validation, v1 XML codec and presentation store.
- DemoRibbon.cs / ThisAddIn.cs: four bound commands, context-enabled menu and linked-file confirmation; ribbon marker v6.
- RibbonIcons.cs: original Excel/CSV/JSON/XML icons.
- tests/s0/DataSourceTests.cs: 41 local cases plus an implemented Office save/reopen scenario.

## Validation — 22 October 2026

Run: `artifacts/validation-20261022-213314-e352710f`.
Source SHA256: `0D129EA6BA875FB41E61A7E21809BE7AD38538A92A05DAD9722D98AB137BFE1F`.

**20/20 local checks pass**, including the 41 source-selection/metadata checks, ribbon XML/callback checks, icons for all 111 controls and AnyCPU/x86/x64 VSTO compile targets. Tests exercise all seven supported extensions, cancelled selection, rejection before persistence, Unicode/ampersand paths, exact XML round trips, missing-source reference retention and malformed metadata.

The interactive suite attempted installed-ribbon, demo-feature-runtime, native-shapes-office and office-1; each failed at PowerPoint COM activation with 0x80070005 in this restricted agent context. Therefore neither actual picker interaction from the newly installed ribbon nor native CustomXMLPart save/reopen is marked passed. The implemented Office test checks four references, duplicate suppression, untouched slides and preserved IDs after reopening, using fixture selections; it does not substitute for a visual picker test.

Rebuild/install current source with `scripts/Install-Local.ps1` in the approved normal desktop environment, then run `scripts/Test-All.ps1 -InteractiveOffice`. In PowerPoint, open a disposable presentation, select each Data Source option, verify its file filter, cancel once, select a matching file and confirm its filename in the pane. Save/reopen the test presentation and verify its four retained references. Existing installed versions will not gain this menu until the rebuilt package is installed. Installed-ribbon validation now requires marker v6 and all four bound source options.

## Linked chart/report generation — 22 October 2026

Choosing a chart reads a source already linked to the active presentation. With multiple sources, a modal selector requires choosing one. With no source, a missing source or invalid values, chart creation stops and shows a message; there is no sample fallback. The ribbon and task-pane column actions both follow this path. The editable-shape column action builds a validated canonical v2 model and scene from linked categories and values, retaining them through the renderer's existing metadata. Embedded sample resources remain isolated test fixtures.

**Generate linked report** appends seven chart slides (column, bar, stacked column, percent column, line, area and pie) to the current presentation, reading the selected source once before adding slides. It never calls `Presentations.Add`. Existing slides remain in the same document. Individual chart actions insert on the current slide, rather than replacing a selected existing chart. Saving the current presentation is still an explicit user action.

### Supported local table formats

All formats require 2–15 columns and 1–49 data rows, with headers first, category/X/task labels in column one and finite numeric values or blanks in remaining columns. Files are limited to 20 MB. This is an explicit bounded table contract, not automatic interpretation of arbitrary files.

- Excel: the first worksheet's used range; evaluated numeric formula values can be read, but macros and external-link updates are disabled.
- CSV: comma separated, UTF-8, header row, quoted fields, escaped quotes and CRLF/LF supported.
- JSON: `{"headers":["Category","Revenue"],"rows":[["East",137.5],["West",-8]]}`. Numeric text and null/empty numeric cells also work.
- XML: `<table><header><cell>Category</cell><cell>Revenue</cell></header><row><cell>East</cell><cell>137.5</cell></row></table>`. DTD and external entities are prohibited.

Column, bar, stacked/percent, line, profile and area use the table's numeric series. Pie/doughnut use the first numeric series. Combo requires two series. Scatter requires numeric X in the first column. Bubble requires exactly X, Y, Size, with nonnegative nonblank sizes. Profile preserves the existing row-oriented plotting behavior.

Custom editable diagrams consume linked values too: waterfall requires Category + signed Change (every row is a cumulative movement starting from zero; there are no implicit total rows); Gantt requires Task + numeric Start + positive Duration (no calendar/date scheduling); Mekko requires nonnegative values, uses category totals for widths, and uses series shares for segment heights, with either percent or absolute total height. These simple mappings are disclosed in command descriptions and validation messages. Diagram snapshots use `urn:powerpoint-intelligence:linked-table:v1` and retain the selected source identity. Labels do not yet have automatic collision avoidance. Charts do not automatically refresh after external file changes; invoking chart generation reads the current file again. No moved-file rebinding policy is implied.

### Current evidence

Focused tests: 57 source/link/parser checks and x64 native renderer core 41/41 plus 9 ribbon preparation checks pass. New parser tests use deliberately different values (137.5 and -8) to catch sample substitution. Office scenarios now link explicit fixtures and assert linked values and seven report slides in the same presentation. Their execution remains subject to desktop COM activation; compiled scenarios are not runtime evidence.

Full run: `artifacts/validation-20261022-215845-fbd835a9`, source SHA256 `A5FDB384BE2F20BD4F8DADF4B05B252EEE0031140C5359017653D714BD938963`. All 20 local checks pass, including 57 source checks, 41 native renderer core tests, 9 renderer preparation tests, explicit x64 execution and AnyCPU/x86/x64 builds. Installed-ribbon, demo-feature-runtime, native-shapes-office and office-1 fail at PowerPoint COM activation with `0x80070005`; no interactive linked-file or save/reopen acceptance is claimed. This run did not install the new package. Reinstall from the normal desktop shell with `scripts/Install-Local.ps1`, then run `scripts/Test-All.ps1 -InteractiveOffice` and visually verify the v6 ribbon.

## Excel column selection — 22 October 2026

Every linked Excel chart/report action now asks which columns to use, before generating native objects. The dialog shows numbered worksheet headers, a category/X/task dropdown and an unchecked list of data-series columns. Choose one category column and at least one different numeric column. Numbering distinguishes duplicate header names. Checked data columns are used in worksheet order; bubble needs Y then Size, and Gantt needs Start then Duration. Changing the category removes it from checked data columns. Cancel or closing the dialog aborts generation without adding charts/slides. For a report, the selection is made once and reused across its seven chart slides. Editable shape columns use the same selector.

Numeric validation happens after projecting the chosen columns: unrelated text columns in the workbook no longer block chart creation. The original workbook and source table are unchanged. Only selected labels and values are embedded in the generated output. The selected data must still satisfy the chart family's existing input rules. Selection is requested again for subsequent chart actions; saved field mapping and automatic refresh remain outside this increment.

Reproduction after installation: link a first-sheet Excel table with Notes (text), Sales (numeric), Region (text), Costs (numeric). Choose a chart; select Region as category and Sales/Costs as series. Confirm that Notes is excluded. Repeat selecting only Costs; verify one series. Cancel once and confirm no new chart. Select Notes as a numeric series and verify an error keeps the dialog open. Generate linked report and confirm one column-selection prompt and seven slides in the same deck. Worksheet used-range limits remain 2–50 rows including headers and 2–15 columns.

Local tests now include projection of a non-first category, selective and multiple series, exclusion of unrelated text, source immutability, invalid/duplicate/out-of-range selections, nonnumeric chosen values and cancellation: 68 source checks total. The modal picker and Excel/PowerPoint runtime flow have not been interactively verified in this restricted agent session.

Validation run: `artifacts/validation-20261022-220826-50b19448`; source SHA256 `3E0CF5D1773844C2606680BA22F65E8F5EC1640D30359191916DEC4B1278A3D4`. All 20 local checks pass, including 68 source checks and AnyCPU/x86/x64 builds. This run omitted interactive Office checks; dialog display and workbook-to-PowerPoint behavior remain unverified. Install current source before expecting the new selector in the existing desktop add-in.

## Installer marker false alarm — 22 October 2026, deployment 1.0.0.8

The desktop install log `artifacts/install-20261022-221234.log` stopped at package validation after a successful build. Read-only inspection of the actual x64 DLL found `PPI-DEMO-RIBBON-20261022-v6` only when UTF-16 decoding began at offset one. `Test-DemoPackage.ps1` decoded the complete PE at offset zero and therefore falsely classified the current payload as stale. This was a validator alignment bug; no rebuild or security-policy change was needed.

The validator now searches the exact UTF-16 marker byte sequence at every byte offset through `scripts/Package-Content.ps1`. Existing source-timestamp, SHA256 manifest dependency and signature-presence checks are unchanged. Eight regression cases exercise even/odd offsets, end-of-file, old/truncated/absent markers and empty input; these run as part of ribbon-definition validation. In PowerShell 7.6.5 the marker cases, complete ribbon regression and validation of the actual signed x64 deployment 1.0.0.8 all pass. Evidence: `artifacts/package-marker-regression.log`, `artifacts/package-marker-ribbon-regression.log`, `artifacts/package-validation-1.0.0.8.log`. A separate Windows PowerShell 5.1 process was blocked by AuthorizationManager in the restricted agent context before script execution; no bypass was attempted. Actual installer/publisher prompt and runtime loading are still desktop validation steps. The verified existing package can be installed using `scripts/Install-Local.ps1 -SkipBuild`.
