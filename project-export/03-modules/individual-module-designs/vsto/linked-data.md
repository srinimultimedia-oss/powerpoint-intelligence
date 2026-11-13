# Linked files, Excel mapping and source manager

<!-- generated-toc -->
## Contents

- [Purpose and implementation status](#purpose-and-implementation-status)
- [Source locations and internal components](#source-locations-and-internal-components)
- [Inputs, outputs, data structures and dependencies](#inputs-outputs-data-structures-and-dependencies)
- [Internal execution and external integration](#internal-execution-and-external-integration)
- [Configuration, validation and failure behavior](#configuration-validation-and-failure-behavior)
- [Security considerations and cleanup](#security-considerations-and-cleanup)
- [Test references](#test-references)
- [Known limitations and technical debt](#known-limitations-and-technical-debt)
- [Extension points](#extension-points)
- [Classes, interfaces and method navigation](#classes-interfaces-and-method-navigation)
<!-- end-generated-toc -->

**VERIFIED — source implementation. Current runtime acceptance is not established by this design review.**

## Purpose and implementation status

Link local files, select categories/series, render native editable objects and explicitly refresh eligible mapped charts.

Implemented local source is distinguished from supported subset/absent external behavior below. See current test evidence in the export testing chapter; do not treat historical logs or method existence as a current runtime pass.

## Source locations and internal components

- `apps/powerpoint-vsto/DataSources.cs:1` — LinkedTable, LinkedTableProjection, LinkedChartData, DataSourceOption, DataSourceLink, DataSources.
- `apps/powerpoint-vsto/ExcelRangeMapping.cs:1` — ExcelRangeMapping, MappedExcelData.
- `apps/powerpoint-vsto/MappedLinkOperations.cs:1` — MappedLinkOperations.
- `apps/powerpoint-vsto/DataLinkManager.cs:1` — LocalLinkFileState, LocalLinkFileStatus, LinkObjectReference, LocalLinkSource, LocalLinkInventory, DataLinkInventory, DataLinkManager.
- `apps/powerpoint-vsto/DemoLinks.cs:1` — DemoWorkbookLink, DemoLinks.

## Inputs, outputs, data structures and dependencies

Inputs: Excel/CSV/JSON/XML path, LinkedTable, mapping and native chart. Outputs: registered and shape-attached DataSourceLink, embedded workbook or native table/diagram. Dependencies: PathPolicy, Excel/PowerPoint COM, XmlPartUpdates and DataLinkInventory. Legacy DemoLinks uses a separate first-sheet .xlsx WorkbookLink namespace.

## Internal execution and external integration

Data Source saves format/path/UTC/GUID after an extension-specific picker without parsing file data. At chart creation Load reads CSV, JSON headers/rows, XML table/header/row/cell or read-only Excel. Unmapped Excel reads first-sheet UsedRange then asks category/series; mapped Excel persists worksheet/A1 rectangle/header signature/zero-based columns. Write activates embedded workbook, performs a single Value2 matrix write, checks cells, binds source and waits for real native series values before and after close. Manager inventories registry/object references including groups, damaged/ambiguous states and local-file metadata; it locates and explicitly refreshes eligible objects. Relink confirms the replacement and retains old registry records.

## Configuration, validation and failure behavior

Source <=20 MiB,1–49 data rows,2–15 columns, <=256 source records; mapped rectangle2–50 rows including headers; table<=400 cells. Current Excel reader is hidden/read-only, AutomationSecurity=3 and UpdateLinks=0. Labels are literal text. XML DTD/resolver disabled. Manager does not probe UNC/network drives. Chart acknowledgement bounded to8sec.

## Security considerations and cleanup

Local native calls run in the current Office user process and require supported selection/ownership; a mutex/selection key is not authentication or RBAC. File/metadata content must remain bounded and error text must not expose data or credentials. Explicit owned-resource cleanup and error/rollback paths are part of the cited implementation, not a guarantee that a failed COM server can always be restored.

## Test references

Lexical module-class matches in test source (not instrumentation coverage; no test is claimed executed here):

- `tests/s0/DataLinkManagerTests.cs:11`.
- `tests/s0/DataSourceTests.cs:26`.
- `tests/s0/DemoContractTests.cs:59`.
- `tests/s0/DemoScenario.cs:118`.
- `tests/s0/ExcelRangeMappingTests.cs:11`.
- `tests/s0/LinkedChartDiagnostic.cs:41`.
- `tests/s0/MappedExcelOfficeTests.cs:23`.
- `tests/s0/XmlPartUpdateTests.cs:25`.

## Known limitations and technical debt

Mapped refresh rejects unmapped/old orientation metadata, tables/diagrams, XY/bubble/pie/doughnut and mixed/secondary-axis charts. Changed headers/dimensions stop refresh, never guess relinking. Registry separates attached XML parts after reopen. Child Excel RCWs are not uniformly explicitly released; cleanup is not a proven process-leak guarantee.

## Extension points

Introduce additional projections only with exact typed mappings, explicit confirmation, formatting retention and native save/reopen/rollback acceptance. Never equate file metadata availability with successful refresh.

## Classes, interfaces and method navigation

Method names below are source declarations; private/internal members are not public API promises. The source hash/review manifest gives exact reviewed bytes.

| Declaration | Kind | Source |
|---|---|---|
| `LinkedTable` | class | `apps/powerpoint-vsto/DataSources.cs:15` |
| `Number` | method | `apps/powerpoint-vsto/DataSources.cs:19` |
| `Validate` | method | `apps/powerpoint-vsto/DataSources.cs:25` |
| `LinkedTableProjection` | class | `apps/powerpoint-vsto/DataSources.cs:32` |
| `Cells` | method | `apps/powerpoint-vsto/DataSources.cs:34` |
| `LinkedChartData` | class | `apps/powerpoint-vsto/DataSources.cs:39` |
| `SelectColumns` | method | `apps/powerpoint-vsto/DataSources.cs:41` |
| `ChooseColumns` | method | `apps/powerpoint-vsto/DataSources.cs:48` |
| `PickMappedColumns` | method | `apps/powerpoint-vsto/DataSources.cs:55` |
| `PickColumns` | method | `apps/powerpoint-vsto/DataSources.cs:56` |
| `Choose` | method | `apps/powerpoint-vsto/DataSources.cs:84` |
| `Load` | method | `apps/powerpoint-vsto/DataSources.cs:94` |
| `FromRows` | method | `apps/powerpoint-vsto/DataSources.cs:114` |
| `Csv` | method | `apps/powerpoint-vsto/DataSources.cs:116` |
| `Excel` | method | `apps/powerpoint-vsto/DataSources.cs:127` |
| `WorkbookValues` | method | `apps/powerpoint-vsto/DataSources.cs:133` |
| `ValidateWorkbookValues` | method | `apps/powerpoint-vsto/DataSources.cs:141` |
| `EmbeddedSourceAddress` | method | `apps/powerpoint-vsto/DataSources.cs:152` |
| `AwaitEmbeddedWorkbook` | method | `apps/powerpoint-vsto/DataSources.cs:157` |
| `ObserveWrite` | method | `apps/powerpoint-vsto/DataSources.cs:165` |
| `Write` | method | `apps/powerpoint-vsto/DataSources.cs:166` |
| `CloseNativeWorkbook` | method | `apps/powerpoint-vsto/DataSources.cs:200` |
| `CloseWorkbook` | method | `apps/powerpoint-vsto/DataSources.cs:206` |
| `WaitForChartValues` | method | `apps/powerpoint-vsto/DataSources.cs:215` |
| `Persist` | method | `apps/powerpoint-vsto/DataSources.cs:233` |
| `DataSourceOption` | class | `apps/powerpoint-vsto/DataSources.cs:240` |
| `DataSourceLink` | class | `apps/powerpoint-vsto/DataSources.cs:254` |
| `DataSources` | class | `apps/powerpoint-vsto/DataSources.cs:262` |
| `Option` | method | `apps/powerpoint-vsto/DataSources.cs:272` |
| `Require` | method | `apps/powerpoint-vsto/DataSources.cs:274` |
| `Validate` | method | `apps/powerpoint-vsto/DataSources.cs:275` |
| `Create` | method | `apps/powerpoint-vsto/DataSources.cs:286` |
| `ChooseAndLink` | method | `apps/powerpoint-vsto/DataSources.cs:297` |
| `PickFile` | method | `apps/powerpoint-vsto/DataSources.cs:303` |
| `Encode` | method | `apps/powerpoint-vsto/DataSources.cs:310` |
| `Decode` | method | `apps/powerpoint-vsto/DataSources.cs:325` |
| `Read` | method | `apps/powerpoint-vsto/DataSources.cs:349` |
| `ReadRegistry` | method | `apps/powerpoint-vsto/DataSources.cs:365` |
| `Save` | method | `apps/powerpoint-vsto/DataSources.cs:371` |
| `ExcelRangeMapping` | class | `apps/powerpoint-vsto/ExcelRangeMapping.cs:13` |
| `Validate` | method | `apps/powerpoint-vsto/ExcelRangeMapping.cs:19` |
| `Dimensions` | method | `apps/powerpoint-vsto/ExcelRangeMapping.cs:26` |
| `Project` | method | `apps/powerpoint-vsto/ExcelRangeMapping.cs:33` |
| `Write` | method | `apps/powerpoint-vsto/ExcelRangeMapping.cs:40` |
| `Read` | method | `apps/powerpoint-vsto/ExcelRangeMapping.cs:42` |
| `MappedExcelData` | class | `apps/powerpoint-vsto/ExcelRangeMapping.cs:50` |
| `Read` | method | `apps/powerpoint-vsto/ExcelRangeMapping.cs:52` |
| `ReadRange` | method | `apps/powerpoint-vsto/ExcelRangeMapping.cs:61` |
| `Choose` | method | `apps/powerpoint-vsto/ExcelRangeMapping.cs:65` |
| `MappedLinkOperations` | class | `apps/powerpoint-vsto/MappedLinkOperations.cs:10` |
| `Parts` | method | `apps/powerpoint-vsto/MappedLinkOperations.cs:12` |
| `ValidateRefresh` | method | `apps/powerpoint-vsto/MappedLinkOperations.cs:14` |
| `Refresh` | method | `apps/powerpoint-vsto/MappedLinkOperations.cs:22` |
| `RelinkSelected` | method | `apps/powerpoint-vsto/MappedLinkOperations.cs:34` |
| `LocalLinkFileState` | enum | `apps/powerpoint-vsto/DataLinkManager.cs:12` |
| `LocalLinkFileStatus` | class | `apps/powerpoint-vsto/DataLinkManager.cs:13` |
| `LinkObjectReference` | class | `apps/powerpoint-vsto/DataLinkManager.cs:19` |
| `LocalLinkSource` | class | `apps/powerpoint-vsto/DataLinkManager.cs:28` |
| `CanRefresh` | method | `apps/powerpoint-vsto/DataLinkManager.cs:35` |
| `LocalLinkInventory` | class | `apps/powerpoint-vsto/DataLinkManager.cs:38` |
| `DataLinkInventory` | class | `apps/powerpoint-vsto/DataLinkManager.cs:45` |
| `Same` | method | `apps/powerpoint-vsto/DataLinkManager.cs:48` |
| `DecodeReferences` | method | `apps/powerpoint-vsto/DataLinkManager.cs:50` |
| `Copy` | method | `apps/powerpoint-vsto/DataLinkManager.cs:56` |
| `Key` | method | `apps/powerpoint-vsto/DataLinkManager.cs:57` |
| `Build` | method | `apps/powerpoint-vsto/DataLinkManager.cs:58` |
| `ProbeLocal` | method | `apps/powerpoint-vsto/DataLinkManager.cs:75` |
| `Capture` | method | `apps/powerpoint-vsto/DataLinkManager.cs:84` |
| `Find` | method | `apps/powerpoint-vsto/DataLinkManager.cs:96` |
| `DataLinkManager` | class | `apps/powerpoint-vsto/DataLinkManager.cs:101` |
| `Grid` | method | `apps/powerpoint-vsto/DataLinkManager.cs:111` |
| `Rescan` | method | `apps/powerpoint-vsto/DataLinkManager.cs:121` |
| `ShowObjects` | method | `apps/powerpoint-vsto/DataLinkManager.cs:127` |
| `UpdateEnabled` | method | `apps/powerpoint-vsto/DataLinkManager.cs:131` |
| `Locate` | method | `apps/powerpoint-vsto/DataLinkManager.cs:132` |
| `RefreshObject` | method | `apps/powerpoint-vsto/DataLinkManager.cs:137` |
| `DemoWorkbookLink` | class | `apps/powerpoint-vsto/DemoLinks.cs:13` |
| `DemoLinks` | class | `apps/powerpoint-vsto/DemoLinks.cs:14` |
| `Insert` | method | `apps/powerpoint-vsto/DemoLinks.cs:17` |
| `Selected` | method | `apps/powerpoint-vsto/DemoLinks.cs:32` |
| `Refresh` | method | `apps/powerpoint-vsto/DemoLinks.cs:38` |
| `Describe` | method | `apps/powerpoint-vsto/DemoLinks.cs:45` |
| `Transfer` | method | `apps/powerpoint-vsto/DemoLinks.cs:50` |
