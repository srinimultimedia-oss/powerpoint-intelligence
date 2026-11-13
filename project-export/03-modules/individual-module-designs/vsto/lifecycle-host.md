# VSTO lifecycle, compatibility and S0 native host

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

Create the in-process add-in, gate Office capabilities, manage per-window task panes and retain S0 clustered-column identity while preserving supported native edits.

Implemented local source is distinguished from supported subset/absent external behavior below. See current test evidence in the export testing chapter; do not treat historical logs or method existence as a current runtime pass.

## Source locations and internal components

- `apps/powerpoint-vsto/ThisAddIn.cs:1` — ThisAddIn, IntelligenceRibbon.
- `apps/powerpoint-vsto/ThisAddIn.Designer.cs:1` — ThisAddIn.
- `apps/powerpoint-vsto/ThisAddIn.Designer.xml:1` — project/configuration/generated lifecycle metadata.
- `apps/powerpoint-vsto/OfficeCompatibility.cs:1` — OfficeSettings, OfficeCompatibilityResult, OfficeCompatibility.
- `apps/powerpoint-vsto/PowerPointHost.cs:1` — OfficeMutationLease, OfficeEventBoundary, PowerPointHost.
- `apps/powerpoint-vsto/Contracts.cs:1` — IPowerPointHost, ChartDocument, ChartSeries.
- `apps/powerpoint-vsto/MetadataCodec.cs:1` — MetadataCodec.
- `apps/powerpoint-vsto/PowerPointIntelligence.csproj:1` — project/configuration/generated lifecycle metadata.
- `apps/powerpoint-vsto/App.config:1` — project/configuration/generated lifecycle metadata.

## Inputs, outputs, data structures and dependencies

Input: Office Application/selection, config, validated S0 ChartDocument. Output: chart, pane/status, CustomerData XML and PPI_CHART_ID tag. Dependencies: Office/VSTO PIAs, .NET Framework4.8, WinForms, System.Configuration and routed modules.

## Internal execution and external integration

Startup creates PowerPointHost, a 750 ms layout timer and 650 ms debounced agenda/element timer. Compatibility loads the add-in DLL config and compares Office 16.0 API/build and process bitness. It subscribes selection/double-click/open/save/activation/close; Shutdown unsubscribes/disposes. Run acquires a process-specific named mutation lease, creates/rebinds the active-window pane, applies compatibility gating and catches failures. S0 AddChart2 writes literal-label embedded workbook values and attaches XML/identity. Check & Sync captures native edits and repairs valid duplicate GUIDs. BeforeSave only processes the active deck.

## Configuration, validation and failure behavior

OfficeProfile auto/office2016/office2019/microsoft365, OfficeArchitecture any/x86/x64, OfficeMinimumVersion16.0. Profiles express policy, not automatic edition discovery. Project embeds Office15 interop types but runtime gates on16.0. STA enforced in host; explicitly rejected COM calls0x80010001 retry at100ms. Event exceptions contained; damaged metadata preserved.

## Security considerations and cleanup

Local native calls run in the current Office user process and require supported selection/ownership; a mutex/selection key is not authentication or RBAC. File/metadata content must remain bounded and error text must not expose data or credentials. Explicit owned-resource cleanup and error/rollback paths are part of the cited implementation, not a guarantee that a failed COM server can always be restored.

## Test references

Lexical module-class matches in test source (not instrumentation coverage; no test is claimed executed here):

- `tests/compatibility/OfficeCompatibilityTests.cs:13`.
- `tests/s0/DemoOfficeTests.cs:36`.
- `tests/s0/DemoScenario.cs:114`.
- `tests/s0/HostIntegrationTests.cs:111`.
- `tests/s0/LinkedChartDiagnostic.cs:27`.
- `tests/s0/MappedExcelOfficeTests.cs:23`.
- `tests/s0/MetadataTests.cs:21`.
- `tests/s0/WorkflowOfficeTests.cs:37`.
- `tests/s0/XmlPartUpdateTests.cs:36`.
- `tests/s1/DomainTests.cs:77`.

## Known limitations and technical debt

S0 supports clustered-column only, <=1000 categories, <=100 series, <=1 MiB XML and top-level shape traversal. Structural native type/count/category changes are not migrated. Compensation can fail; use Undo. No separate command-bus/API service startup is wired by this module.

## Extension points

Extend versioned host contracts with active-deck/UI-STA enforcement and native roundtrip tests; named mutex coordination is not an authorization boundary.

## Classes, interfaces and method navigation

Method names below are source declarations; private/internal members are not public API promises. The source hash/review manifest gives exact reviewed bytes.

| Declaration | Kind | Source |
|---|---|---|
| `ThisAddIn` | class | `apps/powerpoint-vsto/ThisAddIn.cs:13` |
| `RequestComAddInAutomationService` | method | `apps/powerpoint-vsto/ThisAddIn.cs:29` |
| `OnStartup` | method | `apps/powerpoint-vsto/ThisAddIn.cs:37` |
| `ConfigureCompatibility` | method | `apps/powerpoint-vsto/ThisAddIn.cs:52` |
| `OnShutdown` | method | `apps/powerpoint-vsto/ThisAddIn.cs:77` |
| `EnsurePane` | method | `apps/powerpoint-vsto/ThisAddIn.cs:88` |
| `RemovePane` | method | `apps/powerpoint-vsto/ThisAddIn.cs:122` |
| `Run` | method | `apps/powerpoint-vsto/ThisAddIn.cs:133` |
| `InsertRenderedShapes` | method | `apps/powerpoint-vsto/ThisAddIn.cs:152` |
| `Insert` | method | `apps/powerpoint-vsto/ThisAddIn.cs:175` |
| `LinkDataSource` | method | `apps/powerpoint-vsto/ThisAddIn.cs:176` |
| `ChangeTemplatePalette` | method | `apps/powerpoint-vsto/ThisAddIn.cs:190` |
| `ChangeTemplateAlignment` | method | `apps/powerpoint-vsto/ThisAddIn.cs:195` |
| `OpenTemplateLibrary` | method | `apps/powerpoint-vsto/ThisAddIn.cs:196` |
| `OpenDemo` | method | `apps/powerpoint-vsto/ThisAddIn.cs:197` |
| `ShowContextProperties` | method | `apps/powerpoint-vsto/ThisAddIn.cs:212` |
| `RunProductivity` | method | `apps/powerpoint-vsto/ThisAddIn.cs:213` |
| `ElementEnabled` | method | `apps/powerpoint-vsto/ThisAddIn.cs:221` |
| `RunElement` | method | `apps/powerpoint-vsto/ThisAddIn.cs:222` |
| `ElementDoubleClick` | method | `apps/powerpoint-vsto/ThisAddIn.cs:223` |
| `WorkflowEnabled` | method | `apps/powerpoint-vsto/ThisAddIn.cs:229` |
| `RunWorkflow` | method | `apps/powerpoint-vsto/ThisAddIn.cs:242` |
| `RunDemo` | method | `apps/powerpoint-vsto/ThisAddIn.cs:270` |
| `OpenLayoutControls` | method | `apps/powerpoint-vsto/ThisAddIn.cs:276` |
| `OpenWorkspace` | method | `apps/powerpoint-vsto/ThisAddIn.cs:277` |
| `Inspect` | method | `apps/powerpoint-vsto/ThisAddIn.cs:278` |
| `ShowCompatibility` | method | `apps/powerpoint-vsto/ThisAddIn.cs:279` |
| `ShowSelection` | method | `apps/powerpoint-vsto/ThisAddIn.cs:286` |
| `SelectionChanged` | method | `apps/powerpoint-vsto/ThisAddIn.cs:296` |
| `RunHostEvent` | method | `apps/powerpoint-vsto/ThisAddIn.cs:301` |
| `Opened` | method | `apps/powerpoint-vsto/ThisAddIn.cs:308` |
| `Activated` | method | `apps/powerpoint-vsto/ThisAddIn.cs:310` |
| `Closed` | method | `apps/powerpoint-vsto/ThisAddIn.cs:316` |
| `BeforeSave` | method | `apps/powerpoint-vsto/ThisAddIn.cs:318` |
| `CreateRibbonExtensibilityObject` | method | `apps/powerpoint-vsto/ThisAddIn.cs:326` |
| `InternalStartup` | method | `apps/powerpoint-vsto/ThisAddIn.cs:328` |
| `IntelligenceRibbon` | class | `apps/powerpoint-vsto/ThisAddIn.cs:332` |
| `GetCustomUI` | method | `apps/powerpoint-vsto/ThisAddIn.cs:338` |
| `LoadImage` | method | `apps/powerpoint-vsto/ThisAddIn.cs:339` |
| `InsertRenderedShapes` | method | `apps/powerpoint-vsto/ThisAddIn.cs:340` |
| `InsertSample` | method | `apps/powerpoint-vsto/ThisAddIn.cs:341` |
| `LinkDataSource` | method | `apps/powerpoint-vsto/ThisAddIn.cs:342` |
| `GetDataSourceEnabled` | method | `apps/powerpoint-vsto/ThisAddIn.cs:343` |
| `InspectChart` | method | `apps/powerpoint-vsto/ThisAddIn.cs:344` |
| `ShowCompatibility` | method | `apps/powerpoint-vsto/ThisAddIn.cs:345` |
| `ShowWorkspace` | method | `apps/powerpoint-vsto/ThisAddIn.cs:346` |
| `ChangeTemplatePalette` | method | `apps/powerpoint-vsto/ThisAddIn.cs:347` |
| `ChangeTemplateAlignment` | method | `apps/powerpoint-vsto/ThisAddIn.cs:348` |
| `ShowLibraryLayout` | method | `apps/powerpoint-vsto/ThisAddIn.cs:349` |
| `ShowTemplateLibrary` | method | `apps/powerpoint-vsto/ThisAddIn.cs:350` |
| `ShowContextProperties` | method | `apps/powerpoint-vsto/ThisAddIn.cs:351` |
| `ShowDemo` | method | `apps/powerpoint-vsto/ThisAddIn.cs:352` |
| `GetWorkflowEnabled` | method | `apps/powerpoint-vsto/ThisAddIn.cs:353` |
| `RunWorkflowCommand` | method | `apps/powerpoint-vsto/ThisAddIn.cs:354` |
| `RunProductivityTool` | method | `apps/powerpoint-vsto/ThisAddIn.cs:355` |
| `GetElementEnabled` | method | `apps/powerpoint-vsto/ThisAddIn.cs:356` |
| `RunElementCommand` | method | `apps/powerpoint-vsto/ThisAddIn.cs:357` |
| `RunDemoTool` | method | `apps/powerpoint-vsto/ThisAddIn.cs:358` |
| `GetCommandEnabled` | method | `apps/powerpoint-vsto/ThisAddIn.cs:359` |
| `OnLoad` | method | `apps/powerpoint-vsto/ThisAddIn.cs:360` |
| `Refresh` | method | `apps/powerpoint-vsto/ThisAddIn.cs:361` |
| `ThisAddIn` | class | `apps/powerpoint-vsto/ThisAddIn.Designer.cs:9` |
| `Initialize` | method | `apps/powerpoint-vsto/ThisAddIn.Designer.cs:19` |
| `FinishInitialization` | method | `apps/powerpoint-vsto/ThisAddIn.Designer.cs:26` |
| `InitializeDataBindings` | method | `apps/powerpoint-vsto/ThisAddIn.Designer.cs:28` |
| `OnShutdown` | method | `apps/powerpoint-vsto/ThisAddIn.Designer.cs:35` |
| `OfficeSettings` | class | `apps/powerpoint-vsto/OfficeCompatibility.cs:9` |
| `Parse` | method | `apps/powerpoint-vsto/OfficeCompatibility.cs:15` |
| `Load` | method | `apps/powerpoint-vsto/OfficeCompatibility.cs:31` |
| `OfficeCompatibilityResult` | class | `apps/powerpoint-vsto/OfficeCompatibility.cs:50` |
| `OfficeCompatibility` | class | `apps/powerpoint-vsto/OfficeCompatibility.cs:59` |
| `Evaluate` | method | `apps/powerpoint-vsto/OfficeCompatibility.cs:61` |
| `Normalize` | method | `apps/powerpoint-vsto/OfficeCompatibility.cs:88` |
| `Deny` | method | `apps/powerpoint-vsto/OfficeCompatibility.cs:90` |
| `OfficeMutationLease` | class | `apps/powerpoint-vsto/PowerPointHost.cs:14` |
| `GetWindowThreadProcessId` | method | `apps/powerpoint-vsto/PowerPointHost.cs:19` |
| `TryEnter` | method | `apps/powerpoint-vsto/PowerPointHost.cs:20` |
| `ForAutomation` | method | `apps/powerpoint-vsto/PowerPointHost.cs:27` |
| `Dispose` | method | `apps/powerpoint-vsto/PowerPointHost.cs:32` |
| `OfficeEventBoundary` | class | `apps/powerpoint-vsto/PowerPointHost.cs:34` |
| `Run` | method | `apps/powerpoint-vsto/PowerPointHost.cs:37` |
| `PowerPointHost` | class | `apps/powerpoint-vsto/PowerPointHost.cs:48` |
| `CheckThread` | method | `apps/powerpoint-vsto/PowerPointHost.cs:58` |
| `GetSelectedSlideId` | method | `apps/powerpoint-vsto/PowerPointHost.cs:65` |
| `GetSelectedObjectIds` | method | `apps/powerpoint-vsto/PowerPointHost.cs:66` |
| `Shapes` | method | `apps/powerpoint-vsto/PowerPointHost.cs:80` |
| `Part` | method | `apps/powerpoint-vsto/PowerPointHost.cs:90` |
| `Read` | method | `apps/powerpoint-vsto/PowerPointHost.cs:105` |
| `Write` | method | `apps/powerpoint-vsto/PowerPointHost.cs:116` |
| `InsertEditableChart` | method | `apps/powerpoint-vsto/PowerPointHost.cs:166` |
| `SetData` | method | `apps/powerpoint-vsto/PowerPointHost.cs:185` |
| `Find` | method | `apps/powerpoint-vsto/PowerPointHost.cs:220` |
| `ReadChartMetadata` | method | `apps/powerpoint-vsto/PowerPointHost.cs:226` |
| `CopyEditableChartToActiveSlide` | method | `apps/powerpoint-vsto/PowerPointHost.cs:231` |
| `SaveChartMetadata` | method | `apps/powerpoint-vsto/PowerPointHost.cs:275` |
| `UpdateEditableChart` | method | `apps/powerpoint-vsto/PowerPointHost.cs:282` |
| `Capture` | method | `apps/powerpoint-vsto/PowerPointHost.cs:296` |
| `NativeSeriesValues` | method | `apps/powerpoint-vsto/PowerPointHost.cs:331` |
| `SeriesVector` | method | `apps/powerpoint-vsto/PowerPointHost.cs:335` |
| `RetryRejectedCall` | method | `apps/powerpoint-vsto/PowerPointHost.cs:345` |
| `InspectAndRecover` | method | `apps/powerpoint-vsto/PowerPointHost.cs:360` |
| `IPowerPointHost` | interface | `apps/powerpoint-vsto/Contracts.cs:7` |
| `ChartDocument` | class | `apps/powerpoint-vsto/Contracts.cs:17` |
| `ChartSeries` | class | `apps/powerpoint-vsto/Contracts.cs:27` |
| `MetadataCodec` | class | `apps/powerpoint-vsto/MetadataCodec.cs:10` |
| `Validate` | method | `apps/powerpoint-vsto/MetadataCodec.cs:16` |
| `Encode` | method | `apps/powerpoint-vsto/MetadataCodec.cs:34` |
| `Decode` | method | `apps/powerpoint-vsto/MetadataCodec.cs:46` |
| `Sample` | method | `apps/powerpoint-vsto/MetadataCodec.cs:60` |
