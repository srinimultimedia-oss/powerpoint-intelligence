# Module dependency matrix

**INFERRED static reference map:** edges below are exact shared-symbol references/import neighborhoods, not a compiler-resolved callgraph. Shared source inclusion and type names can produce non-call relationships. The component design prose explains consequential runtime direction; test/build paths are not assumed reachable product features.

| From | Referenced module | Evidence source |
|---|---|---|
| OPS-01 | OPS-05 | `tests/s0/ServiceConnectionTests.cs:1` |
| OPS-01 | OPS-06 | `tests/s0/ServiceConnectionTests.cs:1` |
| OPS-01 | VSTO-08 | `apps/powerpoint-vsto/ServiceConnections.cs:1` |
| OPS-01 | VSTO-10 | `apps/powerpoint-vsto/ServiceConnections.cs:1` |
| OPS-01 | VSTO-18 | `apps/powerpoint-vsto/ServiceConnections.cs:1`; `tests/s0/ServiceConnectionTests.cs:1` |
| OPS-01 | chart-calculations | `apps/powerpoint-vsto/ServiceConnections.cs:1`; `tests/s0/ServiceConnectionTests.cs:1` |
| OPS-01 | path-policy | `apps/powerpoint-vsto/ServiceConnections.cs:1`; `tests/s0/ServiceConnectionTests.cs:1` |
| OPS-01 | ts-scene | `apps/powerpoint-vsto/ServiceConnections.cs:1`; `tests/s0/ServiceConnectionTests.cs:1` |
| OPS-02 | VSTO-01 | `apps/powerpoint-vsto/Contracts.cs:1`; `apps/powerpoint-vsto/DemoApi.cs:1` |
| OPS-02 | VSTO-02 | `apps/powerpoint-vsto/DemoApi.cs:1` |
| OPS-02 | chart-calculations | `apps/powerpoint-vsto/Contracts.cs:1` |
| OPS-02 | domain-v2 | `apps/powerpoint-vsto/Contracts.cs:1`; `apps/powerpoint-vsto/DemoApi.cs:1` |
| OPS-03 | OPS-04 | `easySetup/SetupPolicyTests.cs:1` |
| OPS-03 | OPS-05 | `easySetup/Build-EasySetup.ps1:1`; `easySetup/Common.ps1:1`; `easySetup/SetupPolicyTests.cs:1`; `scripts/Install-Local.ps1:1` |
| OPS-03 | OPS-06 | `easySetup/SetupPolicyTests.cs:1` |
| OPS-03 | VSTO-02 | `easySetup/Build-EasySetup.ps1:1`; `scripts/Test-DemoPackage.ps1:1`; `scripts/Test-PayloadRecompile.ps1:1` |
| OPS-03 | VSTO-04 | `scripts/Install-Local.ps1:1` |
| OPS-03 | VSTO-06 | `easySetup/Test-EasySetup.ps1:1` |
| OPS-03 | chart-calculations | `easySetup/Build-EasySetup.ps1:1` |
| OPS-03 | svg-preview | `easySetup/SetupPolicyTests.cs:1` |
| OPS-04 | OPS-05 | `easySetup/Bootstrapper.cs:1` |
| OPS-04 | chart-calculations | `easySetup/Bootstrapper.cs:1` |
| OPS-04 | ts-scene | `easySetup/Bootstrapper.cs:1` |
| OPS-05 | OPS-06 | `scripts/Scene-RunnerCache.ps1:1`; `scripts/Test-BuildIntegrity.ps1:1`; `scripts/Test-Checkmarx113558.ps1:1`; `scripts/Test-PiidxRecovery.ps1:1`; `scripts/Test-SceneRunnerCache.ps1:1`; `scripts/Test-SourceFingerprint.ps1:1` |
| OPS-05 | VSTO-05 | `scripts/Piidx-Recovery.ps1:1`; `scripts/Scene-RunnerCache.ps1:1` |
| OPS-05 | VSTO-06 | `scripts/Piidx-Recovery.ps1:1`; `scripts/Scene-RunnerCache.ps1:1`; `scripts/Test-Checkmarx113558.ps1:1`; `scripts/Test-PiidxRecovery.ps1:1`; `scripts/Test-SceneRunnerCache.ps1:1` |
| OPS-05 | chart-calculations | `scripts/Build-Integrity.ps1:1`; `scripts/Piidx-Recovery.ps1:1`; `scripts/Piidx-TestSuite.ps1:1`; `scripts/Restore-LockedDependencies.ps1:1`; `scripts/Scene-RunnerCache.ps1:1`; `scripts/Source-Fingerprint.ps1:1`; `scripts/Test-SourceFingerprint.ps1:1` |
| OPS-05 | path-policy | `scripts/Scene-RunnerCache.ps1:1` |
| OPS-05 | scene-contract | `scripts/Scene-RunnerCache.ps1:1` |
| OPS-05 | strict-json | `scripts/Scene-RunnerCache.ps1:1` |
| OPS-05 | ts-scene | `scripts/Piidx-Recovery.ps1:1`; `scripts/Scene-RunnerCache.ps1:1` |
| OPS-06 | OPS-05 | `scripts/Collect-SceneThreatEvidence.ps1:1`; `scripts/Diagnose-PowerPointActivation.ps1:1`; `scripts/Test-DiagnosticOutputPath.ps1:1` |
| OPS-06 | VSTO-06 | `scripts/Collect-SceneThreatEvidence.ps1:1`; `scripts/Test-DiagnosticOutputPath.ps1:1` |
| OPS-06 | chart-calculations | `scripts/Collect-SceneThreatEvidence.ps1:1`; `scripts/Diagnostic-OutputPath.ps1:1`; `scripts/Test-DiagnosticOutputPath.ps1:1` |
| OPS-06 | svg-preview | `scripts/Test-DiagnosticOutputPath.ps1:1` |
| OPS-06 | ts-scene | `scripts/Collect-SceneThreatEvidence.ps1:1` |
| OPS-07 | OPS-01 | `scripts/Test-DemoFeatures.ps1:1`; `scripts/Test-Ribbon.ps1:1` |
| OPS-07 | OPS-02 | `scripts/Test-NativeRenderer.ps1:1`; `scripts/Test-S1.ps1:1` |
| OPS-07 | OPS-05 | `scripts/New-PiidxSample.ps1:1`; `scripts/Test-All.ps1:1`; `scripts/Test-CheckProcess.ps1:1`; `scripts/Test-InstalledRibbon.ps1:1`; `scripts/Test-NativeRenderer.ps1:1`; `scripts/Test-Piidx.ps1:1`; `scripts/Test-PiidxCatalog.ps1:1`; `scripts/Test-PiidxCorpus.ps1:1`; `scripts/Test-PiidxFoundation.ps1:1`; `scripts/Test-PiidxPreviews.ps1:1`; `scripts/Test-PiidxSourceReader.ps1:1`; `scripts/Test-PiidxSqliteCapability.ps1:1`; `scripts/Test-SceneContracts.ps1:1`; `scripts/Test-TemplateLibrary.ps1:1` |
| OPS-07 | OPS-06 | `scripts/Test-All.ps1:1`; `scripts/Test-CheckProcess.ps1:1`; `scripts/Test-InstalledRibbon.ps1:1` |
| OPS-07 | VSTO-01 | `scripts/Test-All.ps1:1`; `scripts/Test-DemoFeatures.ps1:1`; `scripts/Test-NativeRenderer.ps1:1`; `scripts/Test-OfficeCompatibility.ps1:1`; `scripts/Test-Ribbon.ps1:1`; `scripts/Test-S0.ps1:1`; `scripts/Test-S1.ps1:1` |
| OPS-07 | VSTO-02 | `scripts/Test-DemoFeatures.ps1:1`; `scripts/Test-InstalledRibbon.ps1:1`; `scripts/Test-Ribbon.ps1:1` |
| OPS-07 | VSTO-03 | `scripts/Test-DemoFeatures.ps1:1`; `scripts/Test-Ribbon.ps1:1` |
| OPS-07 | VSTO-04 | `scripts/Test-All.ps1:1`; `scripts/Test-DemoFeatures.ps1:1`; `scripts/Test-Ribbon.ps1:1` |
| OPS-07 | VSTO-05 | `scripts/Test-DemoFeatures.ps1:1`; `scripts/Test-Ribbon.ps1:1` |
| OPS-07 | VSTO-06 | `scripts/Test-DemoFeatures.ps1:1`; `scripts/Test-Ribbon.ps1:1` |
| OPS-07 | VSTO-07 | `scripts/Test-DemoFeatures.ps1:1`; `scripts/Test-InstalledRibbon.ps1:1`; `scripts/Test-Ribbon.ps1:1` |
| OPS-07 | VSTO-08 | `scripts/Test-DemoFeatures.ps1:1`; `scripts/Test-InstalledRibbon.ps1:1`; `scripts/Test-Ribbon.ps1:1` |
| OPS-07 | VSTO-09 | `scripts/Test-DemoFeatures.ps1:1`; `scripts/Test-Ribbon.ps1:1` |
| OPS-07 | VSTO-10 | `scripts/Test-DemoFeatures.ps1:1`; `scripts/Test-InstalledRibbon.ps1:1`; `scripts/Test-Ribbon.ps1:1` |
| OPS-07 | VSTO-11 | `scripts/Test-DemoFeatures.ps1:1`; `scripts/Test-Ribbon.ps1:1` |
| OPS-07 | VSTO-12 | `scripts/Test-DemoFeatures.ps1:1`; `scripts/Test-Ribbon.ps1:1`; `scripts/Test-TemplateLibrary.ps1:1` |
| OPS-07 | VSTO-13 | `scripts/Test-All.ps1:1`; `scripts/Test-DemoFeatures.ps1:1`; `scripts/Test-Ribbon.ps1:1`; `scripts/Test-TemplateLibrary.ps1:1` |
| OPS-07 | VSTO-15 | `scripts/Test-DemoFeatures.ps1:1`; `scripts/Test-InstalledRibbon.ps1:1`; `scripts/Test-Ribbon.ps1:1` |
| OPS-07 | VSTO-16 | `scripts/Test-DemoFeatures.ps1:1`; `scripts/Test-Ribbon.ps1:1` |
| OPS-07 | VSTO-17 | `scripts/Test-DemoFeatures.ps1:1`; `scripts/Test-Ribbon.ps1:1` |
| OPS-07 | VSTO-18 | `scripts/Test-DemoFeatures.ps1:1`; `scripts/Test-Ribbon.ps1:1` |
| OPS-07 | VSTO-19 | `scripts/Test-NativeRenderer.ps1:1` |
| OPS-07 | chart-calculations | `scripts/Check-Process.ps1:1`; `scripts/Test-All.ps1:1`; `scripts/Test-CheckProcess.ps1:1`; `scripts/Test-InstalledRibbon.ps1:1`; `scripts/Test-NativeRenderer.ps1:1`; `scripts/Test-Piidx.ps1:1`; `scripts/Test-S0.ps1:1`; `scripts/Test-S1.ps1:1`; `scripts/Test-SceneContracts.ps1:1` |
| OPS-07 | domain-v2 | `scripts/Test-NativeRenderer.ps1:1`; `scripts/Test-S1.ps1:1` |
| OPS-07 | native-renderer | `scripts/Test-NativeRenderer.ps1:1` |
| OPS-07 | path-policy | `scripts/Test-NativeRenderer.ps1:1`; `scripts/Test-OfficeCompatibility.ps1:1`; `scripts/Test-S0.ps1:1`; `scripts/Test-S1.ps1:1`; `scripts/Test-SceneContracts.ps1:1`; `scripts/Test-SecurityRemediation.ps1:1` |
| OPS-07 | scene-contract | `scripts/Test-NativeRenderer.ps1:1`; `scripts/Test-SceneContracts.ps1:1` |
| OPS-07 | strict-json | `scripts/Test-NativeRenderer.ps1:1`; `scripts/Test-SceneContracts.ps1:1` |
| OPS-07 | svg-preview | `scripts/Test-SecurityRemediation.ps1:1` |
| OPS-07 | ts-scene | `scripts/Test-All.ps1:1`; `scripts/Test-DemoFeatures.ps1:1` |
| OPS-09 | VSTO-05 | `scripts/Clean-Workspace.ps1:1` |
| OPS-09 | chart-calculations | `scripts/Clean-Workspace.ps1:1` |
| OPS-09 | ts-scene | `scripts/Clean-Workspace.ps1:1` |
| VSTO-01 | OPS-01 | `apps/powerpoint-vsto/ThisAddIn.cs:1` |
| VSTO-01 | OPS-02 | `apps/powerpoint-vsto/Contracts.cs:1`; `apps/powerpoint-vsto/MetadataCodec.cs:1`; `apps/powerpoint-vsto/PowerPointHost.cs:1`; `apps/powerpoint-vsto/ThisAddIn.cs:1` |
| VSTO-01 | OPS-05 | `apps/powerpoint-vsto/ThisAddIn.cs:1` |
| VSTO-01 | OPS-06 | `apps/powerpoint-vsto/ThisAddIn.Designer.cs:1`; `apps/powerpoint-vsto/ThisAddIn.cs:1` |
| VSTO-01 | VSTO-02 | `apps/powerpoint-vsto/ThisAddIn.cs:1` |
| VSTO-01 | VSTO-03 | `apps/powerpoint-vsto/ThisAddIn.cs:1` |
| VSTO-01 | VSTO-04 | `apps/powerpoint-vsto/ThisAddIn.cs:1` |
| VSTO-01 | VSTO-05 | `apps/powerpoint-vsto/ThisAddIn.cs:1` |
| VSTO-01 | VSTO-06 | `apps/powerpoint-vsto/ThisAddIn.cs:1` |
| VSTO-01 | VSTO-07 | `apps/powerpoint-vsto/ThisAddIn.cs:1` |
| VSTO-01 | VSTO-09 | `apps/powerpoint-vsto/ThisAddIn.cs:1` |
| VSTO-01 | VSTO-10 | `apps/powerpoint-vsto/ThisAddIn.cs:1` |
| VSTO-01 | VSTO-11 | `apps/powerpoint-vsto/ThisAddIn.cs:1` |
| VSTO-01 | VSTO-12 | `apps/powerpoint-vsto/ThisAddIn.cs:1` |
| VSTO-01 | VSTO-13 | `apps/powerpoint-vsto/ThisAddIn.cs:1` |
| VSTO-01 | VSTO-14 | `apps/powerpoint-vsto/ThisAddIn.cs:1` |
| VSTO-01 | VSTO-15 | `apps/powerpoint-vsto/ThisAddIn.cs:1` |
| VSTO-01 | VSTO-16 | `apps/powerpoint-vsto/ThisAddIn.cs:1` |
| VSTO-01 | VSTO-18 | `apps/powerpoint-vsto/ThisAddIn.cs:1` |
| VSTO-01 | VSTO-19 | `apps/powerpoint-vsto/ThisAddIn.cs:1` |
| VSTO-01 | chart-calculations | `apps/powerpoint-vsto/Contracts.cs:1`; `apps/powerpoint-vsto/MetadataCodec.cs:1`; `apps/powerpoint-vsto/PowerPointHost.cs:1`; `apps/powerpoint-vsto/ThisAddIn.Designer.cs:1`; `apps/powerpoint-vsto/ThisAddIn.cs:1` |
| VSTO-01 | domain-v2 | `apps/powerpoint-vsto/Contracts.cs:1`; `apps/powerpoint-vsto/MetadataCodec.cs:1`; `apps/powerpoint-vsto/PowerPointHost.cs:1`; `apps/powerpoint-vsto/ThisAddIn.cs:1` |
| VSTO-01 | path-policy | `apps/powerpoint-vsto/OfficeCompatibility.cs:1` |
| VSTO-01 | svg-preview | `apps/powerpoint-vsto/PowerPointHost.cs:1` |
| VSTO-02 | OPS-02 | `apps/powerpoint-vsto/DemoApi.cs:1`; `apps/powerpoint-vsto/DemoTools.cs:1` |
| VSTO-02 | OPS-05 | `apps/powerpoint-vsto/DemoRibbon.cs:1` |
| VSTO-02 | OPS-06 | `apps/powerpoint-vsto/DemoRibbon.cs:1` |
| VSTO-02 | VSTO-01 | `apps/powerpoint-vsto/DemoApi.cs:1`; `apps/powerpoint-vsto/DemoTools.cs:1` |
| VSTO-02 | VSTO-03 | `apps/powerpoint-vsto/DemoRibbon.cs:1`; `apps/powerpoint-vsto/DemoTools.cs:1` |
| VSTO-02 | VSTO-04 | `apps/powerpoint-vsto/DemoTools.cs:1`; `apps/powerpoint-vsto/FeatureInventory.cs:1` |
| VSTO-02 | VSTO-05 | `apps/powerpoint-vsto/RibbonIcons.cs:1` |
| VSTO-02 | VSTO-07 | `apps/powerpoint-vsto/DemoTools.cs:1` |
| VSTO-02 | VSTO-08 | `apps/powerpoint-vsto/DemoTools.cs:1`; `apps/powerpoint-vsto/FeatureInventory.cs:1` |
| VSTO-02 | VSTO-09 | `apps/powerpoint-vsto/DemoTools.cs:1` |
| VSTO-02 | VSTO-10 | `apps/powerpoint-vsto/DemoTools.cs:1` |
| VSTO-02 | VSTO-11 | `apps/powerpoint-vsto/DemoTools.cs:1` |
| VSTO-02 | VSTO-15 | `apps/powerpoint-vsto/DemoTools.cs:1` |
| VSTO-02 | chart-calculations | `apps/powerpoint-vsto/DemoRibbon.cs:1`; `apps/powerpoint-vsto/DemoTools.cs:1`; `apps/powerpoint-vsto/DemoWorkspace.cs:1`; `apps/powerpoint-vsto/FeatureInventory.cs:1`; `apps/powerpoint-vsto/RibbonIcons.cs:1` |
| VSTO-02 | domain-v2 | `apps/powerpoint-vsto/DemoApi.cs:1`; `apps/powerpoint-vsto/DemoTools.cs:1` |
| VSTO-02 | svg-preview | `apps/powerpoint-vsto/DemoTools.cs:1` |
| VSTO-02 | ts-scene | `apps/powerpoint-vsto/DemoTools.cs:1`; `apps/powerpoint-vsto/FeatureInventory.cs:1` |
| VSTO-03 | OPS-05 | `apps/powerpoint-vsto/ExcelRangeMapping.cs:1` |
| VSTO-03 | OPS-06 | `apps/powerpoint-vsto/ExcelRangeMapping.cs:1` |
| VSTO-03 | VSTO-01 | `apps/powerpoint-vsto/DataSources.cs:1`; `apps/powerpoint-vsto/DemoLinks.cs:1` |
| VSTO-03 | VSTO-04 | `apps/powerpoint-vsto/MappedLinkOperations.cs:1` |
| VSTO-03 | VSTO-17 | `apps/powerpoint-vsto/DataSources.cs:1` |
| VSTO-03 | chart-calculations | `apps/powerpoint-vsto/DataLinkManager.cs:1`; `apps/powerpoint-vsto/DataSources.cs:1`; `apps/powerpoint-vsto/DemoLinks.cs:1`; `apps/powerpoint-vsto/ExcelRangeMapping.cs:1`; `apps/powerpoint-vsto/MappedLinkOperations.cs:1` |
| VSTO-03 | path-policy | `apps/powerpoint-vsto/DataSources.cs:1`; `apps/powerpoint-vsto/DemoLinks.cs:1`; `apps/powerpoint-vsto/ExcelRangeMapping.cs:1` |
| VSTO-03 | svg-preview | `apps/powerpoint-vsto/DataSources.cs:1`; `apps/powerpoint-vsto/DemoLinks.cs:1` |
| VSTO-03 | ts-scene | `apps/powerpoint-vsto/DataLinkManager.cs:1`; `apps/powerpoint-vsto/DataSources.cs:1`; `apps/powerpoint-vsto/DemoLinks.cs:1`; `apps/powerpoint-vsto/ExcelRangeMapping.cs:1` |
| VSTO-04 | VSTO-01 | `apps/powerpoint-vsto/NativeSeriesControls.cs:1` |
| VSTO-04 | VSTO-08 | `apps/powerpoint-vsto/ChartAnnotations.cs:1` |
| VSTO-04 | chart-calculations | `apps/powerpoint-vsto/ChartAnnotations.cs:1`; `apps/powerpoint-vsto/NativeChartControls.cs:1`; `apps/powerpoint-vsto/NativeSeriesControls.cs:1` |
| VSTO-04 | svg-preview | `apps/powerpoint-vsto/NativeChartControls.cs:1` |
| VSTO-04 | ts-scene | `apps/powerpoint-vsto/ChartAnnotations.cs:1`; `apps/powerpoint-vsto/NativeChartControls.cs:1` |
| VSTO-05 | VSTO-02 | `apps/powerpoint-vsto/ContextActionRouter.cs:1` |
| VSTO-05 | VSTO-03 | `apps/powerpoint-vsto/ContextActionRouter.cs:1` |
| VSTO-05 | VSTO-04 | `apps/powerpoint-vsto/ContextActionRouter.cs:1`; `apps/powerpoint-vsto/ContextFormatting.cs:1`; `apps/powerpoint-vsto/ContextPanelModel.cs:1`; `apps/powerpoint-vsto/ContextPropertiesController.cs:1`; `apps/powerpoint-vsto/ContextPropertiesView.cs:1`; `apps/powerpoint-vsto/ContextSelectionBridge.cs:1` |
| VSTO-05 | VSTO-06 | `apps/powerpoint-vsto/ContextChartFormatting.cs:1`; `apps/powerpoint-vsto/ContextFormatting.cs:1`; `apps/powerpoint-vsto/ContextPropertiesController.cs:1`; `apps/powerpoint-vsto/ContextSelectionBridge.cs:1` |
| VSTO-05 | VSTO-07 | `apps/powerpoint-vsto/ContextActionRouter.cs:1`; `apps/powerpoint-vsto/ContextSelectionBridge.cs:1` |
| VSTO-05 | VSTO-08 | `apps/powerpoint-vsto/ContextPanelModel.cs:1`; `apps/powerpoint-vsto/ContextSelectionBridge.cs:1` |
| VSTO-05 | VSTO-09 | `apps/powerpoint-vsto/ContextSelectionBridge.cs:1` |
| VSTO-05 | VSTO-10 | `apps/powerpoint-vsto/ContextSelectionBridge.cs:1` |
| VSTO-05 | VSTO-11 | `apps/powerpoint-vsto/ContextActionRouter.cs:1`; `apps/powerpoint-vsto/ContextSelectionBridge.cs:1` |
| VSTO-05 | VSTO-13 | `apps/powerpoint-vsto/ContextActionRouter.cs:1`; `apps/powerpoint-vsto/ContextSelectionBridge.cs:1` |
| VSTO-05 | chart-calculations | `apps/powerpoint-vsto/ContextChartFormatting.cs:1`; `apps/powerpoint-vsto/ContextPanelModel.cs:1`; `apps/powerpoint-vsto/ContextPropertiesController.cs:1`; `apps/powerpoint-vsto/ContextPropertiesView.cs:1`; `apps/powerpoint-vsto/ContextSelectionBridge.cs:1` |
| VSTO-05 | ts-scene | `apps/powerpoint-vsto/ContextPropertiesView.cs:1` |
| VSTO-06 | VSTO-05 | `apps/powerpoint-vsto/NativeFormatting.cs:1` |
| VSTO-06 | chart-calculations | `apps/powerpoint-vsto/NativeFormatting.cs:1` |
| VSTO-07 | VSTO-06 | `apps/powerpoint-vsto/TextShapeElements.cs:1` |
| VSTO-07 | VSTO-08 | `apps/powerpoint-vsto/ElementCommands.cs:1`; `apps/powerpoint-vsto/TextShapeElements.cs:1` |
| VSTO-07 | VSTO-09 | `apps/powerpoint-vsto/ElementCommands.cs:1` |
| VSTO-07 | VSTO-10 | `apps/powerpoint-vsto/ElementCommands.cs:1` |
| VSTO-07 | VSTO-17 | `apps/powerpoint-vsto/TextShapeElements.cs:1` |
| VSTO-07 | chart-calculations | `apps/powerpoint-vsto/ElementCommands.cs:1`; `apps/powerpoint-vsto/TextShapeElements.cs:1` |
| VSTO-07 | ts-scene | `apps/powerpoint-vsto/ElementCommands.cs:1`; `apps/powerpoint-vsto/TextShapeElements.cs:1` |
| VSTO-08 | VSTO-04 | `apps/powerpoint-vsto/ProcessFlowTools.cs:1` |
| VSTO-08 | VSTO-06 | `apps/powerpoint-vsto/ProcessFlowTools.cs:1` |
| VSTO-08 | VSTO-07 | `apps/powerpoint-vsto/ProcessFlowTools.cs:1` |
| VSTO-08 | VSTO-10 | `apps/powerpoint-vsto/ProcessFlowTools.cs:1` |
| VSTO-08 | VSTO-17 | `apps/powerpoint-vsto/ProcessFlowTools.cs:1` |
| VSTO-08 | ts-scene | `apps/powerpoint-vsto/ProcessFlowTools.cs:1` |
| VSTO-09 | VSTO-17 | `apps/powerpoint-vsto/IndicatorElements.cs:1` |
| VSTO-09 | chart-calculations | `apps/powerpoint-vsto/IndicatorElements.cs:1` |
| VSTO-10 | VSTO-04 | `apps/powerpoint-vsto/UniversalConnectors.cs:1` |
| VSTO-10 | VSTO-08 | `apps/powerpoint-vsto/UniversalConnectors.cs:1` |
| VSTO-10 | VSTO-17 | `apps/powerpoint-vsto/UniversalConnectors.cs:1` |
| VSTO-10 | chart-calculations | `apps/powerpoint-vsto/UniversalConnectors.cs:1` |
| VSTO-10 | ts-scene | `apps/powerpoint-vsto/UniversalConnectors.cs:1` |
| VSTO-11 | VSTO-17 | `apps/powerpoint-vsto/ChapterAgenda.cs:1` |
| VSTO-11 | chart-calculations | `apps/powerpoint-vsto/ChapterAgenda.cs:1`; `apps/powerpoint-vsto/ChapterAgendaModel.cs:1`; `apps/powerpoint-vsto/PresentationAgenda.cs:1` |
| VSTO-11 | svg-preview | `apps/powerpoint-vsto/ChapterAgenda.cs:1` |
| VSTO-11 | ts-scene | `apps/powerpoint-vsto/ChapterAgenda.cs:1`; `apps/powerpoint-vsto/ChapterAgendaModel.cs:1`; `apps/powerpoint-vsto/PresentationAgenda.cs:1` |
| VSTO-12 | OPS-05 | `apps/powerpoint-vsto/TemplateCatalog.cs:1` |
| VSTO-12 | VSTO-05 | `apps/powerpoint-vsto/TemplateCatalog.cs:1` |
| VSTO-12 | VSTO-06 | `apps/powerpoint-vsto/TemplateCatalog.cs:1` |
| VSTO-12 | chart-calculations | `apps/powerpoint-vsto/TemplateCatalog.cs:1` |
| VSTO-12 | ts-scene | `apps/powerpoint-vsto/TemplateCatalog.cs:1` |
| VSTO-13 | VSTO-05 | `apps/powerpoint-vsto/TemplateLibrary.cs:1` |
| VSTO-13 | VSTO-06 | `apps/powerpoint-vsto/TemplateLibrary.cs:1` |
| VSTO-13 | VSTO-12 | `apps/powerpoint-vsto/TemplateLibrary.cs:1` |
| VSTO-13 | VSTO-14 | `apps/powerpoint-vsto/TemplateLibrary.cs:1` |
| VSTO-13 | catalog-query | `apps/powerpoint-vsto/TemplateLibrary.cs:1` |
| VSTO-13 | chart-calculations | `apps/powerpoint-vsto/TemplateLibrary.cs:1` |
| VSTO-13 | source-extraction | `apps/powerpoint-vsto/TemplateLibrary.cs:1` |
| VSTO-13 | ts-scene | `apps/powerpoint-vsto/TemplateLibrary.cs:1` |
| VSTO-14 | VSTO-13 | `apps/powerpoint-vsto/LocalLibraryGallery.cs:1`; `apps/powerpoint-vsto/LocalLibraryWorkspace.cs:1` |
| VSTO-14 | catalog | `apps/powerpoint-vsto/LocalLibraryGallery.cs:1`; `apps/powerpoint-vsto/LocalLibraryWorkspace.cs:1` |
| VSTO-14 | catalog-query | `apps/powerpoint-vsto/LocalLibraryGallery.cs:1`; `apps/powerpoint-vsto/LocalLibraryWorkspace.cs:1` |
| VSTO-14 | chart-calculations | `apps/powerpoint-vsto/LocalLibraryGallery.cs:1`; `apps/powerpoint-vsto/LocalLibraryWorkspace.cs:1` |
| VSTO-14 | library-config | `apps/powerpoint-vsto/LocalLibraryGallery.cs:1`; `apps/powerpoint-vsto/LocalLibraryWorkspace.cs:1` |
| VSTO-14 | library-refresh | `apps/powerpoint-vsto/LocalLibraryWorkspace.cs:1` |
| VSTO-14 | preview-cache | `apps/powerpoint-vsto/LocalLibraryGallery.cs:1`; `apps/powerpoint-vsto/LocalLibraryWorkspace.cs:1` |
| VSTO-15 | chart-calculations | `apps/powerpoint-vsto/ProductivityTools.cs:1` |
| VSTO-16 | chart-calculations | `apps/powerpoint-vsto/PresentationQuality.cs:1` |
| VSTO-17 | chart-calculations | `apps/powerpoint-vsto/XmlPartUpdates.cs:1` |
| VSTO-18 | OPS-01 | `apps/powerpoint-vsto/ServiceConnections.cs:1` |
| VSTO-18 | VSTO-08 | `apps/powerpoint-vsto/ServiceConnections.cs:1` |
| VSTO-18 | VSTO-10 | `apps/powerpoint-vsto/ServiceConnections.cs:1` |
| VSTO-18 | chart-calculations | `apps/powerpoint-vsto/ServiceConnections.cs:1` |
| VSTO-18 | path-policy | `apps/powerpoint-vsto/ServiceConnections.cs:1` |
| VSTO-18 | ts-scene | `apps/powerpoint-vsto/ServiceConnections.cs:1` |
| VSTO-19 | OPS-02 | `apps/powerpoint-vsto/RendererRibbonCommand.cs:1` |
| VSTO-19 | VSTO-01 | `apps/powerpoint-vsto/RendererRibbonCommand.cs:1` |
| VSTO-19 | chart-calculations | `apps/powerpoint-vsto/RendererRibbonCommand.cs:1` |
| VSTO-19 | domain-v2 | `apps/powerpoint-vsto/RendererRibbonCommand.cs:1` |
| VSTO-19 | native-renderer | `apps/powerpoint-vsto/RendererRibbonCommand.cs:1` |
| VSTO-19 | scene-contract | `apps/powerpoint-vsto/RendererRibbonCommand.cs:1` |
| VSTO-19 | ts-scene | `apps/powerpoint-vsto/RendererRibbonCommand.cs:1` |
| catalog | OPS-05 | `src/Library.Core/PiidxCatalog.cs:1` |
| catalog | catalog-query | `src/Library.Core/PiidxCatalog.cs:1` |
| catalog | chart-calculations | `src/Library.Core/PiidxCatalog.cs:1` |
| catalog | path-policy | `src/Library.Core/PiidxCatalog.cs:1` |
| catalog | piidx-schema | `src/Library.Core/PiidxCatalog.cs:1` |
| catalog | preview-cache | `src/Library.Core/PiidxCatalog.cs:1` |
| catalog | source-extraction | `src/Library.Core/PiidxCatalog.cs:1` |
| catalog | sqlite | `src/Library.Core/PiidxCatalog.cs:1` |
| catalog | ts-scene | `src/Library.Core/PiidxCatalog.cs:1` |
| catalog-query | catalog | `src/Library.Core/PiidxNavigation.cs:1`; `src/Library.Core/PiidxSearch.cs:1` |
| catalog-query | chart-calculations | `src/Library.Core/LibraryModels.cs:1`; `src/Library.Core/PiidxNavigation.cs:1`; `src/Library.Core/PiidxSearch.cs:1` |
| catalog-query | preview-cache | `src/Library.Core/PiidxNavigation.cs:1`; `src/Library.Core/PiidxSearch.cs:1` |
| catalog-query | source-extraction | `src/Library.Core/PiidxSearch.cs:1` |
| chart-calculations | svg-preview | `packages/chart-core/src/calculations.mjs:1`; `packages/chart-engine/src/basic-charts.ts:1` |
| chart-calculations | ts-scene | `packages/chart-core/src/calculations.mjs:1`; `packages/chart-engine/src/basic-charts.ts:1` |
| domain-v2 | OPS-02 | `src/Chart.Domain/ChartCodec.cs:1`; `src/Chart.Domain/ChartDocument.cs:1`; `src/Chart.Domain/S0Migration.cs:1` |
| domain-v2 | OPS-05 | `src/Chart.Domain/ChartCodec.cs:1` |
| domain-v2 | OPS-06 | `src/Chart.Domain/ChartCodec.cs:1` |
| domain-v2 | VSTO-01 | `src/Chart.Domain/ChartCodec.cs:1`; `src/Chart.Domain/ChartDocument.cs:1`; `src/Chart.Domain/S0Migration.cs:1` |
| domain-v2 | chart-calculations | `src/Chart.Domain/ChartCodec.cs:1`; `src/Chart.Domain/ChartDocument.cs:1`; `src/Chart.Domain/S0Migration.cs:1` |
| domain-v2 | svg-preview | `src/Chart.Domain/ChartCodec.cs:1` |
| domain-v2 | ts-scene | `src/Chart.Domain/ChartCodec.cs:1` |
| domain-v3 | OPS-02 | `src/Chart.Domain/ChartDocumentV3.cs:1`; `src/Chart.Domain/ChartV3Codec.cs:1` |
| domain-v3 | VSTO-01 | `src/Chart.Domain/ChartDocumentV3.cs:1`; `src/Chart.Domain/ChartV3Codec.cs:1` |
| domain-v3 | chart-calculations | `src/Chart.Domain/ChartDocumentV3.cs:1`; `src/Chart.Domain/ChartV3Codec.cs:1` |
| domain-v3 | domain-v2 | `src/Chart.Domain/ChartDocumentV3.cs:1`; `src/Chart.Domain/ChartV3Codec.cs:1` |
| domain-v3 | strict-json | `src/Chart.Domain/ChartV3Codec.cs:1` |
| library-config | catalog | `src/Library.Core/LibraryConfiguration.cs:1` |
| library-config | catalog-query | `src/Library.Core/LibraryConfiguration.cs:1` |
| library-config | path-policy | `src/Library.Core/LibraryConfiguration.cs:1` |
| library-config | preview-cache | `src/Library.Core/LibraryConfiguration.cs:1` |
| library-config | source-extraction | `src/Library.Core/LibraryConfiguration.cs:1` |
| library-refresh | catalog | `src/Library.Core/LibraryRefreshMonitor.cs:1` |
| library-refresh | catalog-query | `src/Library.Core/LibraryRefreshMonitor.cs:1` |
| library-refresh | chart-calculations | `src/Library.Core/LibraryRefreshMonitor.cs:1` |
| library-refresh | preview-cache | `src/Library.Core/LibraryRefreshMonitor.cs:1` |
| native-renderer | OPS-02 | `src/Chart.Renderers/NativeRenderingCore.cs:1` |
| native-renderer | OPS-06 | `src/Chart.Renderers/PowerPointSceneSurface.cs:1` |
| native-renderer | VSTO-01 | `src/Chart.Renderers/NativeRenderingCore.cs:1` |
| native-renderer | chart-calculations | `src/Chart.Renderers/NativeRenderingCore.cs:1`; `src/Chart.Renderers/PowerPointSceneSurface.cs:1` |
| native-renderer | domain-v2 | `src/Chart.Renderers/NativeRenderingCore.cs:1` |
| native-renderer | scene-contract | `src/Chart.Renderers/NativeRenderingCore.cs:1`; `src/Chart.Renderers/PowerPointSceneSurface.cs:1` |
| native-renderer | ts-scene | `src/Chart.Renderers/NativeRenderingCore.cs:1`; `src/Chart.Renderers/PowerPointSceneSurface.cs:1` |
| path-policy | chart-calculations | `src/Library.Core/PathPolicy.cs:1` |
| piidx-schema | sqlite | `src/Library.Core/PiidxSchema.cs:1` |
| preview-cache | catalog | `src/Library.Core/PiidxPreviewCache.cs:1` |
| preview-cache | catalog-query | `src/Library.Core/PiidxPreviewCache.cs:1` |
| scene-contract | chart-calculations | `src/Chart.Renderers/SceneCodec.cs:1`; `src/Chart.Renderers/SceneContracts.cs:1` |
| scene-contract | strict-json | `src/Chart.Renderers/StrictSceneJson.cs:1` |
| scene-contract | svg-preview | `src/Chart.Renderers/SceneCodec.cs:1` |
| scene-contract | ts-scene | `src/Chart.Renderers/SceneCodec.cs:1`; `src/Chart.Renderers/SceneContracts.cs:1` |
| source-extraction | VSTO-06 | `src/Library.Core/PresentationSource.cs:1` |
| source-extraction | chart-calculations | `src/Library.Core/PresentationSource.cs:1` |
| source-extraction | path-policy | `src/Library.Core/PresentationSource.cs:1` |
| source-extraction | svg-preview | `src/Library.Core/PresentationSource.cs:1` |
| sqlite | OPS-05 | `src/Library.Core/NativeSqlite.cs:1` |
| sqlite | OPS-06 | `src/Library.Core/NativeSqlite.cs:1` |
| sqlite | chart-calculations | `src/Library.Core/NativeSqlite.cs:1` |
| sqlite | svg-preview | `src/Library.Core/NativeSqlite.cs:1` |
| strict-json | svg-preview | `src/Shared.Contracts/BoundedJson.cs:1` |
| svg-preview | chart-calculations | `packages/chart-engine/src/svg-preview.ts:1` |
| ts-scene | OPS-02 | `packages/chart-engine/src/scene-builder.ts:1` |
| ts-scene | OPS-05 | `packages/chart-engine/src/scene-contract.ts:1`; `packages/chart-engine/src/strict-scene-json.ts:1` |
| ts-scene | OPS-06 | `packages/chart-engine/src/scene-contract.ts:1`; `packages/chart-engine/src/strict-scene-json.ts:1` |
| ts-scene | VSTO-01 | `packages/chart-engine/src/scene-builder.ts:1` |
| ts-scene | VSTO-04 | `packages/chart-engine/src/scene-contract.ts:1` |
| ts-scene | chart-calculations | `packages/chart-engine/src/scene-builder.ts:1`; `packages/chart-engine/src/scene-contract.ts:1`; `packages/chart-engine/src/strict-scene-json.ts:1`; `packages/chart-engine/tools/render-chart-v2.mjs:1` |
| ts-scene | domain-v2 | `packages/chart-engine/src/scene-builder.ts:1` |
| ts-scene | scene-contract | `packages/chart-engine/src/scene-builder.ts:1`; `packages/chart-engine/src/scene-contract.ts:1` |
| ts-scene | svg-preview | `packages/chart-engine/src/scene-builder.ts:1`; `packages/chart-engine/src/scene-contract.ts:1`; `packages/chart-engine/src/strict-scene-json.ts:1` |
