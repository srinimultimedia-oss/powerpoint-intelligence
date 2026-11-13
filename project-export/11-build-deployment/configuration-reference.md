# Configuration reference

## Add-in configuration

AppSettings source: `apps/powerpoint-vsto/App.config:1`.

| Key | Default and semantics | Consumer |
|---|---|---|
| OfficeProfile | auto; office2016/office2019/microsoft365 allowed | OfficeSettings.Parse allowlist |
| OfficeArchitecture | any; x86/x64 process check | OfficeCompatibility.Evaluate |
| OfficeMinimumVersion |16.0; optional four-part build |16.0 only; executable build may be required |
| TemplateLibraryFolder | Development source local repository path; builder rewrites relative templates | Resolve beside deployed DLL; explicit missing root reports unavailable |
| LibraryIndexPath | Empty uses local per-user default | LibraryConfiguration |
| LibrarySourcesConfig | Empty uses per-user source registration default | LibraryConfiguration |
| LibraryRefreshSeconds |60 | Refresh monitor/configuration |
| LibraryWatchRootLimit |32 | Watcher root cap |

OfficeSettings.Load explicitly maps add-in .dll.config, never edits POWERPNT.EXE.config, rejects unknown compatibility settings and permits library keys for independent resolution. `apps/powerpoint-vsto/OfficeCompatibility.cs:33`.

## Service profiles

LocalAppData/PowerPointIntelligence/service-connections.v1.xml: three nonsecret profiles, version1/strict namespace/bounds/no DTD. Empty profiles are valid not-configured values. Not a credential storage format. `apps/powerpoint-vsto/ServiceConnections.cs:35`.

## Build and test flags

Project: PpiForcePayloadCompile, PpiEasySetupBuild, PpiSkipEasySetup, SignManifests, ManifestCertificateThumbprint, ApplicationVersion, Platform, Configuration. Test-All: InteractiveOffice, OfficeRuns=3, LocalTimeoutSeconds=300, OfficeTimeoutSeconds=3600. Focused demo tests expose AgendaOnly/ElementsOnly/PropertiesOnly/LinkedChartDiagnostic/PauseOwnAddIn/MappedExcelOnly; invalid combinations fail before Office use. `scripts/Test-All.ps1:1` `scripts/Test-DemoFeatures.ps1:1`.

Branding does not rename persisted XML namespaces, tags, COM/assembly/storage/installer identities. `docs/DECISIONS.md:4`. No Azure secret/token environment configuration is implemented.
