# PIIDX gallery integration checkpoint — 2026-10-29

The independent SQLite core is now compiled into the VSTO add-in and linked to its existing Library command and thumbnail gallery. This is an implementation checkpoint, not final migration acceptance.

## Current implementation

- Indexed library/presentation/category navigation, paged lexical search, author/layout/tag/date/aspect filters, scoped duplicate hiding and sorting use the PIIDX catalog.
- Registration, refresh, rebuild, cancellation, watcher debouncing and polling run without background Office calls. Control updates and native preview dispatch explicitly return to the form UI thread.
- Validated PNG derivatives persist in the catalog. Selected references carry canonical source hashes and native slide identities into the original hidden read-only source-session insertion path.
- Native insertion targets the existing presentation. The implementation verifies source freshness and records PIIDX slide/schema tags. Actual installed Office execution of this new path remains unverified.
- The old serialized text cache was removed from production. Its implementation survives only as a regression fixture. Migration reads the old registration root, indexes original files, preserves the old cache and writes a source-hashed migration receipt; cached text is never imported.
- No installed vendor-folder fallback or proprietary index parser is used. Existing descriptive feature-inventory/history strings do not load vendor components.
- The procedural Designed templates dialog remains available from the gallery toolbar.

## Configuration

The deployed add-in configuration accepts `LibraryIndexPath` (empty uses `%LOCALAPPDATA%/PowerPointIntelligence/Library/catalog.piidx`), `LibrarySourcesConfig` (optional separate XML), `LibraryRefreshSeconds` (5–86400, default 60), and `LibraryWatchRootLimit` (0–128, default 32). Relative paths resolve against the deployed configuration directory; paths are literal, not environment-variable expressions. `TemplateLibraryFolder` remains an explicit managed root. Missing configured roots are reported; there is no vendor fallback.

Example external file:

```xml
<Libraries xmlns="urn:powerpoint-intelligence:libraries:v1" version="1">
  <Library path="C:\Company\PresentationAssets" name="Company library" recursive="true" />
</Libraries>
```

Managed registrations cannot be removed through the gallery; edit deployment configuration. SQLite should stay in a local writable location; source folders may be shared. Network deployment, full enterprise policy enforcement and production-scale service behavior are not certified by these tests.

For rollback, retain the prior installed package/configuration and the untouched legacy registration. Do not downgrade a newer unknown PIIDX schema in place. Close the gallery/add-in before replacing its catalog; keep source presentations unchanged. Existing chart metadata and native formatting are not migrated by the index engine.

## Evidence

- `artifacts/piidx/suite-70889325669b475c9516d94758f9af8a/results.json`: all seven groups pass: foundation 28, extraction 24, catalog 46, preview 14, edge 4, workspace/configuration/migration/watchers 20, corpus 7 (143 assertions). Actual Windows x64 execution under runtime 10.0.11; this is not proof of native CLR4/x86 execution.
- `artifacts/piidx/integrated-ribbon-final.log`: eight integrated native WinForms gallery assertions pass, alongside the existing ribbon/element/icon contracts. A synthetic bitmap provider validates dispatch only; its pixels are not Office rendering evidence.
- `artifacts/piidx/integrated-build-AnyCPU.log`, `integrated-build-x86.log`, `integrated-build-x64.log`: production compilation passes for all three targets. Compilation does not certify installation, CLR4 runtime behavior or Office 2016/2019/365 compatibility.
- `scripts/Test-PiidxIsolation.ps1`: production indexing dependency-isolation check passes; the obsolete cache fixture and research are excluded from production compilation.
- `artifacts/piidx/activation-current/activation.json`: Office registration resolves to the installed executable in both registry views, but activation still fails at COM class-factory creation with `0x80070005`; this process has an AppContainer token. No security/antivirus policy was changed.

## Remaining gates

The broader migration remains incomplete. Installed current-package loading, actual PIIDX thumbnail export/insertion/edit/save/reopen, native CLR4/x86 runtime checks, full post-integration regressions, enterprise deployment, comprehensive final file accounting, final acceptance matrix/reports and final integration commits remain open. Historical quarantined-runner and cache-receipt holds remain in force. Existing owner files and source assets are preserved. The current icon and ribbon contract passes do not establish a visible installed ribbon.