# Interoperability, threading and process boundaries

## Object model/architecture

The compiled VSTO assembly targets .NET Framework4.8, C#7.3, AnyCPU/x86/x64 configurations and Office/PowerPoint15PIAs with EmbedInteropTypes (`apps/powerpoint-vsto/PowerPointIntelligence.csproj:9`, `:18`, `:30`). Office runtime is external; VSTO tool/runtime references are not bundled as if self-contained .NET runtime. Process bitness must match Office, not Windows OS; explicit profile checks gate Office16API/build/bitness (`apps/powerpoint-vsto/OfficeCompatibility.cs:61`). Compatibility unit/build checks do not certify every installed Office edition/build.

## STA and UI marshalling

PowerPointHost captures creator managed thread ID and requires STA on access (`apps/powerpoint-vsto/PowerPointHost.cs:58`). DemoAutomation InsertChartJson enforces the same principle (`apps/powerpoint-vsto/DemoApi.cs:35`). WinForms timers execute on the UI message thread; context observation is immutable/volatile so external COM tests do not inspect Office/WinForms from the wrong thread. LocalLibraryWorkspace performs catalog I/O on background tasks and returns via UiAwaiter before native Office actions (`apps/powerpoint-vsto/LocalLibraryWorkspace.cs:17`). Disposal fallback must be reviewed carefully; background continuation does not authorize COM access after the form closes.

## Native handles and COM value shapes

user32 interop supplies window bounds, foreground owner and HWND->process identity. IUnknown-based document identity is used only as an opaque live key; acquired raw pointer is released in finally (`apps/powerpoint-vsto/ContextSelectionBridge.cs:30`). Native series can return typed one-dimensional SAFEARRAYs with nonzero lower bound or a scalar; PowerPointHost.SeriesVector normalizes Array.Cast<object> and rejects multidimensional arrays (`apps/powerpoint-vsto/PowerPointHost.cs:335`). Excel ranges are rank2Value2 arrays and must honor actual lower bounds.

## Embedded workbook sequencing

Activate -> acquire workbook -> write/verify/calculate -> bind native source -> wait actual chart acknowledgement -> close -> verify retained chart values. Local link writes retry known workbook-busy acquisition errors; Close uses en-US late-bound optional-argument invocation to avoid locale/COM dispatch ambiguity (`apps/powerpoint-vsto/DataSources.cs:199`). Only RPC_E_CALL_REJECTED0x80010001 is replayed for mutation close; specific disconnected close codes are tolerated only when native committed values independently verify (`:206`). Arbitrary RPC failures must not be blindly retried.

## Ownership and cleanup

Owned hidden source presentations are closed by SourceTemplateSession.Dispose; already-user-open sources are never adopted. Task panes/forms/timers/images/catalog monitors are disposed at their boundaries. Source Excel file reads close read-only books and release application RCW; subordinate RCW cleanup is not uniformly explicit. The app does not call global FinalReleaseComObject on Office-owned application objects. Do not claim a uniform complete COM lifetime strategy where the source uses mixed dynamic/typed proxies.

## Coordination and trust

OfficeMutationLease uses a process-specific Local named mutex. Tests discover actual PowerPoint PID from HWND and coordinate with installed automatic writers; unrelated add-ins are not constrained by it. Publisher/VSTO installation trust and Windows activation policy are external boundaries. A source/build pass is distinct from successfully activating PowerPoint under an AppContainer/desktop token. Existing diagnostics retain HRESULT-only operation descriptions to avoid workbook/metadata disclosure. No COM DCOM permission weakening, antivirus exclusion or trust-store modification is prescribed by this design.

## Tests and unresolved native coverage

See native test source for scalar/SAFEARRAY normalization, worksheet lower bounds, stale selection, grid/indicator/connectors, copied XML ownership and scene primitive round trips. This documentation leaf runs no Office automation. Per-monitor DPI, every chart family, focus and security-token behavior remain explicit environment-dependent verification items.
