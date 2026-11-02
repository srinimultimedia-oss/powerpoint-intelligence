# Office compatibility configuration

The owner requested Office 2016 ("Office 16"), Office 2019 ("Office 19") and Microsoft 365 on 22 October 2026. These are configurable Windows PowerPoint targets. This does not assert that each product/build has passed the add-in's runtime tests.

## Configure the add-in

Edit `apps/powerpoint-vsto/App.config` before building/publishing. The deployed sidecar is `PowerPointIntelligence.dll.config`, next to the add-in DLL. The add-in loads that exact sidecar rather than POWERPNT.EXE.config. Include it in the signed deployment; after deployment changes follow the approved republish/re-sign process. Restart PowerPoint to apply settings.

```xml
<appSettings>
  <add key="OfficeProfile" value="office2019" />
  <add key="OfficeArchitecture" value="any" />
  <add key="OfficeMinimumVersion" value="16.0" />
</appSettings>
```

| Setting | Accepted values | Behavior |
|---|---|---|
| OfficeProfile | `auto`, `office2016`, `office2019`, `microsoft365` | `auto` uses the common Office 16 API; explicit values record the intended deployment edition. All three families use the same chart/metadata path. |
| OfficeArchitecture | `any`, `x86`, `x64` | Restricts the actual PowerPoint process architecture. Does not infer Office bitness from Windows or convert a 32-bit DLL into a 64-bit DLL. |
| OfficeMinimumVersion | `16.0` or a validated `16.0.build.revision` | Optional build floor. Do not guess a production minimum. A build-specific floor cannot be satisfied using only Application.Version `16.0`. |

Missing sidecar/settings default to `auto`, `any`, `16.0`. An explicitly invalid profile, architecture, minimum or unreadable/malformed config disables chart commands. The **Office Compatibility** ribbon button shows status in the task pane. Existing presentations/charts are not altered by this policy check. No registry or trust settings are written. Successful policy evaluation means only that the configured constraints permit the host; it does not mean that runtime acceptance passed.

Microsoft documents that Office 2016, Office 2019 and Microsoft 365 all share version **16.0**. Office 2021/2024 also share it. Therefore `Application.Version`, folder `Office16`, registry `16.0` and build thresholds cannot reliably label an installed product family. Explicit profile selection records administrator intent; it does **not** verify a license or reject another edition with the same API. This avoids falsely reporting Microsoft 365 as Office 2016. Product/edition identity must be confirmed on the target installation before deployment acceptance.

The runtime check reads PowerPoint.Application.Version and the file version of POWERPNT.EXE under that running application's Path. It checks consistency and uses the executable build for a configured build floor. If the file version cannot be read, the API baseline can still be used with the baseline setting; stricter floors fail safely. This is not a fleet inventory or Microsoft 365 update-channel detector.

## Build architecture and interop

Use the existing .NET Framework 4.8 target with Visual Studio's Office development workload:

```powershell
MSBuild apps/powerpoint-vsto/PowerPointIntelligence.csproj /t:Build /p:Platform=AnyCPU
MSBuild apps/powerpoint-vsto/PowerPointIntelligence.csproj /t:Build /p:Platform=x86
MSBuild apps/powerpoint-vsto/PowerPointIntelligence.csproj /t:Build /p:Platform=x64
```

Outputs are separated under `bin/<Platform>/<Configuration>`. AnyCPU (Prefer32Bit false) allows a managed-only add-in to run within either Office process architecture; architecture-specific builds must match the installed Office. Native dependencies/installer/runtime prerequisites must match that Office architecture as well. No installer/signing choice is made here.

The project's OfficeVersion `15.0` and PIA assembly version `15.0.0.0` are development-tool/interop references, not a claim that the product targets Office 2013. They are retained rather than replacing them with a nonexistent "Office 19" API version. EmbedInteropTypes remains enabled. The used native chart/CustomerData APIs are shared; actual behavior still needs testing on each requested product/build. The adapter test runner discovers development PIAs or accepts explicit `-OfficeInteropPath` and `-PowerPointInteropPath`; it no longer requires one fixed GAC location and also embeds interop types.

## Verification matrix

| Requested product | x86 | x64 | Current runtime evidence |
|---|---|---|---|
| Office 2016 | Configurable | Configurable | NOT VERIFIED |
| Office 2019 | Configurable | Configurable | NOT VERIFIED |
| Microsoft 365 desktop PowerPoint | Configurable | Configurable | NOT VERIFIED |

For **each** selected product/architecture/build/channel: record About PowerPoint details, validate signing/loading, ribbon and status pane, insert native editable chart, edit embedded data, save, exit/restart PowerPoint, reopen and verify metadata/data/formatting, duplicate-slide/cross-presentation copy and safe missing/corrupt metadata behavior. Also test an intentionally mismatched architecture setting and a minimum above the running build: chart commands must disable and status must explain why. Use the existing S0 Windows runbook and record actual results rather than marking the matrix green from unit tests.

`scripts/Test-OfficeCompatibility.ps1 -Platform AnyCPU|x86|x64` runs pure configuration/version/architecture tests without automating Office. Windows CI executes these tests in all three executable architectures. These tests verify policy/configuration code only. Full VSTO build/load, sidecar deployment, ribbon integration and native Office behavior remain unverified until the required development/test environments are available.

## Microsoft references

- [Office product releases share version 16.0](https://learn.microsoft.com/en-us/deployoffice/install-different-office-visio-and-project-versions-on-the-same-computer)
- [Office primary interop assemblies](https://learn.microsoft.com/en-us/visualstudio/vsto/office-primary-interop-assemblies?view=visualstudio)
- [VSTO Runtime lifecycle and Office compatibility](https://learn.microsoft.com/en-us/visualstudio/vsto/visual-studio-tools-for-office-runtime?view=visualstudio)
