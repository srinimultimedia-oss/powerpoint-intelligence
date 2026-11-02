# Renderer integration investigation — 22 October 2026

Acceptance remains OPEN. No Office acceptance is inferred from compilation or fake-surface tests.

## Observed activation failure

`scripts/Diagnose-PowerPointActivation.ps1` records only process/security/registration metadata, not presentation contents. Current evidence: `artifacts/activation-diagnostics/activation.json`, `permissions.json`, and `security-log-access.json`.

At 16:02 local time, ProgID resolution succeeded. `Activator.CreateInstance` failed obtaining the PowerPoint COM class factory with UnauthorizedAccessException / 0x80070005, before reading Application.Version, selecting a slide, or invoking the renderer. Both registry views resolve CLSID `{91493441-5A91-11CF-8700-00AA0060263B}` to the installed Office16 POWERPNT executable. Click-to-Run reports x64, build 16.0.20527.20088. Registry key and executable security descriptors are readable; registration readability is not permission to activate COM.

The caller is SRINI\srini in a genuine Windows x64 process. GetTokenInformation reports TokenIsAppContainer=1, TokenHasRestrictions=1, TokenElevationType=3, TokenElevation=0. LOCALAPPDATA is redirected to a sandbox package. These are verified execution restrictions, not evidence that PowerPoint is broken. AppContainer access includes package/capability access checks in addition to the user's permissions ([Microsoft documentation](https://learn.microsoft.com/en-us/windows/win32/secauthz/implementing-an-appcontainer)). An isolation boundary is a supported explanation for the access denial; the exact denying COM access rule is NOT established. A normal approved desktop-token comparison is still needed to establish causality and verify remediation. Code Integrity and Defender event log reads also return access denied in this context. No DCOM grants, policy exclusions, registry permission changes, certificate trust changes, or elevation workarounds were applied.

## Remaining local check

`scene-contracts-x64` compiles the net48 library and explicitly x64 executable successfully, records SHA256 in compiled.json, then fails launching SceneContractTests.exe with ResourceUnavailable / 0xfffffffe. The executable subsequently disappears. The cause is unproven; no security product is blamed without accessible event evidence. AnyCPU/x86 scene tests and explicit x64 compatibility tests execute. Compilation alone does not satisfy the x64 scene execution gate.

## Changes

- Added repeatable, read-only activation diagnosis with isolated stage and token information.
- Added `Intelligence > Create > Elements > Editable shape columns (sample)` with original column icon and compatibility-enabled callback. The command loads validated embedded canonical/scene samples, assigns a fresh chart identity, recomputes the scene hash, validates before mutation, and calls NativeSceneRenderer/PowerPointSceneSurface on the Office UI thread. No Node subprocess or external service is used in PowerPoint.
- Added renderer, canonical DTO, strict parser and Newtonsoft runtime dependencies to the VSTO project. Sample resources are embedded with explicit names.
- Ribbon build marker advanced to v4. Package freshness fingerprint now covers embedded JSON and renderer/domain/shared source dependencies.
- The command is a bounded local sample, not a general production command-bus integration. Axes/legend are deliberately absent and disclosed in its tooltip. Other families remain outside this integration increment.

## Validation boundary

41/41 fake-surface renderer tests and 35/35 canonical domain tests pass. Ribbon definition/icon/callback checks pass with 106 controls and all existing 90 demo bindings. Neither source XML nor callback presence proves installed ribbon appearance. The installed package has not been promoted to this source build. The interactive suite and current full-run results are recorded separately below.

The owned-document native test covers creation, CustomerData persistence, save/close/reopen, native edit detection and restoration, duplicate identity rejection, and rollback. It cannot reach those assertions in this process because activation fails first. Do not label those scenarios passing. Native geometry edits currently diverge from canonical geometry and are reported by ReadVerified; automatic canonical reconciliation is not implemented.

## Required legitimate desktop validation

From an approved normal Windows desktop PowerShell session, close user presentations before the owned-document tests, then run:

```powershell
.\scripts\Diagnose-PowerPointActivation.ps1 -OutputDirectory artifacts/activation-desktop
.\scripts\Test-All.ps1 -InteractiveOffice
```

Compare token flags, activation stage, x64 execution, and any administrator-accessible security events before changing anything. If desktop activation also fails, the AppContainer comparison does not establish the cause; continue investigating the actual Office/COM permission or runtime error. Install the current signed build through the existing visible VSTO installer only after local checks pass. Verify the v4 installed receipt and actual renderer command in PowerPoint, create sample shapes, edit a rectangle, save, close, reopen, and confirm editable native shapes and retained canonical metadata. Test-InstalledRibbon cannot substitute for visual command execution.

## Acceptance

Not complete. No verified activation fix, 20/20 result, installed v4 ribbon or end-to-end edit/save/reopen acceptance is claimed.

## Executed full-suite evidence (16:03–16:07)

Run: `artifacts/validation-20261022-160354-d2db390b/results.json` and per-check logs. Immutable source fingerprint: `A526149A8D80AE6D0ABD1B9F8A3283D51E82168DE6F764E1185A9FD51286E239`.

- Local checks: **19/20 PASS**. Only scene-contracts-x64 failed. Its compiled executable SHA256 was `72BE5A0B44966754BC2D7911861419415A40489490314A420B265F3800AE7F1E`; compilation receipt is retained inside the snapshot artifacts.
- Renderer core: **41/41 PASS**, including rollback, metadata validation, owner-thread and reentrancy checks.
- Canonical domain: **35/35 PASS**. Scene contracts AnyCPU/x86, chart engine, type checking, schemas, calculations, cross-language, S0 unit, demo compile, ribbon and diff checks pass.
- VSTO Compile target: **AnyCPU/x86/x64 PASS**. Explicit x64 compatibility executable: **36/36 PASS**, including process architecture assertion. The machine has target x64 PowerPoint; this does not establish Office integration execution.
- Interactive checks: installed-ribbon, demo-feature-runtime, native-shapes-office, and office-1 all **FAIL at 0x80070005 COM activation**. No renderer mutation or Office round-trip assertions were reached. S0 activation retries stopped after the first denial.

Owner agreed to run the desktop-token comparison. Results are pending and must not be inferred. The feature remains unaccepted.

## x64 assembly/package build

The full x64 Build produced PowerPointIntelligence.dll and copied Newtonsoft.Json.dll. PE inspection confirms Amd64; both canonical/scene resources are embedded with their required logical names (artifacts/activation-diagnostics/renderer-assembly.json). The build then **failed at the VSTO manifest stage** because signed ClickOnce manifests are mandatory. This was an unsigned verification build, not an installation candidate. The existing Install-Local.ps1 signing/visible-installer flow remains required in an approved desktop context. No unsigned package was installed, and a compiled DLL does not count as successful deployable package validation.

The installed-ribbon check was subsequently updated to require v4 and the new renderer command explicitly; the full-suite snapshot above predates that check-only update. It has not passed against an installed v4 package.

## Follow-up: ribbon preparation and stronger native round trip

The actual ribbon sample preparation is now tested with seven assertions for fresh identity, complete node remapping, hash recomputation, exact retained canonical/geometry/style/diagnostic values, independent snapshots and missing-host rejection. Explicit x64 execution passed 41 core + 7 preparation tests, with architecture assertion, execution.log and receipt hashes under artifacts/native-renderer-x64. The interactive test now invokes the ribbon's shared insertion implementation and saves/reopens a manually edited native bar a second time while verifying that canonical metadata is unchanged and divergence remains detectable. These new COM assertions remain unexecuted because activation still returns 0x80070005.

The owner subsequently prioritized investigation of the missing scene-contracts-x64 executable and a reported antivirus alert. See scene-antivirus-investigation-2026-10-22.md for preserved hashes, read-only collector, control results and remaining scanner evidence. No replacement of the x64 check, policy exclusion or restore was performed. The next full-suite run must use current source; the prior 19/20 run is evidence for its recorded immutable snapshot only.

## Protected retest and direct UI evidence, 22 October 2026 evening

McAfee quarantine is now confirmed by exact hash-matched detection records; the x64 missing-file symptom is explained. After the owner re-enabled McAfee, the same x64 executable passed without rebuilding. A fresh full suite (artifacts/validation-20261022-204800-08191651, source EA3A7E46C070A98E698EBBD39BEC3CD3E3FCAEB39800424ECC3990789ECC2239) passed **20/20 local checks**, including x64 scene execution and 41 renderer core + seven ribbon preparation tests. The four Office checks still fail at COM class-factory activation, 0x80070005. Details and limitations are in scene-antivirus-investigation-2026-10-22.md; no lasting antivirus classification fix is inferred.

The supported computer-use desktop API successfully launched PowerPoint and showed the installed Intelligence ribbon with original icons. Its Create > Elements menu exposes only Managed sample column, **not Editable shape columns (sample)**. The installed receipt is version 1.0.0.3 from 15:06; current renderer-source loading is not accepted. Desktop UI operation succeeds independently of this AppContainer's denied COM client activation; no COM permissions were changed.

An owned legacy sample was inserted through that installed ribbon. Its embedded workbook's Q2 Actual value was edited from 20 to 22.5, saved locally, closed, reopened in PowerPoint and selected again as a native chart (Chart Design/Format context and series handles visible). Closed-package checks confirm the original GUID, revision 1, edited value in canonical XML, native chart cache and embedded XLSX, clustered-column type and two series. Evidence: artifacts/activation-diagnostics/legacy-ui-persistence.json, artifacts/ui-evidence/legacy-native-chart-reopened.png and artifacts/ui-legacy-native-edit-20261022.pptx. The owned test deck and empty PowerPoint instance were closed normally afterward. This proves one round trip of the **previously installed S0 chart**, not the new grouped-shape renderer, duplicate/corrupt-metadata scenarios or an Office product matrix.

Current acceptance: local 20/20 gate verified for the recorded protected run; x64 renderer/core gate verified; legacy desktop UI round trip verified narrowly. New renderer signed deployment, installed command execution and native-group edit/save/reopen remain unverified. The complete renderer feature is NOT accepted.

## Final boundary for this increment

Current-source fingerprint audit confirms the worktree's apps/src/scripts/tests/packages match the protected 20/20 snapshot (artifacts/activation-diagnostics/current-source-audit.json); subsequent changes were documentation/evidence only. The current x64 executable remains present. No new matching McAfee quarantine was recorded through the last log read. The generic McAfee classification has not been adjudicated by the vendor.

After closing the owned Office test normally, Install-Local.ps1 was retried for deployment 1.0.0.4. It failed **before build/installation**, at New-SelfSignedCertificate: `CertEnroll::CX509Enrollment::_CreateRequest: Access is denied. 0x80070005`. Log: artifacts/install-20261022-211228.log. No new certificate/package was installed and the existing installed receipt remains 1.0.0.3. This is separate from the repaired x64 missing-file test. A normal approved desktop install/test execution was requested because this restricted token cannot create the signing request. No security-control change or new permission grant is proposed.
