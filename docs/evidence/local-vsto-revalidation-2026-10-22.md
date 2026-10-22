# Local VSTO revalidation — 22 October 2026

The owner reported that Visual Studio and PowerPoint were installed and requested revalidation, continuing the local installation/testing task.

## Prerequisites now verified

- Visual Studio Community **2026**, version **18.11.12315.225**, registered complete, launchable, no reboot required. Path: `C:\Program Files\Microsoft Visual Studio\18\Community`.
- Office targets exist under `MSBuild/Microsoft/VisualStudio/v18.0/OfficeTools/Microsoft.VisualStudio.Tools.Office.targets`.
- .NET Framework **4.8 targeting pack** exists, including mscorlib.dll. The previous missing-workload/targeting-pack blocker is resolved.
- PowerPoint executable exists, file version **16.0.20527.20088**. The owner identifies it as Office 2019; the executable version alone does not prove edition.

## First full-toolchain findings and fixes

The initial project build found CS0012 (missing Microsoft.Office.Tools reference) and CS0103 (undefined BeginInitialization/EndInitialization helpers). The installed Microsoft PowerPoint VSTO project template was inspected locally to confirm the required references and lifecycle pattern.

Changes: add the Office.Tools and Office.Tools.v4.0.Framework references; do not copy runtime-provided reference assemblies locally; initialize the task-pane collection with BeginInit/EndInit; connect startup handlers in FinishInitialization; and restore the actual PowerPoint HostPackage/ThisAddIn host-item metadata from the installed template.

## Executed checks

| Check | Actual result |
|---|---|
| x64 Release `/t:Compile` using installed Visual Studio MSBuild | **PASS**, exit 0, 0 warnings/errors. Assembly under obj/x64/Release/PowerPointIntelligence.dll. This is compilation only. |
| Normal VSTO Build without a signer | Assembly compiled; deployment stopped because ClickOnce manifest signing is mandatory. |
| Rebuild with temporary local development signer | Failed while Office's clean/unregister task accessed the VSTO solution-metadata store: E_ACCESSDENIED / 0x80070005. |
| Build with temporary local development signer | Certificate import failed/cancelled: MSB3326 and MSB3321. No completed deployment package was generated. |
| C# domain tests with installed compiler | Executable launched but reported FileLoadException; no test pass count was reported. Earlier CI passes are separate evidence. |
| x64 compatibility tests with installed compiler | Executable launch blocked by Windows Application Control; no cases executed in this revalidation. Earlier local/CI passes remain historical evidence. |
| Whitespace | `git diff --check` passed. |
| Actual VSTO install/load, ribbon, pane and native persistence | **NOT VERIFIED**; registration/deployment did not complete. |

Temporary private-key PFX files were deleted in finally blocks after the signing attempts. No manual trusted-root/publisher or application-control policy changes were made. Normal MSBuild certificate/registration operations were attempted and their failures recorded; alternative registry or COM activation workarounds were not used.

## Remaining installation step

Use Visual Studio in the normal Windows desktop session to open `apps/powerpoint-vsto/PowerPointIntelligence.csproj`, select an approved signing certificate (or an explicitly approved local-development test certificate) in Signing, and build x64 Release. VSTO may need to complete its normal per-user certificate/registration prompts. An administrator may need to approve the blocked binaries/COM access under the machine's existing policy. Do not disable policy or mark an unexecuted host test passed.

After a successful signed deployment, verify the COM Add-ins listing and ribbon/pane and run the full S0 native chart round-trip runbook. Until then, the correct status is **prerequisites installed, source compiles, plug-in not installed/validated**.
