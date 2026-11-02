# Office configuration evidence — 22 October 2026

Owner request: configurable Office 2016 ("Office 16"), Office 2019 ("Office 19") and Microsoft 365 versions. Product-family scope is recorded as AD-010, with exact builds/channels/signing/licensing still unresolved.

| Check | Result |
|---|---|
| Compatibility helper/tests compilation | Passed locally with Roslyn 4.14.0 for AnyCPU, x86 and x64 |
| Compatibility test execution | Passed 36 tests per architecture in Windows CI: AnyCPU, x86 and x64 (108 executions total); requested x86/x64 process architecture was asserted |
| S0 adapter/integration harness compilation using discovered embedded PIAs | Passed locally |
| Whitespace | `git diff --check` passed |
| Full VSTO compilation/load | Not verified; no Office development targets/targeting pack available locally |
| Office 2016 native behavior | Not verified on either architecture |
| Office 2019 native behavior | Not verified on either architecture |
| Microsoft 365 native behavior | Not verified on either architecture |

Tests cover all profile/architecture combinations, explicit architecture restrictions, minimum-build comparisons, missing/invalid/malformed settings, non-16 APIs, missing/full/inconsistent executable build information and the lack of edition detection from the shared version number. CI runs pure .NET test executables in AnyCPU, x86 and x64 configurations; no Office automation is performed on a server.

Executed source: `dce9412e52c1d0a203b7f123f1c20ccd81a9660e`. [Windows compatibility run 39726879521](https://github.com/srinimultimedia-oss/powerpoint-intelligence/actions/runs/39726879521), jobs 119652701027 (AnyCPU), 119652701367 (x86), 119652701346 (x64). Each job log reports `Office compatibility tests passed: 36`. Runners: Windows Server 2025, windows-2025-vs2026 images 20261018.275.1 / 20261007.264.1. S0 metadata regression run 39726879423 and chart-core run 39726879440 succeeded. Evidence-only follow-up commits do not broaden what those tests prove.

The changes configure compatibility policy, startup/ribbon gating, sidecar settings and build architecture. They do not certify deployment or native chart/persistence behavior. Complete the per-product/build matrix in [the validation guide](../implementation/office-compatibility.md) before claiming tested compatibility.
