# Office configuration evidence — 22 October 2026

Owner request: configurable Office 2016 ("Office 16"), Office 2019 ("Office 19") and Microsoft 365 versions. Product-family scope is recorded as AD-010, with exact builds/channels/signing/licensing still unresolved.

| Check | Result |
|---|---|
| Compatibility helper/tests compilation | Passed locally with Roslyn 4.14.0 for AnyCPU, x86 and x64 |
| Compatibility test execution | Pending Windows CI; no local execution claimed |
| S0 adapter/integration harness compilation using discovered embedded PIAs | Passed locally |
| Whitespace | `git diff --check` passed |
| Full VSTO compilation/load | Not verified; no Office development targets/targeting pack available locally |
| Office 2016 native behavior | Not verified on either architecture |
| Office 2019 native behavior | Not verified on either architecture |
| Microsoft 365 native behavior | Not verified on either architecture |

Tests cover all profile/architecture combinations, explicit architecture restrictions, minimum-build comparisons, missing/invalid/malformed settings, non-16 APIs, missing/full/inconsistent executable build information and the lack of edition detection from the shared version number. CI runs pure .NET test executables in AnyCPU, x86 and x64 configurations; no Office automation is performed on a server.

The changes configure compatibility policy, startup/ribbon gating, sidecar settings and build architecture. They do not certify deployment or native chart/persistence behavior. Complete the per-product/build matrix in [the validation guide](../implementation/office-compatibility.md) before claiming tested compatibility.
