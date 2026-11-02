# Corpus basename rejection on .NET Framework — 2 November 2026

## Confirmed cause

After the Unicode bootstrap encoding correction, desktop Windows PowerShell reported `ArgumentException: Illegal characters in path` from SceneArtifactPath.Resolve. The first guard called Path.IsPathRooted before checking raw invalid filename/control characters. .NET Framework 4.8 throws on the newline fixture at that call; modern .NET did not expose the same early exception. The test expected the domain InvalidDataException and correctly failed instead of swallowing arbitrary exceptions.

A regression added to the existing stable SceneContractTests.exe runner reproduced the exact ArgumentException on genuine Windows x64 .NET Framework before the fix. Evidence: `artifacts/corpus-validation-net48-before.log`.

## Fix

Removed the redundant IsPathRooted call from the basename guard. Raw filename/control checks and explicit rejection of every slash, backslash and colon already reject all rooted, UNC, device and drive-relative names. No platform path parsing occurs until these name checks complete. Length, Windows device-name checks, output containment and duplicate-name checks remain intact. Tests were not relaxed to accept arbitrary ArgumentException failures.

Added newline, NUL, quote, pipe, less-than and greater-than cases to the existing native scene runner; the bootstrap adds the five previously missing fixtures using ASCII-safe numeric construction where needed. No application code, deployment DLL, trust setting or antivirus policy changed. Stable scene test executable names and source-review receipts were retained. No missing/quarantined binary was restored.

## Executed verification

- Available PowerShell security suite: **38 corpus-path**, **30 path-authority**, **31 diagnostic-path** assertions and **22 Python security tests** pass. Evidence: `artifacts/corpus-validation-security-after.log`. Junction fixture creation remains explicitly skipped in this restricted session.
- Genuine Windows .NET Framework scene runner: **154 assertions + 73 cross-language round trips on each of AnyCPU, x86 and x64**, plus actual malicious corpus traversal rejection. Evidence: `artifacts/corpus-validation-net48-{AnyCPU,x86,x64}-after.log`.
- The six new raw-character tests require InvalidDataException; the pre-fix x64 failure and post-fix passes verify this is the actual runtime incompatibility, not only a simulated source check.
- Direct Windows PowerShell 5.1 script-file execution remains unavailable in this agent token under the previously recorded AuthorizationManager restriction. This change is verified through the actual .NET Framework executables, but complete desktop PowerShell wrapper execution remains for the user. No execution-policy bypass was used.

No add-in reinstall is needed for these test/helper changes. Run in a fresh process so Add-Type compiles current source:

```powershell
powershell.exe -NoProfile -File .\scripts\Test-SecurityRemediation.ps1
```
