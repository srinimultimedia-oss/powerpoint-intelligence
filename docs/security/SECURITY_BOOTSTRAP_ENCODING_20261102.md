# Security bootstrap encoding correction — 2 November 2026

The desktop `Test-SecurityRemediation.ps1` failure, `Unsafe corpus basename accepted`, was traced to a BOM-less UTF-8 test script containing literal Unicode fixtures. Windows PowerShell 5.1 uses the Windows ANSI code page for such files. A Windows-1252 reproduction converts the intended `LPT\u00b9` into `LPT\u00c2\u00b9`; the latter is a different, legal basename and the unchanged C# validator correctly accepts it. This reproduced the test failure without weakening or changing the validator.

The bootstrap is now ASCII, builds its Delta and superscript fixtures using numeric character codes, and rejects future non-ASCII source bytes with an explicit encoding message. Reserved-name tests now cover all three superscripts for both COM and LPT. Failure diagnostics identify fixture character codes, so a future encoding mismatch is visible. No production application/renderer/validator code or cached executable was changed; no add-in reinstall is required for this script-only fix.

Verification:

- Actual bootstrap UTF-8 and Windows-1252 decodings are identical: `artifacts/security-bootstrap-decoder.log`.
- Actual security suite passes in the available PowerShell runtime: 33 corpus-path assertions, 30 path-authority assertions, 31 diagnostic-path assertions and 22 Python security tests. Junction fixture creation is still explicitly skipped by this restricted session. Evidence: `artifacts/security-bootstrap-encoding.log`.
- A normal Windows PowerShell 5.1 `-NoProfile -File` attempt is still blocked in this agent session by `AuthorizationManager check failed`, before script execution. Evidence: `artifacts/security-bootstrap-powershell51.log`. No execution-policy/security bypass was attempted; desktop Windows PowerShell execution remains to be verified by the user.

Use a fresh process because Add-Type assemblies cannot be replaced in the existing PowerShell process:

```powershell
Set-Location 'C:\Srini\powerpoint-intelligence\powerpoint-intelligence'
powershell.exe -NoProfile -File .\scripts\Test-SecurityRemediation.ps1
```
