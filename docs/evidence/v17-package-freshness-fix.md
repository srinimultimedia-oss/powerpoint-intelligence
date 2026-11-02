# v17 rollback package freshness fix

The owner installation log artifacts/install-20261027-004159.log reports a stale package after building deployment 1.0.0.24. Direct inspection confirmed that the deployed x64 DLL and cached intermediate DLL still contained the v19 marker while DemoRibbon.cs contained v17. Restoring snapshot files preserved older timestamps, causing incremental compilation to reuse newer outputs. The stale-package guard correctly rejected this DLL; it was not a marker-search false positive.

The earlier Compile-target success reports only established successful target execution, not a fresh v17 DLL. This evidence supersedes that interpretation.

## Fix

Install-Local.ps1 now passes PpiForcePayloadCompile=true. The project ForcePayloadCompilation target runs before CoreCompile and invalidates only the intermediate assembly. MSBuild then compiles the restored source and performs the existing signed build/manifest/package verification. The deployment revision still increases, allowing ClickOnce to update to restored functionality without lowering its deployment version. No marker/hash/signature/trust guard was relaxed.

A trial full Rebuild verification hit the restricted agent's registry-access boundary inside VSTO Clean. The final fix does not use Rebuild/Clean, which include VSTO registration/security cleanup targets. The successful verification used only Compile with forced intermediate invalidation; no Office activation or registration operation was requested by that verification.

## Executed evidence

- The actual forced x64 intermediate DLL contains PPI-DEMO-RIBBON-20261026-v17 and no v19 marker: artifacts/v17-forced-compile.log.
- Test-PayloadRecompile.ps1 reproduces reuse with an intermediate DLL dated newer than source, then proves forced compilation replaces that output and embeds the current source marker. PASS: artifacts/rollback-compile-regression.log and artifacts/rollback-compile-x64.log. No new executable was created.
- All 8 existing package-marker byte-alignment/rejection cases pass.
- git diff --check passes.

The deployed bin package remains unaccepted until a normal desktop signed build regenerates it. No agent installation or signed-package success is claimed. Run .\scripts\Install-Local.ps1 without SkipBuild in normal Windows PowerShell with PowerPoint closed. The existing package guard must pass before the visible VSTO publisher prompt.
