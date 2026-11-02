# Workspace cleanup and runtime files

The installed x64 add-in's deployment payload is in `apps/powerpoint-vsto/bin/x64/Release`. `artifacts/runtime-file-inventory.csv` lists the six runtime/deployment files and exact SHA256 hashes for review. This is an inventory, not an antivirus allowlist or safety verdict. No antivirus settings were changed.

The PowerPoint add-in does not need `SceneContractTests.exe`, `TemplateLibraryTests.exe`, or other test runners to load or run. Those are development validation tools. McAfee has detected scene-test binaries; keep them out of a runtime exclusion decision until the exact detections receive vendor/security review. An exclusion for the whole repository, artifacts tree, Downloads or PowerPoint process would cover much more than the deployment payload and is not recommended here.

Cleanup removed duplicate dependency caches (`node_modules`) inside older validation snapshots, unused downloaded compiler/PIA caches and intermediate VSTO `obj` files. The latest two validation snapshots were retained in full. Source, documentation, tests, current project dependencies, installed Release files, installation receipts, corrected presentations, previews, original validation logs/source snapshots and antivirus evidence were preserved. Older snapshots may need dependency restoration to rerun because their generated dependency cache was removed; logs and source fingerprints remain.

Use `scripts/Clean-Workspace.ps1` to preview future cleanup. Use `scripts/Clean-Workspace.ps1 -Apply` to apply that same bounded cleanup. It keeps the latest two snapshot dependency trees by default and refuses targets outside this repository or directories containing reparse points. Cleanup reports record paths, counts and reclaimed sizes under `artifacts/cleanup-*.json`. It does not compile, execute, restore or relocate test runners, nor change exclusions or quarantine.

## Stable scene runner reuse

`Test-All.ps1` now uses the original repository's stable `artifacts/scene-contracts-AnyCPU`, `scene-contracts-x86` and `scene-contracts-x64` runner caches. No additional SceneContractTests.exe is compiled in every validation snapshot. One executable per architecture retains the architecture test matrix. Fixtures/results stay isolated under each snapshot's `artifacts/scene-results-<architecture>`.

Source, compiler, framework and dependency hashes form a build key. Reuse verifies every cached binary against its receipt without changing its timestamp. Only an empty cache permits initial deterministic compilation. Source/toolchain changes, legacy receipts, altered binaries and missing files stop without overwriting/restoring anything. A failed launch/disappeared file also stops for evidence review; ordinary executed test failures can rerun an intact matching binary. No option bypasses provenance or scanner gates.

Ten inert-file cache policy tests pass and now run inside ribbon-definition regression, which also passes. The actual existing x64 file was checked with CompileOnly: it was retained byte-for-byte, and neither compiled nor executed, because its legacy receipt has no build key. Its scanner history remains unresolved; this change does not approve that binary or adopt an unverified file. A reviewed provenance transition/rebuild remains required for this legacy cache. No antivirus settings were changed.

## Stable template-library runner reuse

`TemplateLibraryTests.exe` now follows the same shared build-key/hash cache policy as the scene runner. The default x64 path is `artifacts/template-library-x64/TemplateLibraryTests.exe`. Test-All passes the original repository as RunnerRoot for both local and interactive template checks, so both use that same executable instead of compiling one in each source snapshot. Result previews remain per-run; no extra executable is copied with them.

The cache key covers the template catalog, library adapter, test source, architecture, compiler and referenced .NET 4.8/Office assemblies. Identical inputs reuse a verified binary. Source changes, missing binaries, unknown/legacy receipts and failed launches stop without overwriting or automatically adopting a file. Different architecture checks require separate binaries; the normal template check defaults to x64.

The existing x64 TemplateLibraryTests.exe was preserved byte-for-byte and not executed. It currently lacks a provenance receipt and therefore stops at that guard. Exclusion from antivirus does not supply missing build provenance. Ten inert cache policy tests and script parsing pass. No new TemplateLibraryTests.exe was generated for this change. Runtime reuse of the existing legacy binary remains pending provenance/security review.
