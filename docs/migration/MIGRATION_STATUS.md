# PIIDX migration progress — full acceptance pending

Working root: C:/Srini/powerpoint-intelligence/powerpoint-intelligence. Branch: feature/piidx-library-engine. Preserve the complete owner goal; this checkpoint is not a smaller definition of success.

## Verified

The immutable initial inventory accounts for all 15 files at baseline commit 78b9f87ef2218d6d19d7e8049d0f223281e9ed76 and merged untracked/ignored/admin files. Historical research is isolated. No production TCIDX reader was found; the old optional installed-vendor path and gallery XML cache still require removal/integration changes.

Core implemented: strict versioned SQLite schema, original presentation extraction, persistent multi-library indexing, incremental/forced rebuild, stable IDs, source fingerprint guards, sanitized jobs/errors, OS-lock-based interrupted-job recovery, safe ranked FTS5 search, source selection guards and persistent validated PNG caching.

scripts/Test-Piidx.ps1 passes 123 x64 assertions: 28 foundation, 24 sources, 46 catalog, 14 previews, 4 edge cases, 7 corpus. Latest evidence: artifacts/piidx/suite-5fdea41f329b4c8c854231032362bedd/results.json. The corpus has 62 packages/497 slides/13049 objects/2078 tags. The suite also tests 1001-slide and long Unicode path inputs. Native Office rendering/insertion was not executed. Architecture compilation passes AnyCPU/x86/x64; x86 execution remains NOT EXECUTED — ENVIRONMENT UNAVAILABLE after Windows AuthorizationManager rejection, without bypassing policy.

## Required next work

1. Integrate the core into the existing VSTO gallery/ribbon, real native preview production/cache consumption, editable insertion and UI registration/removal/search/navigation.
2. Implement automatic refresh and enterprise configuration, migrate legacy remembered folders through original sources, preserve rollback information, remove the vendor-path fallback and obsolete XML cache usage.
3. Expand integration/edge coverage, native full-text/visual preservation and every relevant source object type; current package fixtures are extraction tests, not native round-trip acceptance.
4. Wire project/build/source fingerprints, package dependencies/installation checks, dependency isolation tests and CI. Preserve scene quarantine holds and stable runner policies.
5. Run all available regressions plus real Windows/Office insert/edit/save/reopen checks. AppContainer COM activation denial is a separate baseline; normal desktop evidence is required where this token cannot execute.
6. Finish schema/architecture/API/deployment documentation, final tracked/new-file accounting, dependency elimination and implementation reports, the acceptance matrix and final migration commits.

The production VSTO UI is now linked to the new core in source, with integrated gallery tests and three architecture compilation checks passing. See GALLERY_INTEGRATION_CHECKPOINT.md for the current evidence and remaining gates. No claim is made that the full migration or any Office acceptance gate is complete. Owner source, assets and historical acceptance evidence remain preserved.
A sample native-source index and evidence are in artifacts/piidx/sample-native-ffc4208f95e8412b9f26ecaf27d77f99. The input is a previously owner-generated Office test deck; indexing it did not execute Office again. Test-All.ps1 now invokes the PIIDX suite against captured source while reusing the stable test DLL at RunnerRoot; unrelated existing checks remain intact.

29 October recovery update: owner-approved x64 legacy source review is complete; the runner passes and reuses the same path. Latest local suite is **23/23 PASS** (artifacts/validation-20261029-160704-ee2c43de). See docs/implementation/x64-reviewed-recovery-20261029.md. This supersedes the earlier x64 hold; it does not establish current native Office acceptance.
