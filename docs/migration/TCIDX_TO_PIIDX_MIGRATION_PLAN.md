# TCIDX to PIIDX migration plan

1. Preserve baseline inventory and owner changes; audit source, projects, UI, configuration, scripts, binary asset/package metadata and historical evidence.
2. Independently inspect the supplied research sample without executing vendor code. Keep research out of the runtime build.
3. Specify SQLite PIIDX schema, migrations and compatibility, validate an FTS5-capable native provider, then implement the reader/writer/search interfaces with bounded resource use.
4. Implement source-file hashing and stable native slide identity, metadata/object extraction, transactional incremental indexing, cancellation/progress, errors/jobs and recovery.
5. Integrate multiple-library registration/configuration, automatic refresh and indexed search into the existing gallery/ribbon. Reuse serialized UI-thread PowerPoint preview/insertion services, add persistent preview caching and freshness guards.
6. Migrate the old remembered folder/XML registration through original presentations; preserve rollback metadata. Do not import TCIDX content or mutate sources.
7. Remove vendor installed-directory fallback and vendor-dependent test fixture selection after replacement is connected. Add isolated dependency checks; preserve historical research and unrelated features.
8. Run format/search/corruption/refresh/cancellation/large/unicode tests, compile all architectures, and perform desktop PowerPoint editing/insertion/save/reopen regression tests where authorized execution is available. Mark environment-denied cases explicitly.
9. Re-scan every file, publish required specification/API/architecture/coverage/elimination/final reports and acceptance matrix, and commit completed changes on the appropriate branch without silently committing unrelated owner modifications.

Current phase: audited baseline and provider feasibility. Full migration is not accepted. Live Office integration, x86 execution and installer verification remain evidence gates.