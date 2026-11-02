# PIIDX engine architecture — core implemented, UI migration pending

Repository: C:/Srini/powerpoint-intelligence/powerpoint-intelligence. Branch: feature/piidx-library-engine. Original audit baseline: 78b9f87ef2218d6d19d7e8049d0f223281e9ed76; initial owner merged code was largely untracked and is preserved separately from migration commits.

The owner explicitly selects SQLite for the local PIIDX document index. Q-006 remains open for enterprise server database/HA/operational ownership; this embedded-index implementation does not resolve those infrastructure questions.

## Implemented components

- NativeSqlite.cs: System32 OS SQLite P/Invoke with SafeHandle, typed binding, single-statement checks, native query interruption, automatic-rollback-aware cleanup and bounded streaming reads.
- PiidxSchema.cs: initial versioned schema migration; exact schema/header/version and optional deep SQLite/foreign-key verification.
- PresentationSource.cs: bounded original PPTX/POTX discovery/extraction, raw native IDs distinct from positions, source hashing, typed objects/tags, strict/transitional XML and cancellation. See source extraction scope/limits.
- LibraryModels.cs: typed catalog registrations, progress/results, references and search requests/pages.
- PiidxCatalog.cs: persistent multi-library registration, per-library incremental/forced rebuild, stable catalog identities, source revalidation before commit, sanitized jobs/errors, logical content/source fingerprints and stale-reference protection.
- PiidxSearch.cs: FTS5 safe lexical query construction, ranked/bounded paging and filters, coherent read snapshots and source verification.
- PiidxPreviewCache.cs: persistent bounded derivative PNG cache, corruption/dimension/checksum checks, source-version invalidation and eviction. Real Office image production is an integration responsibility, not simulated by this cache.

## Ownership and transactions

Each operation owns its own SQLite connection. All catalog mutations acquire an exclusive OS writer-lock file handle, not merely a timestamp/marker. Only after acquiring that handle may a refresh classify prior running jobs as interrupted. Tests prove that a live held handle prevents recovery from stale job rows alone. The persistent lock filename is not evidence that a process is live or stopped.

Source discovery is bounded; presentations are parsed one at a time. Refresh uses BEGIN IMMEDIATE, records running jobs outside the payload transaction, commits content/digests/completion together, and rolls back the previous library snapshot on cancellation/error/source races. Missing sources are retained as missing, not fabricated or rebound by guesswork. Unique exact-hash renames retain IDs; ambiguous renames do not. Read snapshots remain available during refresh.

Derivative previews have individual PNG/hash/version checks and are intentionally outside the canonical payload digest to permit inexpensive regeneration. They are not source data and their corruption must not erase the source catalog. SQLite cache pages are reused after eviction; payload limits are not an immediate file-shrink promise.

## Verification

The six-group core suite passes 123 x64 assertions. It persists the entire copied 62-package corpus (497 slides, 13049 objects, 2078 tags), tests concurrent search/refresh, cancellation/rollback, failed/corrupt source/index handling, stale selection, missing/restored sources, multi-library removal, PNG cache persistence, a 1001-slide input and a >275-character Unicode source path. One run measured initial indexing 8.7 s and unchanged refresh 0.24 s; this is machine-specific evidence, not an enterprise SLA.

The same sources compile with the .NET Framework compiler for AnyCPU/x86/x64. Executed tests use the available x64 PowerShell runtime. Actual Framework-hosted/x86 and Office 2016/2019/365 runtime matrices remain acceptance gates.

## Remaining integration, not hidden scope reduction

Connect the core to the existing gallery/ribbon with registration/removal, SQL search/paging, automatic refresh/debounce, enterprise configuration and migration/rollback of the legacy remembered folder. Keep preview/insertion COM serialized on Office UI STA, preserve current presentation and native source formatting, and verify save/reopen. Remove the installed-vendor source fallback and obsolete production XML cache only after replacement is connected. Update project/source fingerprints, build/installer scripts, dependency isolation tests, release documentation and final repository coverage. Future semantic search must use a configured provider boundary; no live AI behavior is claimed here.