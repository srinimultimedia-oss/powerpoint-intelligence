# PIIDX engine architecture — implementation status

Repository: C:/Srini/powerpoint-intelligence/powerpoint-intelligence. Branch: feature/piidx-library-engine. Baseline: 78b9f87ef2218d6d19d7e8049d0f223281e9ed76.

The owner goal explicitly selects SQLite for the local PIIDX document index. Q-006 in DECISIONS remains open for enterprise service/database infrastructure; this embedded-index choice does not resolve server hosting/backup/HA ownership.

Implemented boundary: src/Library.Core/NativeSqlite.cs owns safe native connections, statements and bounded parameter/row conversion. PiidxSchema.cs owns initial schema migration and strict validation. Neither references Office, vendor code or index research. The native dependency is Windows System32 winsqlite3.dll; x64 FTS5 is verified on this machine. Tests compile as .NET Framework libraries and execute in the available x64 PowerShell CLR. Framework-hosted and x86 runtime execution are not yet verified.

PresentationSource.cs now implements bounded discovery and typed PPTX/POTX extraction, verified against all 62 local packages / 497 slides and synthetic edge cases. See PIIDX_SOURCE_EXTRACTION.md. It is not yet connected to persisted indexing.

Next boundaries, not yet implemented: transactional index writer/reader, safe ranked search, refresh/job/cancellation/recovery orchestration, library registration and enterprise configuration. Connect these to existing LocalLibraryForm, replacing LocalLibraryCache and the installed-vendor source fallback. Keep source preview/insertion on the existing Office UI STA and preserve source hashes/formatting/native edits.

All original user source files are preserved. The merged implementation was largely untracked at baseline; migration commits must not silently absorb unrelated owner changes. Research samples/analysis remain isolated from production compilation and packaging. Existing native template/context acceptance evidence remains historical until the integrated migration passes fresh runtime checks.

Completion requires the complete acceptance matrix from the goal, including real PPTX/POTX indexing, multiple libraries, refresh/error/cancellation/corruption tests, preview/search/gallery/native insertion, packaging, all available regressions and desktop Office save/reopen. Database foundation checks alone do not satisfy these gates.