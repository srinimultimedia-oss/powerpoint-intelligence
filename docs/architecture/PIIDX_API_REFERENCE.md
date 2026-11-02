# PIIDX API reference — implemented core

Internal C# namespace: PowerPointIntelligence.Library. Target: .NET Framework 4.8. The gallery/ribbon connection is still pending; these are implemented, executable core APIs, not claims of PowerPoint runtime acceptance.

## Catalog

- PiidxCatalog.Open(path): opens an existing .piidx read-only for validation, verifies schema, SQLite integrity, header identity and canonical content/source fingerprints in a read transaction. Foreign/corrupt files are rejected and preserved.
- OpenOrCreate(path): creates only a new/empty file under the exclusive writer protocol, or validates an existing catalog. Never silently recreates a corrupt index.
- Libraries(): returns typed persistent registrations, recursion/enabled state and last refresh UTC.
- Register(root,name,recursive): verifies a direct accessible directory, reuses an existing case-insensitive root identity or creates one, and commits registration/digest changes. It does not mutate source presentations.
- Remove(libraryId): cascades catalog/FTS/object/tag/preview records, preserves original files and other libraries, and keeps historical job references through their defined foreign keys.
- Refresh(libraryId,cancellation,progress,force=false): transactional library rebuild/incremental refresh. Returns job ID and file/slide counts. force=true reparses originals while retaining catalog identities. Any source-read failure/cancellation/change observed after capture rolls back the library snapshot; sanitized job/error status remains recorded.
- VerifyIntegrity(): deep schema, SQLite page/foreign-key and canonical content/source-digest verification. Derivative PNG cache bytes are verified separately when read; a bad preview does not discard the source catalog.

## Search and source references

- Search(request,cancellation): returns total count plus a bounded result page (1–1000 rows). AllWords supports quoted phrases and explicit suffix stars; Phrase enforces adjacency; Prefix prefixes unquoted terms; ExactText compares complete literal title/body values. User operators/punctuation are tokenized/quoted and bound, never concatenated into SQL. FTS5 bm25 weights title/body/keywords 5/1/2. Exact category, normalized tag, library/presentation, author/layout filters are implemented. Empty query browses indexed enabled sources. Missing sources are excluded.
- Slide(slideId,cancellation): looks up identity directly, independent of the selected page or source slide position.
- VerifySource(reference,cancellation): rejects changed source/hash/position/native ID or stale indexed selection before insertion. It performs no COM calls.
- Stable slide IDs derive from persisted presentation identity and native PowerPoint slide ID. Unique full-content-hash renames retain presentation identity; ambiguous matches are not guessed. Reordered native slides keep identity while position changes. Missing rows remain recorded and can regain their identities when restored.

## Preview cache

- PutPreview(reference,png,width,height,cancellation): persists validated PNG bytes only for the current indexed reference, checks dimensions/CRC/size, and evicts oldest cache rows above a 256 MiB payload budget. This method does not claim that caller-provided pixels came from PowerPoint; the native preview provider must establish that at integration.
- Preview(slideId,width,height,cancellation): returns independent verified PNG bytes, null for a missing/stale derivative, or an explicit error for corruption. The cache is keyed by slide identity, dimensions and source slide fingerprint. Rebuild/changed source rows invalidate dependent previews.
- PNG bounds: 8 MiB per image, 4096×4096 maximum. PNG signature/chunk bounds/CRC/header/dimensions/critical chunks/ending are checked before UI decoding. Full pixel decoding and native visual fidelity are separate runtime gates.

## Native SQLite and extraction

SqliteDatabase owns System32-only winsqlite3.dll connections/statements via SafeHandle; typed UTF-8/blob/numeric/null binding; single-statement enforcement; bounded row values; read-only access; busy timeout; trusted-schema/foreign-key settings; and cancellable native long queries. RollbackIfActive handles SQLite auto-rollback after interruption. Source extraction APIs and explicit limits are in PIIDX_SOURCE_EXTRACTION.md. Production code never opens a proprietary index.

## Reproduction and evidence

Run scripts/Test-Piidx.ps1 from the repository. It compiles/reuses one stable artifacts/piidx/PiidxTestSuite.dll with the Visual Studio C# 7.3 compiler, executes six groups and records source hashes, binary hash, runtime, architecture and results. Individual Test-PiidxFoundation/SourceReader/Catalog/Previews/Corpus scripts use the same suite. Loaded assemblies with changed inputs require a fresh PowerShell session; missing/modified/unreceipted assemblies fail closed rather than being restored blindly. Test-All includes piidx-engine and keeps the DLL at the actual repository RunnerRoot while executing captured sources.

Current x64 results: 28 foundation + 24 extraction + 46 catalog + 14 preview + 4 edge + 7 corpus assertions. The corpus contains 62 packages, 497 slides, 13049 objects and 2078 tags. Compiler builds cover AnyCPU/x86/x64; x86 runtime remains unexecuted because Windows rejected the script host. No Office rendering/insertion is inferred from these results.