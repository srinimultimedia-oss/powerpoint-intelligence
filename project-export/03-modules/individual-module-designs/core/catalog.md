# Transactional PIIDX registration and refresh

**VERIFIED source review.** This design describes inspected code, not current runtime/Office execution.

## 1. Purpose

Persists independent libraries and incremental source snapshots with cross-process writer ownership and logical integrity.

## 2. Implementation status

Production local .piidx file, disposable connection per operation; no proprietary index dependency.

## 3. Source locations

- `src/Library.Core/PiidxCatalog.cs:1`

## 4. Internal components

PathPolicy, schema/SQLite, PresentationSource, SHA256. Open validates schema/content; Register reuses case-insensitive root IDs. Writer .writer.lock denies sharing and waits<=5 s cancellably. Refresh records running job, discovers/hashes sources, begins immediate transaction, preserves unique-hash rename/native-slide reorder identity, replaces changed slides, marks missing sources, rehashes captured files, updates digest/FTS/job, commits. Failures retain previous snapshot and write sanitized errors.

## 5. Classes and interfaces

- `src/Library.Core/PiidxCatalog.cs:13` — `internal sealed partial class PiidxCatalog`

## 6. Important methods and functions

Declaration locator below is lexical and may include overload/call-like lines; semantic workflow is described separately.

- `src/Library.Core/PiidxCatalog.cs:18` — `internal static PiidxCatalog Open(string path)`
- `src/Library.Core/PiidxCatalog.cs:24` — `internal static PiidxCatalog OpenOrCreate(string path)`
- `src/Library.Core/PiidxCatalog.cs:54` — `internal IReadOnlyList<LibraryRegistration> Libraries()`
- `src/Library.Core/PiidxCatalog.cs:58` — `internal string Register(string root,string name,bool recursive)`
- `src/Library.Core/PiidxCatalog.cs:77` — `internal void Remove(string libraryId)`
- `src/Library.Core/PiidxCatalog.cs:84` — `internal void VerifyIntegrity()`
- `src/Library.Core/PiidxCatalog.cs:152` — `internal LibraryRefreshResult Refresh(string libraryId,CancellationToken cancellation,IProgress<LibraryIndexProgress> progress=null,bool force=false)`

## 7. Inputs and outputs

Index/library/path+cancel/progress -> registrations/LibraryRefreshResult and persisted authoritative rows.

## 8. Dependencies

PathPolicy, schema/SQLite, PresentationSource, SHA256.

## 9. Workflow

Open validates schema/content; Register reuses case-insensitive root IDs. Writer .writer.lock denies sharing and waits<=5 s cancellably. Refresh records running job, discovers/hashes sources, begins immediate transaction, preserves unique-hash rename/native-slide reorder identity, replaces changed slides, marks missing sources, rehashes captured files, updates digest/FTS/job, commits. Failures retain previous snapshot and write sanitized errors.

## 10. Integration points

Index/library/path+cancel/progress -> registrations/LibraryRefreshResult and persisted authoritative rows. Production local .piidx file, disposable connection per operation; no proprietary index dependency.

## 11. Exceptions and recovery

Corrupt/foreign indexes preserved; cancellation/source race rolls back. Stale running markers become interrupted only after actual OS writer handle acquired. Remove deletes index rows, never sources.

## 12. Configuration dependencies

.piidx extension,1000 libraries, direct registered roots/index/lock rejectreparse. SlideId deterministic truncatedSHA256(presentationId/nativeSlideId).

## 13. Security considerations

Digests cover libraries/presentations/slides/objects/tags/FTS and source rows; unkeyed consistency, not cryptographic authorization. Bound SQL values, index/lock owned path.

## 14. Unit and integration test definitions

- `tests/library/piidx/PiidxCatalogTests.cs:1`
- `tests/library/piidx/PiidxCorpusTests.cs:1`

Inspected definitions only. See the parent current test report for execution; historical passes do not establish current acceptance.

## 15. Known limitations

Full integrity recomputation and transactional refresh scale with catalog; missing data retained but hidden; rename identity only unambiguous content hash.

## 16. Technical debt

No production size/latency SLO. Long refresh retains writer lock/transaction; unchanged sources still rehashed.

## 17. Extension points

Add explicit migration and performance tests; preserve committed-reader snapshots and source-race refusal.
