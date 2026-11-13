# PIIDX query models, search and navigation

**VERIFIED source review.** This design describes inspected code, not current runtime/Office execution.

## 1. Purpose

Full-text/metadata search, deterministic pages and freshness-protected slide identity navigation.

## 2. Implementation status

Production local backend. UI rendering/slide insertion belongs to VSTO gallery.

## 3. Source locations

- `src/Library.Core/LibraryModels.cs:1`
- `src/Library.Core/PiidxSearch.cs:1`
- `src/Library.Core/PiidxNavigation.cs:1`

## 4. Internal components

PiidxCatalog, FTS5 bm25, parameterized SQL, cancellation. Convert query to quoted FTS words/phrases/prefixes or exact title/text predicate; add library/presentation/category/date/aspect/tag/author/layout filters. Duplicate representative chosen inside filtered candidates. Count+page read transaction with stable source/position/id tie breaks. VerifySource compares selected/current identity/path/hash/position and actual file fingerprint.

## 5. Classes and interfaces

- `src/Library.Core/LibraryModels.cs:5` — `internal sealed class LibraryRegistration`
- `src/Library.Core/LibraryModels.cs:14` — `internal sealed class LibraryIndexProgress`
- `src/Library.Core/LibraryModels.cs:22` — `internal sealed class LibraryRefreshResult`
- `src/Library.Core/LibraryModels.cs:30` — `internal sealed class LibrarySlideReference`
- `src/Library.Core/LibraryModels.cs:51` — `internal enum LibrarySortMode`
- `src/Library.Core/LibraryModels.cs:52` — `internal enum LibraryAspect`
- `src/Library.Core/LibraryModels.cs:53` — `internal enum LibrarySearchMode`
- `src/Library.Core/LibraryModels.cs:54` — `internal sealed class LibrarySearchRequest`
- `src/Library.Core/LibraryModels.cs:72` — `internal sealed class LibrarySearchPage`
- `src/Library.Core/PiidxSearch.cs:11` — `internal sealed partial class PiidxCatalog`
- `src/Library.Core/PiidxNavigation.cs:7` — `internal sealed class LibraryCategorySummary`
- `src/Library.Core/PiidxNavigation.cs:12` — `internal sealed class LibraryPresentationSummary`
- `src/Library.Core/PiidxNavigation.cs:20` — `internal sealed partial class PiidxCatalog`

## 6. Important methods and functions

Declaration locator below is lexical and may include overload/call-like lines; semantic workflow is described separately.

- `src/Library.Core/PiidxSearch.cs:30` — `internal LibrarySearchPage Search(LibrarySearchRequest request,CancellationToken cancellation)`
- `src/Library.Core/PiidxSearch.cs:84` — `internal LibrarySlideReference Slide(string slideId,CancellationToken cancellation)`
- `src/Library.Core/PiidxSearch.cs:98` — `internal void VerifySource(LibrarySlideReference slide,CancellationToken cancellation)`
- `src/Library.Core/PiidxNavigation.cs:18` — `public override string ToString()`
- `src/Library.Core/PiidxNavigation.cs:22` — `internal IReadOnlyList<LibraryCategorySummary> Categories(string libraryId,CancellationToken cancellation)`
- `src/Library.Core/PiidxNavigation.cs:26` — `internal IReadOnlyList<LibraryPresentationSummary> Presentations(string libraryId,CancellationToken cancellation)`

## 7. Inputs and outputs

LibrarySearchRequest -> LibrarySearchPage(total,IReadOnlyList slide refs); Slide/VerifySource and Categories/Presentations provide identity/navigation.

## 8. Dependencies

PiidxCatalog, FTS5 bm25, parameterized SQL, cancellation.

## 9. Workflow

Convert query to quoted FTS words/phrases/prefixes or exact title/text predicate; add library/presentation/category/date/aspect/tag/author/layout filters. Duplicate representative chosen inside filtered candidates. Count+page read transaction with stable source/position/id tie breaks. VerifySource compares selected/current identity/path/hash/position and actual file fingerprint.

## 10. Integration points

LibrarySearchRequest -> LibrarySearchPage(total,IReadOnlyList slide refs); Slide/VerifySource and Categories/Presentations provide identity/navigation. Production local backend. UI rendering/slide insertion belongs to VSTO gallery.

## 11. Exceptions and recovery

ArgumentException for invalid enums/bounds/queries; missing source FileNotFoundException; stale selection InvalidOperationException; cancellation.

## 12. Configuration dependencies

Query<=4096 chars,<=64 parsedterms; page1..1000 default100, offset<=10 M; wide/standard ratio tolerance0.02.

## 13. Security considerations

Never raw callerSQL; canonical nonemptyGUIDs. Search requires indexed sources+enabledlibraries. Direct Slide identity path is not RBAC or enterprise authorization.

## 14. Unit and integration test definitions

- `tests/library/piidx/PiidxCatalogTests.cs:1`
- `tests/library/piidx/PiidxWorkspaceTests.cs:1`

Inspected definitions only. See the parent current test report for execution; historical passes do not establish current acceptance.

## 15. Known limitations

Duplicate fingerprint is source-derived, not pixel/visual equivalence. FTS relevance is lexical, not AI semantic search.

## 16. Technical debt

SQLite Unicode folding and invariant tag normalization need explicit multilingual expectations.

## 17. Extension points

Add filters as bound predicates with count/page/paging/duplicate composition tests.
