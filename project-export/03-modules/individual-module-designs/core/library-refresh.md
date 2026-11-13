# Background filesystem refresh coordinator

**VERIFIED source review.** This design describes inspected code, not current runtime/Office execution.

## 1. Purpose

Debounces filesystem events and polls enabled library roots without background Office calls.

## 2. Implementation status

Production Task/timer service; UI caller must marshal notices.

## 3. Source locations

- `src/Library.Core/LibraryRefreshMonitor.cs:1`

## 4. Internal components

FileSystemWatcher,Timer,Task,ConcurrentDictionary,CancellationTokenSource. 500 ms timer; mark dirty on changes/renames/watch error; debounce 750 ms; periodic poll all; one busy Task drains dirty IDs and calls Refresh. Reset replaces watchers; Dispose disables and cancels and releases token source after task completion.

## 5. Classes and interfaces

- `src/Library.Core/LibraryRefreshMonitor.cs:10` — `internal sealed class LibraryRefreshNotice`
- `src/Library.Core/LibraryRefreshMonitor.cs:16` — `internal sealed class LibraryRefreshMonitor : IDisposable`

## 6. Important methods and functions

Declaration locator below is lexical and may include overload/call-like lines; semantic workflow is described separately.

- `src/Library.Core/LibraryRefreshMonitor.cs:32` — `internal LibraryRefreshMonitor(PiidxCatalog catalog,int refreshSeconds,int watchRootLimit,Action<LibraryRefreshNotice> notify)`
- `src/Library.Core/LibraryRefreshMonitor.cs:37` — `internal void Reset()`
- `src/Library.Core/LibraryRefreshMonitor.cs:59` — `internal void Request(string id)`
- `src/Library.Core/LibraryRefreshMonitor.cs:60` — `internal void RequestAll()`
- `src/Library.Core/LibraryRefreshMonitor.cs:92` — `public void Dispose()`

## 7. Inputs and outputs

Catalog+refresh/watch limits+notice callback -> asynchronous result/error notices; Current/IsBusy expose task state.

## 8. Dependencies

FileSystemWatcher,Timer,Task,ConcurrentDictionary,CancellationTokenSource.

## 9. Workflow

500 ms timer; mark dirty on changes/renames/watch error; debounce 750 ms; periodic poll all; one busy Task drains dirty IDs and calls Refresh. Reset replaces watchers; Dispose disables and cancels and releases token source after task completion.

## 10. Integration points

Catalog+refresh/watch limits+notice callback -> asynchronous result/error notices; Current/IsBusy expose task state. Production Task/timer service; UI caller must marshal notices.

## 11. Exceptions and recovery

Watcher errors emit watcher-unavailable-polling-retained; sanitized exception type notices; controlled cancellation; disposed UI callback exceptions tolerated.

## 12. Configuration dependencies

Poll 5–86,400 seconds / watch 0–128; roots beyond the cap remain polled.

## 13. Security considerations

No COM calls or source writes; one-writer catalog lock controls mutation.

## 14. Unit and integration test definitions

- `tests/library/piidx/PiidxWorkspaceTests.cs:1`

Inspected definitions only. See the parent current test report for execution; historical passes do not establish current acceptance.

## 15. Known limitations

Continuous changes may delay debounce; long refresh delays notices; Dispose does not wait synchronously.

## 16. Technical debt

No production event-storm/latency SLO evidence.

## 17. Extension points

Preserve one-worker execution and UI marshalling if adding scheduling/backoff.
