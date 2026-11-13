# Thumbnail gallery and PIIDX workspace UI

**VERIFIED — source implementation. Current runtime acceptance is not established by this design review.**

## Purpose and implementation status

Browse indexed local slides with thumbnails, filters/navigation and controlled insertion in the current presentation.

Implemented local source is distinguished from supported subset/absent external behavior below. See current test evidence in the export testing chapter; do not treat historical logs or method existence as a current runtime pass.

## Source locations and internal components

- `apps/powerpoint-vsto/LocalLibraryGallery.cs:1` — LibraryGallery, LocalLibraryForm.
- `apps/powerpoint-vsto/LocalLibraryWorkspace.cs:1` — LocalLibraryForm, UiAwaiter, LibraryChoice.

## Inputs, outputs, data structures and dependencies

Inputs: local configured sources, search/filter/page selections and SourceReference. Outputs: gallery, catalog/navigation changes and inserted native slides. Dependencies: Library.Core PIIDX/source/preview/refresh APIs and SourceTemplateSession; no proprietary index runtime.

## Internal execution and external integration

LocalLibraryForm uses large-icon ListView/paging and selected larger preview. Workspace loads profile/catalog, indexes configured sources on background tasks and exposes library/presentation/search/tag controls and coalesced refresh signals. Catalog/file work uses cancellation and background tasks; native preview/insertion returns to captured UI thread through explicit awaiter. Queue renders one preview at a time, caches up to96thumbnails and shows placeholders/errors. Insertion revalidates fingerprint and selection before keep-source/destination-design path. Dispose cancels work and releases monitor/catalog/source/images/timers.

## Configuration, validation and failure behavior

LibraryIndexPath, LibrarySourcesConfig, LibraryRefreshSeconds, LibraryWatchRootLimit and TemplateLibraryFolder. Cache identities include source revisions. Native preview uses Office; indexed metadata alone does not establish rendering.

## Security considerations and cleanup

Local native calls run in the current Office user process and require supported selection/ownership; a mutex/selection key is not authentication or RBAC. File/metadata content must remain bounded and error text must not expose data or credentials. Explicit owned-resource cleanup and error/rollback paths are part of the cited implementation, not a guarantee that a failed COM server can always be restored.

## Test references

Lexical module-class matches in test source (not instrumentation coverage; no test is claimed executed here):

- `tests/library/GalleryTests.cs:19`.
- `tests/library/PiidxGalleryTests.cs:30`.

## Known limitations and technical debt

Disposal/UI-thread fallback races need stress tests. File watchers are hints, not atomic snapshots. Native previews can be denied under restricted token. Cancellation/source revision mismatch prevents stale navigation/insertion; errors remain explicit per-thumbnail.

## Extension points

Extend typed search modes/pagination with cancellation and STA-boundary tests; retain distinction between package extraction and native Office rendering.

## Classes, interfaces and method navigation

Method names below are source declarations; private/internal members are not public API promises. The source hash/review manifest gives exact reviewed bytes.

| Declaration | Kind | Source |
|---|---|---|
| `LibraryGallery` | class | `apps/powerpoint-vsto/LocalLibraryGallery.cs:16` |
| `Caption` | method | `apps/powerpoint-vsto/LocalLibraryGallery.cs:19` |
| `Page` | method | `apps/powerpoint-vsto/LocalLibraryGallery.cs:21` |
| `LocalLibraryForm` | class | `apps/powerpoint-vsto/LocalLibraryGallery.cs:24` |
| `ShowAssets` | method | `apps/powerpoint-vsto/LocalLibraryGallery.cs:79` |
| `Filter` | method | `apps/powerpoint-vsto/LocalLibraryGallery.cs:88` |
| `MovePage` | method | `apps/powerpoint-vsto/LocalLibraryGallery.cs:94` |
| `Key` | method | `apps/powerpoint-vsto/LocalLibraryGallery.cs:96` |
| `Frame` | method | `apps/powerpoint-vsto/LocalLibraryGallery.cs:97` |
| `ShowPage` | method | `apps/powerpoint-vsto/LocalLibraryGallery.cs:106` |
| `UpdateDetails` | method | `apps/powerpoint-vsto/LocalLibraryGallery.cs:116` |
| `RememberThumbnail` | method | `apps/powerpoint-vsto/LocalLibraryGallery.cs:121` |
| `RenderNextPreview` | method | `apps/powerpoint-vsto/LocalLibraryGallery.cs:126` |
| `RefreshIndex` | method | `apps/powerpoint-vsto/LocalLibraryGallery.cs:135` |
| `Insert` | method | `apps/powerpoint-vsto/LocalLibraryGallery.cs:136` |
| `Dispose` | method | `apps/powerpoint-vsto/LocalLibraryGallery.cs:143` |
| `LocalLibraryForm` | class | `apps/powerpoint-vsto/LocalLibraryWorkspace.cs:14` |
| `UiAwaiter` | struct | `apps/powerpoint-vsto/LocalLibraryWorkspace.cs:17` |
| `GetAwaiter` | method | `apps/powerpoint-vsto/LocalLibraryWorkspace.cs:21` |
| `GetResult` | method | `apps/powerpoint-vsto/LocalLibraryWorkspace.cs:23` |
| `OnCompleted` | method | `apps/powerpoint-vsto/LocalLibraryWorkspace.cs:24` |
| `LibraryChoice` | class | `apps/powerpoint-vsto/LocalLibraryWorkspace.cs:46` |
| `ToString` | method | `apps/powerpoint-vsto/LocalLibraryWorkspace.cs:49` |
