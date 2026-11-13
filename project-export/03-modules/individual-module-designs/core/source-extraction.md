# Read-only PPTX/POTX metadata extraction

**VERIFIED source review.** This design describes inspected code, not current runtime/Office execution.

## 1. Purpose

Reads OOXML ZIP structure, text, tags, layouts/masters and object transforms without Office rendering.

## 2. Implementation status

Production independent reader; Transitional and Strict namespaces handled.

## 3. Source locations

- `src/Library.Core/PresentationSource.cs:1`

## 4. Internal components

ZipArchive, secureXmlReader/XDocument, PathPolicy, SHA256. No OpenXML SDK or COM dependency. Discover bounds trees and skips linked descendants. Read hashes file, parses slide list and relationships/dimensions/core metadata, selects markup-compatibility fallback or first choice, extracts title/tags/layout/master and objects. Structural XML object key tolerates duplicate/zero native IDs. Recheck size/time and full bytes after extraction.

## 5. Classes and interfaces

- `src/Library.Core/PresentationSource.cs:16` — `internal sealed class IndexedObject`
- `src/Library.Core/PresentationSource.cs:23` — `internal sealed class IndexedSlide`
- `src/Library.Core/PresentationSource.cs:31` — `internal sealed class IndexedPresentation`
- `src/Library.Core/PresentationSource.cs:39` — `internal static class PresentationSource`

## 6. Important methods and functions

Declaration locator below is lexical and may include overload/call-like lines; semantic workflow is described separately.

- `src/Library.Core/PresentationSource.cs:46` — `internal static string Fingerprint(string path,CancellationToken cancellation)`
- `src/Library.Core/PresentationSource.cs:57` — `internal static IEnumerable<string> Discover(string root,bool recursive,CancellationToken cancellation)`
- `src/Library.Core/PresentationSource.cs:127` — `internal static IndexedPresentation Read(string path,CancellationToken cancellation)`

## 7. Inputs and outputs

Direct library directory or selected source+CancellationToken -> IndexedPresentation/Slide/Object DTOs and SHA fingerprints.

## 8. Dependencies

ZipArchive, secureXmlReader/XDocument, PathPolicy, SHA256. No OpenXML SDK or COM dependency.

## 9. Workflow

Discover bounds trees and skips linked descendants. Read hashes file, parses slide list and relationships/dimensions/core metadata, selects markup-compatibility fallback or first choice, extracts title/tags/layout/master and objects. Structural XML object key tolerates duplicate/zero native IDs. Recheck size/time and full bytes after extraction.

## 10. Integration points

Direct library directory or selected source+CancellationToken -> IndexedPresentation/Slide/Object DTOs and SHA fingerprints. Production independent reader; Transitional and Strict namespaces handled.

## 11. Exceptions and recovery

InvalidDataException for unsupported package/namespace/IDs/external or ambiguous relationships/limits; XmlException forDTD; IOException on source race; cancellation propagated.

## 12. Configuration dependencies

1 GiB file,100000 members,40 MiB XMLpart,10000 slides,20000 objects/slide,500000 objects/presentation; aggregate slide/tag/object text bounds; discovery10000dirs/100000files.

## 13. Security considerations

DTD prohibited/resolver null; external relationship never fetched; package reads never extract arbitrary ZIP files to filesystem. Source read-shared so hash race detection matters.

## 14. Unit and integration test definitions

- `tests/library/piidx/PiidxSourceTests.cs:1`
- `tests/library/piidx/PiidxEdgeTests.cs:1`
- `tests/library/piidx/PiidxCorpusTests.cs:1`

Inspected definitions only. See the parent current test report for execution; historical passes do not establish current acceptance.

## 15. Known limitations

No PPT/PPTM; not native fidelity/rendering. Group geometry is parent-space EMU, not transformed absolute coordinates. No aggregate total decompressed package budget across all parts.

## 16. Technical debt

Extraction fidelity limited to selected metadata; source scaling benchmarks do not imply Office rendering scalability.

## 17. Extension points

Extend only via explicit package relationships/namespaces/limits and source-preservation tests.
