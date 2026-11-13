# Validated PNG derivative cache

**VERIFIED source review.** This design describes inspected code, not current runtime/Office execution.

## 1. Purpose

Persists bounded PNG thumbnails tied to exact indexed source/slide fingerprints.

## 2. Implementation status

Production derivative storage. Does not generate images or establish native rendering fidelity.

## 3. Source locations

- `src/Library.Core/PiidxPreviewCache.cs:1`

## 4. Internal components

Catalog/SQLite, CRC32 chunk framing, SHA256. Clone caller bytes, verify PNG signature/chunk order/types/CRC/size/encoding; transaction requires current selection fingerprints; replace same-size row, evict oldest-rendered rows above256MiB. Read checks matching slide fingerprint, checksum and PNG framing before returning.

## 5. Classes and interfaces

- `src/Library.Core/PiidxPreviewCache.cs:8` — `internal static class PiidxPng`
- `src/Library.Core/PiidxPreviewCache.cs:24` — `if(type.Any(c`
- `src/Library.Core/PiidxPreviewCache.cs:43` — `internal sealed partial class PiidxCatalog`

## 6. Important methods and functions

Declaration locator below is lexical and may include overload/call-like lines; semantic workflow is described separately.

- `src/Library.Core/PiidxPreviewCache.cs:16` — `internal static void Validate(byte[] png,int width,int height)`
- `src/Library.Core/PiidxPreviewCache.cs:41` — `internal static string Hash(byte[] bytes)`
- `src/Library.Core/PiidxPreviewCache.cs:45` — `internal void PutPreview(LibrarySlideReference slide,byte[] png,int width,int height,CancellationToken cancellation)`
- `src/Library.Core/PiidxPreviewCache.cs:64` — `internal byte[] Preview(string slideId,int width,int height,CancellationToken cancellation)`

## 7. Inputs and outputs

LibrarySlideReference+PNG+size+cancel -> stored row; Preview -> independent byte[] or null for absent/stale.

## 8. Dependencies

Catalog/SQLite, CRC32 chunk framing, SHA256.

## 9. Workflow

Clone caller bytes, verify PNG signature/chunk order/types/CRC/size/encoding; transaction requires current selection fingerprints; replace same-size row, evict oldest-rendered rows above256MiB. Read checks matching slide fingerprint, checksum and PNG framing before returning.

## 10. Integration points

LibrarySlideReference+PNG+size+cancel -> stored row; Preview -> independent byte[] or null for absent/stale. Production derivative storage. Does not generate images or establish native rendering fidelity.

## 11. Exceptions and recovery

InvalidDataException for damaged PNG/hash; stale selection InvalidOperationException; cancellation beforecommit.

## 12. Configuration dependencies

PNG<=8 MiB; width/height1..4096; aggregate image data256MiB; image/png only.

## 13. Security considerations

Corrupt derivative cannot authorize dropping source records. Previews outside authoritative logicaldigest are independently hashed/validated.

## 14. Unit and integration test definitions

- `tests/library/piidx/PiidxPreviewTests.cs:1`

Inspected definitions only. See the parent current test report for execution; historical passes do not establish current acceptance.

## 15. Known limitations

PNG validator does not decompress IDAT or measure image fidelity; test fixture is synthetic1px. Eviction by render timestamp, not last use; logicalbytes not physicaldatabase size.

## 16. Technical debt

Physical file compaction and native rendering coverage remain operational considerations.

## 17. Extension points

New providers must preserve current-snapshot checks and native renderer acceptance evidence.
