# Independent presentation source extraction

Implemented in src/Library.Core/PresentationSource.cs. This layer discovers and reads original PPTX/POTX packages without Office automation, vendor software, a proprietary index or source mutation. It is not yet connected to a persisted PIIDX writer/gallery.

## Actual behavior

- Iterative bounded discovery with cancellation; supports recursive roots, filters PPTX/POTX case-insensitively and excludes reparse points.
- SHA256 before/after reading and source size/UTC timestamp guards reject observed concurrent source changes. Streams file hashing; holds one presentation's metadata at a time.
- Parses transitional and strict Open XML namespaces, slide order/native IDs, package relationships, dimensions, author/keywords, title/text, layout/master names, user-defined tag parts and supported native object records.
- Keeps native slide identity separate from position. Native object IDs may legitimately be zero/repeated in the supplied source packages; those values are retained, with separate structural keys scoped to each slide. This prevents dropping objects or pretending ambiguous native IDs are unique.
- Preserves raw parent-coordinate transforms, rotation/flips, group relationships and source-XML fingerprints. It does not claim those coordinates are absolute slide coordinates.
- Handles multiple tag parts and observed 32K metadata values. Tag values are not printed in diagnostics. Native opaque data is not decoded or executed.
- Rejects DTD/external entity expansion and external slide/layout/tag fetches. Source documents remain unchanged. Markup-compatibility fallback is preferred for metadata; native source rendering/insertion must continue using the original package.

## Bounds

Maximum presentation: 1 GiB. XML part: 40 MiB/characters. Package members: 100,000. Slides: 10,000 per presentation. Objects: 20,000 per slide / 500,000 per presentation. Slide text: 640,000 characters; presentation slide text: 64 MiB; accumulated object text: 128 MiB. Tags: 512 per slide; names 256 characters, values 65,536 characters; total presentation tag text 32 MiB. These are explicit resource bounds, not performance certification.

## Verified

scripts/Test-PiidxSourceReader.ps1 passes 24 assertions on the available x64 runtime. It covers synthetic PPTX/POTX, native IDs across reordered slides, Unicode filenames/text, empty slides, core/object/transform/tag metadata, strict OOXML, external source rejection, DTD rejection and cancellation. It also reads all 62 copied local packages / 497 slides and verifies unique scoped object keys without discarding repeated native IDs. .NET Framework compiler builds passed for AnyCPU/x86/x64; runtime execution was x64 only.

No Office rendering, insertion, native object modification, animation playback or save/reopen was performed by these tests. Synthetic package extraction is not native PowerPoint round-trip acceptance. Full ZIP/OOXML conformance validation, deep chart/embedded-document content extraction, corrupt/large/read-only/long-path race coverage and persistent catalog identity still need broader engine/integration tests.

The PowerPoint tag-list model is documented by [Microsoft](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.taglist?view=openxml-3.0.1). Actual source-package tests, not this reference alone, establish the multiple-part/large-value observations.