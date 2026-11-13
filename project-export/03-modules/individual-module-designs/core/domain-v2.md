# Canonical chart v2 and explicit S0 migration

**VERIFIED source review.** This design describes inspected code, not current runtime/Office execution.

## 1. Purpose

Closed identified clustered-column DTO, structural/semantic codec and explicit legacy conversion; no COM mutation.

## 2. Implementation status

VSTO explicitly compiles v2 ChartDocument/ChartCodec. Standalone net48 Chart.Domain project additionally compiles S0Migration and links legacy wire types. Host migration is not automatic.

## 3. Source locations

- `src/Chart.Domain/ChartDocument.cs:1`
- `src/Chart.Domain/ChartCodec.cs:1`
- `src/Chart.Domain/S0Migration.cs:1`
- `src/Chart.Domain/Chart.Domain.csproj:1`

## 4. Internal components

Json.NET13.0.3; invariant culture; MetadataCodec/Contracts for migration. Exact object members and token types checked before deserialization; GUID/revision/global semantic IDs/category-series dimensions/axes/colors/finite point bounds validated. S0Migration creates deterministic IDs from chart ID and ordinal positions, applies explicit layout/style, and clones through Encode/Decode so caller objects do not alias.

## 5. Classes and interfaces

- `src/Chart.Domain/ChartDocument.cs:8` — `public sealed class ChartDocument`
- `src/Chart.Domain/ChartDocument.cs:21` — `public sealed class EmbeddedDataset`
- `src/Chart.Domain/ChartDocument.cs:27` — `public sealed class ChartCategory`
- `src/Chart.Domain/ChartDocument.cs:33` — `public sealed class ChartSeries`
- `src/Chart.Domain/ChartDocument.cs:40` — `public sealed class ChartAxis`
- `src/Chart.Domain/ChartDocument.cs:48` — `public sealed class ChartLayout`
- `src/Chart.Domain/ChartDocument.cs:57` — `public sealed class ChartStyle`
- `src/Chart.Domain/ChartCodec.cs:13` — `public static class ChartCodec`
- `src/Chart.Domain/ChartCodec.cs:141` — `private sealed class StrictReader : JsonTextReader`
- `src/Chart.Domain/S0Migration.cs:8` — `public static class S0Migration`

## 6. Important methods and functions

Declaration locator below is lexical and may include overload/call-like lines; semantic workflow is described separately.

- `src/Chart.Domain/ChartCodec.cs:41` — `public static void Validate(ChartDocument model)`
- `src/Chart.Domain/ChartCodec.cs:48` — `internal static void ValidateContent(ChartDocument model, bool cartesian, Func<string, bool> textRule = null)`
- `src/Chart.Domain/ChartCodec.cs:84` — `public static string Encode(ChartDocument model)`
- `src/Chart.Domain/ChartCodec.cs:91` — `public static ChartDocument Decode(string json)`
- `src/Chart.Domain/ChartCodec.cs:110` — `internal static void CheckShape(JToken token, Type type, Func<JToken, Type, Type> resolve = null)`
- `src/Chart.Domain/ChartCodec.cs:143` — `public StrictReader(TextReader input) : base(input)`
- `src/Chart.Domain/ChartCodec.cs:144` — `public override bool Read()`
- `src/Chart.Domain/S0Migration.cs:12` — `public static ChartDocument FromXml(string xml, ChartLayout layout, ChartStyle style)`

## 7. Inputs and outputs

Canonical JSON <=1,048,576 UTF-16 characters -> ChartDocument or validation failure. Encode accepts typed model. Migration requires legacy XML and caller-supplied layout/style.

## 8. Dependencies

Json.NET13.0.3; invariant culture; MetadataCodec/Contracts for migration.

## 9. Workflow

Exact object members and token types checked before deserialization; GUID/revision/global semantic IDs/category-series dimensions/axes/colors/finite point bounds validated. S0Migration creates deterministic IDs from chart ID and ordinal positions, applies explicit layout/style, and clones through Encode/Decode so caller objects do not alias.

## 10. Integration points

Canonical JSON <=1,048,576 UTF-16 characters -> ChartDocument or validation failure. Encode accepts typed model. Migration requires legacy XML and caller-supplied layout/style. VSTO explicitly compiles v2 ChartDocument/ChartCodec. Standalone net48 Chart.Domain project additionally compiles S0Migration and links legacy wire types. Host migration is not automatic.

## 11. Exceptions and recovery

InvalidDataException coded failures; malformed JsonException becomes document.invalid-json. Unsupported version/type preserved through rejection.

## 12. Configuration dependencies

Version2, clustered-column only, categories1..1000/series1..100, two category/value axes, point dimensions0<d<=1 M.

## 13. Security considerations

TypeNameHandling.None and MissingMemberHandling.Error; duplicate properties/unknown fields/coercion rejected. Size/depth32 limits. V2 custom reader is less strict than the explicit v3/scene grammar and must be tested separately.

## 14. Unit and integration test definitions

- `tests/s1/DomainTests.cs:1`
- `packages/chart-domain/test/contracts.test.mjs:1`

Inspected definitions only. See the parent current test report for execution; historical passes do not establish current acceptance.

## 15. Known limitations

No log/time/break axis model; no automatic default native formatting capture in migration. V2 supports one chart family.

## 16. Technical debt

C#/AJV semantic checks duplicated; standalone project uses explicit compile list, so adding DTO files does not deploy them.

## 17. Extension points

Add new schema/migration and golden cross-language tests; never reinterpret historical v2 silently.
