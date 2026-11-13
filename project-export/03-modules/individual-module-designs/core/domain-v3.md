# Authored chart v3 options and codec

**VERIFIED source review.** This design describes inspected code, not current runtime/Office execution.

## 1. Purpose

Extends model shape to six bar profiles, line, area and pie with discriminated options.

## 2. Implementation status

Authored source is not included in either standalone Chart.Domain.csproj or production VSTO compile items. No deployment/runtime completeness claim.

## 3. Source locations

- `src/Chart.Domain/ChartDocumentV3.cs:1`
- `src/Chart.Domain/ChartV3Codec.cs:1`

## 4. Internal components

ChartCodec common validation, BoundedJson/ContractText, Json.NET OptionsReader. Strict bounded JSON -> whitelist options.kind bar/line/area/pie -> exact DTO shape -> common ID/dimension/style rules -> family-specific constraints. V2 migration adds empty categoryColors and explicit bar options, then codec round-trips.

## 5. Classes and interfaces

- `src/Chart.Domain/ChartDocumentV3.cs:6` — `public sealed class ChartDocumentV3`
- `src/Chart.Domain/ChartDocumentV3.cs:20` — `public sealed class ChartStyleV3`
- `src/Chart.Domain/ChartDocumentV3.cs:27` — `public abstract class ChartOptions`
- `src/Chart.Domain/ChartDocumentV3.cs:30` — `public sealed class BarOptions : ChartOptions`
- `src/Chart.Domain/ChartDocumentV3.cs:33` — `public sealed class LineOptions : ChartOptions`
- `src/Chart.Domain/ChartDocumentV3.cs:39` — `public sealed class AreaOptions : ChartOptions`
- `src/Chart.Domain/ChartDocumentV3.cs:45` — `public sealed class PieOptions : ChartOptions`
- `src/Chart.Domain/ChartV3Codec.cs:13` — `public static class ChartV3Codec`
- `src/Chart.Domain/ChartV3Codec.cs:86` — `private sealed class OptionsReader : JsonConverter`

## 6. Important methods and functions

Declaration locator below is lexical and may include overload/call-like lines; semantic workflow is described separately.

- `src/Chart.Domain/ChartV3Codec.cs:36` — `public static void Validate(ChartDocumentV3 model)`
- `src/Chart.Domain/ChartV3Codec.cs:65` — `public static ChartDocumentV3 Decode(string json)`
- `src/Chart.Domain/ChartV3Codec.cs:71` — `public static string Encode(ChartDocumentV3 model)`
- `src/Chart.Domain/ChartV3Codec.cs:79` — `public static ChartDocumentV3 FromV2(string json, double gapRatio)`
- `src/Chart.Domain/ChartV3Codec.cs:89` — `public override bool CanConvert(Type type)`
- `src/Chart.Domain/ChartV3Codec.cs:90` — `public override object ReadJson(JsonReader reader, Type type, object existing, JsonSerializer serializer)`
- `src/Chart.Domain/ChartV3Codec.cs:92` — `public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)`

## 7. Inputs and outputs

V3 JSON/typed DTO -> validated model; FromV2 requires explicit gapRatio and yields an independent clone.

## 8. Dependencies

ChartCodec common validation, BoundedJson/ContractText, Json.NET OptionsReader.

## 9. Workflow

Strict bounded JSON -> whitelist options.kind bar/line/area/pie -> exact DTO shape -> common ID/dimension/style rules -> family-specific constraints. V2 migration adds empty categoryColors and explicit bar options, then codec round-trips.

## 10. Integration points

V3 JSON/typed DTO -> validated model; FromV2 requires explicit gapRatio and yields an independent clone. Authored source is not included in either standalone Chart.Domain.csproj or production VSTO compile items. No deployment/runtime completeness claim.

## 11. Exceptions and recovery

InvalidDataException v3/options codes, JsonException -> v3.dto. Signed percent stacks and negative/multiple-series pies fail.

## 12. Configuration dependencies

Version3; bar gap0<=g<1; line break/connect/zero; area baseline 0; pie startAngle[-360,360]; category colors only for pie.

## 13. Security considerations

Explicit options type mapping, no executable type materialization. Valid Unicode and finite numbers before use.

## 14. Unit and integration test definitions

- `packages/chart-engine/test/basic-charts.test.mjs:1`

Inspected definitions only. See the parent current test report for execution; historical passes do not establish current acceptance.

## 15. Known limitations

Adjacent normalized-calculation tests are not dedicated v3 codec runtime coverage. No v3 JSON schema or host migration integration established.

## 16. Technical debt

Uncompiled source can drift unnoticed; package model and authored v3 are distinct contracts.

## 17. Extension points

Include explicitly in build graph and add dedicated v3 schema/codec/migration fixtures before exposing to users.
