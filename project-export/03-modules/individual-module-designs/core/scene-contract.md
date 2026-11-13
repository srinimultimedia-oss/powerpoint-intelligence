# Scene DTOs and cross-language hash boundary

<!-- generated-toc -->
## Contents

- [1. Purpose](#1-purpose)
- [2. Implementation status](#2-implementation-status)
- [3. Source locations](#3-source-locations)
- [4. Internal components](#4-internal-components)
- [5. Classes and interfaces](#5-classes-and-interfaces)
- [6. Important methods and functions](#6-important-methods-and-functions)
- [7. Inputs and outputs](#7-inputs-and-outputs)
- [8. Dependencies](#8-dependencies)
- [9. Workflow](#9-workflow)
- [10. Integration points](#10-integration-points)
- [11. Exceptions and recovery](#11-exceptions-and-recovery)
- [12. Configuration dependencies](#12-configuration-dependencies)
- [13. Security considerations](#13-security-considerations)
- [14. Unit and integration test definitions](#14-unit-and-integration-test-definitions)
- [15. Known limitations](#15-known-limitations)
- [16. Technical debt](#16-technical-debt)
- [17. Extension points](#17-extension-points)
<!-- end-generated-toc -->

**VERIFIED source review.** This design describes inspected code, not current runtime/Office execution.

## 1. Purpose

Closed point-based geometry and paint/diagnostic model with deterministic verified SHA-256.

## 2. Implementation status

Compiled production source. Six DTO geometry types do not imply six native renderer capabilities.

## 3. Source locations

- `src/Chart.Renderers/SceneContracts.cs:1`
- `src/Chart.Renderers/SceneCodec.cs:1`
- `src/Chart.Renderers/StrictSceneJson.cs:1`

## 4. Internal components

BoundedJson/ContractText, Json.NET, SHA-256; no Office objects. Strict parse -> exact fields/ranges/chart-scoped unique IDs/diagnostic references -> recompute canonical binary hash -> explicit GeometryReader subtype. Hash excludes outer renderHash; ordinal object keys, ordered arrays, big-endian doubles, normalized negative zero, domain separator PPI-SCENE-1.

## 5. Classes and interfaces

- `src/Chart.Renderers/SceneContracts.cs:7` — `public sealed class RenderResult`
- `src/Chart.Renderers/SceneContracts.cs:21` — `public sealed class ScenePlot`
- `src/Chart.Renderers/SceneContracts.cs:30` — `public sealed class SceneNode`
- `src/Chart.Renderers/SceneContracts.cs:37` — `public sealed class ScenePaint`
- `src/Chart.Renderers/SceneContracts.cs:45` — `public sealed class SceneDiagnostic`
- `src/Chart.Renderers/SceneContracts.cs:51` — `public abstract class SceneGeometry`
- `src/Chart.Renderers/SceneContracts.cs:54` — `public class BoxGeometry : SceneGeometry`
- `src/Chart.Renderers/SceneContracts.cs:62` — `public sealed class TextGeometry : BoxGeometry`
- `src/Chart.Renderers/SceneContracts.cs:70` — `public sealed class PathGeometry : SceneGeometry`
- `src/Chart.Renderers/SceneContracts.cs:73` — `public sealed class ScenePoint`
- `src/Chart.Renderers/SceneContracts.cs:79` — `public sealed class SectorGeometry : SceneGeometry`
- `src/Chart.Renderers/SceneCodec.cs:14` — `public static class SceneCodec`
- `src/Chart.Renderers/SceneCodec.cs:164` — `private sealed class GeometryReader : JsonConverter`
- `src/Chart.Renderers/StrictSceneJson.cs:5` — `internal static class StrictSceneJson`
- `src/Chart.Renderers/StrictSceneJson.cs:9` — `internal static class SceneText`

## 6. Important methods and functions

Declaration locator below is lexical and may include overload/call-like lines; semantic workflow is described separately.

- `src/Chart.Renderers/SceneCodec.cs:106` — `public static RenderResult Decode(string json)`
- `src/Chart.Renderers/SceneCodec.cs:113` — `public static void ValidateJsonSyntax(string json)`
- `src/Chart.Renderers/SceneCodec.cs:120` — `public static string Encode(RenderResult scene)`
- `src/Chart.Renderers/SceneCodec.cs:125` — `public static string ComputeHash(RenderResult scene)`
- `src/Chart.Renderers/SceneCodec.cs:167` — `public override bool CanConvert(Type type)`
- `src/Chart.Renderers/SceneCodec.cs:168` — `public override object ReadJson(JsonReader reader, Type type, object existingValue, JsonSerializer serializer)`
- `src/Chart.Renderers/SceneCodec.cs:181` — `public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)`
- `src/Chart.Renderers/StrictSceneJson.cs:7` — `internal static JToken Parse(string json)`
- `src/Chart.Renderers/StrictSceneJson.cs:11` — `internal static bool Valid(string value, int maximum)`
- `src/Chart.Renderers/StrictSceneJson.cs:12` — `internal static bool NonWhitespace(string value)`

## 7. Inputs and outputs

Sealed scene JSON <=8,388,608 chars -> RenderResult; draft DTO may compute hash but Encode/Decode require verified lowercase hash.

## 8. Dependencies

BoundedJson/ContractText, Json.NET, SHA-256; no Office objects.

## 9. Workflow

Strict parse -> exact fields/ranges/chart-scoped unique IDs/diagnostic references -> recompute canonical binary hash -> explicit GeometryReader subtype. Hash excludes outer renderHash; ordinal object keys, ordered arrays, big-endian doubles, normalized negative zero, domain separator PPI-SCENE-1.

## 10. Integration points

Sealed scene JSON <=8,388,608 chars -> RenderResult; draft DTO may compute hash but Encode/Decode require verified lowercase hash. Compiled production source. Six DTO geometry types do not imply six native renderer capabilities.

## 11. Exceptions and recovery

InvalidDataException scene.* on invalid types/IDs/hash/bounds/Unicode; JsonException becomes scene.dto.

## 12. Configuration dependencies

Schema1/pt, nodes<= 2,00000, diagnostics<=10000, path points<=10000, text4096/font128 codepoints, font size1..1000.

## 13. Security considerations

Whitelisted rectangles/ellipses/polylines/polygons/sectors/text, hex colors or null, no executable nodes or URL paint. Unkeyed hash proves consistency, not authentication/permission.

## 14. Unit and integration test definitions

- `tests/s3/SceneContractTests.cs:1`
- `tests/s3/validate-roundtrip.mjs:1`
- `packages/chart-engine/test/scene-schema.test.mjs:1`
- `packages/chart-engine/test/scene-contract.test.mjs:1`

Inspected definitions only. See the parent current test report for execution; historical passes do not establish current acceptance.

## 15. Known limitations

JSON Schema cannot establish hash/unique semantic linkage; native adapter only rectangle profile.

## 16. Technical debt

Canonical bytes are compatibility-critical; a JSON-string hash is not equivalent.

## 17. Extension points

Add geometry only with matching schema/DTO/hash fixtures and explicit host capability.
