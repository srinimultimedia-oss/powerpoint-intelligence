# Bounded JSON and Unicode utilities

**VERIFIED source review.** This design describes inspected code, not current runtime/Office execution.

## 1. Purpose

Strict JSON grammar and Unicode validation before scene/v3 DTO conversion.

## 2. Implementation status

Linked source compiled into production VSTO; no independent runtime service.

## 3. Source locations

- `src/Shared.Contracts/BoundedJson.cs:1`

## 4. Internal components

Json.NET.Linq, System.Text/IO/Globalization; no COM/network/filesystem operations. Recursive cursor parser accepts JSON whitespace only, enforces depth32, ordinal duplicate keys, exact number/escape grammar, finite invariant double values and no trailing content. NonWhitespace mirrors ECMAScript rather than culture whitespace.

## 5. Classes and interfaces

- `src/Shared.Contracts/BoundedJson.cs:11` — `internal sealed class BoundedJson`
- `src/Shared.Contracts/BoundedJson.cs:108` — `internal static class ContractText`

## 6. Important methods and functions

Declaration locator below is lexical and may include overload/call-like lines; semantic workflow is described separately.

- `src/Shared.Contracts/BoundedJson.cs:19` — `internal static JToken Parse(string json, int maximumCharacters, string prefix)`
- `src/Shared.Contracts/BoundedJson.cs:110` — `internal static bool Valid(string value, int maxCodePoints)`
- `src/Shared.Contracts/BoundedJson.cs:122` — `internal static bool NonWhitespace(string value)`

## 7. Inputs and outputs

String + maxCharacters + error prefix -> JToken; ContractText checks paired surrogates/codepoint count/non-whitespace.

## 8. Dependencies

Json.NET.Linq, System.Text/IO/Globalization; no COM/network/filesystem operations.

## 9. Workflow

Recursive cursor parser accepts JSON whitespace only, enforces depth32, ordinal duplicate keys, exact number/escape grammar, finite invariant double values and no trailing content. NonWhitespace mirrors ECMAScript rather than culture whitespace.

## 10. Integration points

String + maxCharacters + error prefix -> JToken; ContractText checks paired surrogates/codepoint count/non-whitespace. Linked source compiled into production VSTO; no independent runtime service.

## 11. Exceptions and recovery

InvalidDataException prefix.invalid-json or prefix.json-size; no partial DTO escapes.

## 12. Configuration dependencies

Caller-specific length cap; nesting32.

## 13. Security considerations

Reject JavaScript comments/single quotes/trailing commas/constructors, unpaired surrogates, nonfinite numbers; no reflection/type loading.

## 14. Unit and integration test definitions

- `tests/s3/SceneContractTests.cs:1`
- `packages/chart-engine/test/scene-schema.test.mjs:1`
- `tests/s3/scene-fixtures.mjs:1`

Inspected definitions only. See the parent current test report for execution; historical passes do not establish current acceptance.

## 15. Known limitations

Full token tree allocation; numeric tokens represented as double, integer range enforced later.

## 16. Technical debt

C# and TS handwritten grammar must remain behaviorally aligned.

## 17. Extension points

Change only with shared positive/negative syntax corpus and codepoint boundary fixtures.
