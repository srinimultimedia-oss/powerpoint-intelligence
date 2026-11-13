# Published chart, scene and rendered-record schemas

**VERIFIED source review.** This design describes inspected code, not current runtime/Office execution.

## 1. Purpose

Distributes closed structural validation specifications for the canonical v2, semantic scene v1 and persisted rendered record v1 boundaries.

## 2. Implementation status

Authored versioned schemas are consumed by AJV tests and .NET XSD tests. They are specifications, not independent services or authorization policies.

## 3. Source locations

- `packages/chart-domain/schema/chart-document.v2.schema.json:1`
- `packages/chart-engine/schema/render-result.v1.schema.json:1`
- `packages/chart-engine/schema/rendered-chart.v1.xsd:1`

## 4. Internal components

AJV2020 for JSON Schema draft2020-12; System.Xml.Schema for XSD; no Office or network retrieval is required for local schema use. Chart schema defines closed required object fields, version/type constants, semantic ID text patterns and layout/value bounds. Scene schema defines a closed geometry union and paint/diagnostic fields; XSD fixes Canonical/Scene/Bindings element order and binding attribute shapes. Runtime validators add uniqueness, dimensional equality, cross-reference, exact projection and hash checks unavailable in the schemas.

## 5. Classes and interfaces

Static exported functions only; no class type in these files.

## 6. Important methods and functions

Declaration locator below is lexical and may include overload/call-like lines; semantic workflow is described separately.



## 7. Inputs and outputs

JSON models or rendered XML -> structural acceptance/error from a caller validator; runtime codecs perform additional semantic/hash/ownership checks.

## 8. Dependencies

AJV2020 for JSON Schema draft2020-12; System.Xml.Schema for XSD; no Office or network retrieval is required for local schema use.

## 9. Workflow

Chart schema defines closed required object fields, version/type constants, semantic ID text patterns and layout/value bounds. Scene schema defines a closed geometry union and paint/diagnostic fields; XSD fixes Canonical/Scene/Bindings element order and binding attribute shapes. Runtime validators add uniqueness, dimensional equality, cross-reference, exact projection and hash checks unavailable in the schemas.

## 10. Integration points

JSON models or rendered XML -> structural acceptance/error from a caller validator; runtime codecs perform additional semantic/hash/ownership checks. Authored versioned schemas are consumed by AJV tests and .NET XSD tests. They are specifications, not independent services or authorization policies.

## 11. Exceptions and recovery

Schema validation fails unsupported structures. Consumer error codes/throwing behavior belong to AJV wrapper or runtime codec rather than these passive files.

## 12. Configuration dependencies

Chart version2; scene version1; rendered XML version1 and urn namespace; canonical/scene XML text length limits1MiB/8MiB and at most2000 bindings.

## 13. Security considerations

AdditionalProperties=false and whitelist unions constrain data. A structurally valid renderHash is not a verified digest; schema validation never authorizes native document mutation.

## 14. Unit and integration test definitions

- `packages/chart-domain/test/contracts.test.mjs:1`
- `packages/chart-engine/test/scene-schema.test.mjs:1`
- `tests/s3/NativeRendererTests.cs:1`

Inspected definitions only. See the parent current test report for execution; historical passes do not establish current acceptance.

## 15. Known limitations

No authored v3 JSON schema in inspected tree; XSD does not establish embedded JSON semantics or native ownership. External schema URI identifiers are identifiers, not runtime download requirements.

## 16. Technical debt

Three format specifications plus two language codecs must evolve together; compatibility cannot be established from schema alone.

## 17. Extension points

Version schemas explicitly and extend independent positive/negative/cross-language corpus before consumer rollout.
