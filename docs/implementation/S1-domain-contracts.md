# S1: domain and contracts — first slice

The owner explicitly accepted S0 in chat on 21 October 2026 and then requested S1. This authorizes progression; it does not establish any missing Windows runtime evidence. S0 AT-001–005 remain unverified as recorded. No PR was merged by that acceptance.

## Scope reconciliation

The controlling DOCX section 13 defines S1 as JSON Schema, DTOs, semantic IDs, versioning, validation and migrations, with golden serialization round trips as its exit evidence. FEATURES.md also mentions geometry and AGENTS.md mentions command bus/rendering. Follow the controlling document's sequence: calculations are S2, layout/rendering S3, and command execution S4. This first S1 slice implements the embedded clustered-column domain profile; it does not assert the entire canonical model or S1 milestone complete.

## Contract

- JSON version 2 (`urn:powerpoint-intelligence:chart-document:v2`) is distinct from S0 XML version 1. Never reinterpret an old XML payload as a new JSON document.
- Pure .NET Framework 4.8 project under src/Chart.Domain, with pinned Newtonsoft.Json 13.0.3; no COM, Office, service, database or AI dependency. Legacy DTO/codec source is linked for explicit import, not copied into a competing S0 wire implementation. The VSTO project continues to use its existing S0 DTOs and is not switched to v2 by this change.
- The JSON Schema and .NET DTOs cover embedded category labels/IDs, ordered series and nullable data values, exactly one category and one linear value axis, layout in points, explicit legend visibility and per-series RGB colors.
- Chart IDs are nonempty GUIDs. Child IDs are stored semantic identities, unique within a document. Changing a label or array position must not regenerate an existing v2 identity. IDs are case-sensitive child identifiers; the chart GUID remains a GUID string. No native shape ID is used.
- Shape validation rejects missing/unknown members and type coercion. Semantic validation checks category/value dimensions, unique IDs, axis roles/scales, finite numbers, positive bounded sizes and matching color counts. Numeric bounds/counts are defensive limits for this profile, not an approved production performance envelope.
- Canonical encoding uses declared property order, invariant culture and preserved list order; decoding/encoding again is byte-stable within this codec. It is not a general JSON canonicalization standard or an audit signature format.
- Unknown/future versions are rejected without writing files. No automatic down-conversion. Annotations, overrides, linked sources/provenance and additional chart families are intentionally rejected until their typed contracts and tests are implemented in further slices. Thus this is a bounded profile rather than the complete section 7 model.

## Migration and persistence boundary

`S0Migration.FromXml(xml, layout, style)` validates legacy XML before import. It preserves chart ID, revision, label/series order, names, signed values and nulls. Category and series identities are deterministic chart-ID-plus-ordinal IDs on this one-time import; afterwards they are stored identities and must not be recomputed on reordering. Legacy series role IDs are not copied into the v2 ID field because S0 allowed strings outside the new semantic-ID grammar. No claims are made about preserving unmodeled S0 fields.

S0 did not serialize layout/formatting, so migration requires explicit layout and style inputs. It refuses missing inputs and never guesses corporate colors or resets native formatting. Import clones all data/defaults and does not mutate XML or caller-owned objects. A repeated import of identical XML and defaults yields identical JSON.

The codec is an in-memory persistence boundary only. Attaching v2 JSON to PPTX, automatic migration on reopen, native chart projection, host selection, rollback, and process round-trip remain separate host integration work and are not verified here. No open presentation ZIP is modified.

## Checks

1. `npm ci --ignore-scripts` then `npm test` in packages/chart-domain.
2. `./scripts/Test-S1.ps1` on a machine permitted to execute compiled tests; `-Compiler <csc.exe>` selects Roslyn explicitly. `-CompileOnly` reports compilation only.
3. `MSBuild src/Chart.Domain/Chart.Domain.csproj /restore /t:Build` with .NET Framework 4.8 targeting pack and current MSBuild.
4. The Windows S1 workflow runs C# tests/build, schema tests, and validates actual C# output against the published schema and JS semantic validator.
5. Keep S0 metadata and existing chart-core workflows passing. Record results in docs/evidence/S1-2026-10-21.md.

No CHT-001, DAT-001, EDT-012 or S0 host acceptance criterion is marked passed by domain-only tests.
