# Canonical XML ownership and compensating replacement

**VERIFIED — source implementation. Current runtime acceptance is not established by this design review.**

## Purpose and implementation status

Update shape/deck XML without trusting stale COM namespace proxies or losing shared copied owners.

Implemented local source is distinguished from supported subset/absent external behavior below. See current test evidence in the export testing chapter; do not treat historical logs or method existence as a current runtime pass.

## Source locations and internal components

- `apps/powerpoint-vsto/XmlPartUpdates.cs:1` — XmlPartUpdates.

## Inputs, outputs, data structures and dependencies

Inputs: live owner, previous CustomXMLPart and new XML; output: verified current part or explicit compensating failure. Dependencies: Office CustomerData/CustomXMLPart and System.Xml; shared by managed text/grid/agenda/connectors.

## Internal execution and external integration

Validate size/DTD first. AttachedParts normalizes GUIDs, resolves every unique part through live owner collection and extracts namespace from serialized XML. Equivalent compares namespaced nodes/attributes/text while respecting significant whitespace. Shape preflights the canonical attachment, scans shared owners, detaches/adds/verifies target and restores copied owners. Deck requires one current registry, replaces/verifies it and restores on failure. ReplaceAttachment aggregates cleanup/restoration exceptions.

## Configuration, validation and failure behavior

XML<=1000000 chars,<=64 parts/shape,owner scan<=10000 objects/depth8,XMLcomparisonancestor depth<=128. Namespace cannot change via replacement. Messages report phase/counts, not XML content.

## Security considerations and cleanup

Local native calls run in the current Office user process and require supported selection/ownership; a mutex/selection key is not authentication or RBAC. File/metadata content must remain bounded and error text must not expose data or credentials. Explicit owned-resource cleanup and error/rollback paths are part of the cited implementation, not a guarantee that a failed COM server can always be restored.

## Test references

Lexical module-class matches in test source (not instrumentation coverage; no test is claimed executed here):

- `tests/s0/UniversalConnectorOfficeTests.cs:39`.
- `tests/s0/XmlPartUpdateTests.cs:10`.

## Known limitations and technical debt

COM compensation is not an atomic OOXML package transaction. Supported shape XML update requires ungrouped owner. Copied shared-part/restoration quirks require native regression tests; incomplete restore directs Undo.

## Extension points

Use helper for existing mutable parts rather than blind LoadXML; add positive/negative semantic-equivalence and shared-owner native tests.

## Classes, interfaces and method navigation

Method names below are source declarations; private/internal members are not public API promises. The source hash/review manifest gives exact reviewed bytes.

| Declaration | Kind | Source |
|---|---|---|
| `XmlPartUpdates` | class | `apps/powerpoint-vsto/XmlPartUpdates.cs:11` |
| `Validate` | method | `apps/powerpoint-vsto/XmlPartUpdates.cs:13` |
| `Add` | method | `apps/powerpoint-vsto/XmlPartUpdates.cs:15` |
| `SamePartId` | method | `apps/powerpoint-vsto/XmlPartUpdates.cs:17` |
| `Content` | method | `apps/powerpoint-vsto/XmlPartUpdates.cs:19` |
| `AttributeKey` | method | `apps/powerpoint-vsto/XmlPartUpdates.cs:25` |
| `Equivalent` | method | `apps/powerpoint-vsto/XmlPartUpdates.cs:26` |
| `Has` | method | `apps/powerpoint-vsto/XmlPartUpdates.cs:35` |
| `AttachedParts` | method | `apps/powerpoint-vsto/XmlPartUpdates.cs:47` |
| `RootNamespace` | method | `apps/powerpoint-vsto/XmlPartUpdates.cs:52` |
| `VerifyAttachment` | method | `apps/powerpoint-vsto/XmlPartUpdates.cs:59` |
| `Shape` | method | `apps/powerpoint-vsto/XmlPartUpdates.cs:63` |
| `Deck` | method | `apps/powerpoint-vsto/XmlPartUpdates.cs:76` |
