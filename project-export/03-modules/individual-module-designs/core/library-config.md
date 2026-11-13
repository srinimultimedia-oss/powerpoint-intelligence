# Library profile and explicit legacy registration migration

**VERIFIED source review.** This design describes inspected code, not current runtime/Office execution.

## 1. Purpose

Loads whitelisted app/enterprise library settings and migrates only approved old source registration by fresh extraction.

## 2. Implementation status

Production profile/migration helper; does not import proprietary/cached slide data.

## 3. Source locations

- `src/Library.Core/LibraryConfiguration.cs:1`

## 4. Internal components

PathPolicy, secure XML, catalog and source fingerprints. Load defaults then recognized settings; optional template root and version1 enterprise XML. Legacy migration extracts folder only, matches independently configured/registered approved root, reindexes original source, verifies index and unchanged legacy hash, writes CreateNew receipt. Original legacy file retained.

## 5. Classes and interfaces

- `src/Library.Core/LibraryConfiguration.cs:13` — `internal sealed class LibrarySourceDefinition`
- `src/Library.Core/LibraryConfiguration.cs:20` — `internal sealed class LibraryRuntimeProfile`

## 6. Important methods and functions

Declaration locator below is lexical and may include overload/call-like lines; semantic workflow is described separately.

- `src/Library.Core/LibraryConfiguration.cs:32` — `internal static LibraryRuntimeProfile Load(string configPath,string baseDirectory,string localAppData)`
- `src/Library.Core/LibraryConfiguration.cs:55` — `internal static IReadOnlyList<LibrarySourceDefinition> ParseSources(string text,string baseDirectory)`
- `src/Library.Core/LibraryConfiguration.cs:70` — `internal static string LegacyFolder(string text)`
- `src/Library.Core/LibraryConfiguration.cs:79` — `internal string MigrateLegacy(PiidxCatalog catalog,CancellationToken cancellation)`

## 7. Inputs and outputs

Config path/baseDirectory/LocalAppData -> LibraryRuntimeProfile; catalog+cancel -> migrated ID or null when receipt exists.

## 8. Dependencies

PathPolicy, secure XML, catalog and source fingerprints.

## 9. Workflow

Load defaults then recognized settings; optional template root and version1 enterprise XML. Legacy migration extracts folder only, matches independently configured/registered approved root, reindexes original source, verifies index and unchanged legacy hash, writes CreateNew receipt. Original legacy file retained.

## 10. Integration points

Config path/baseDirectory/LocalAppData -> LibraryRuntimeProfile; catalog+cancel -> migrated ID or null when receipt exists. Production profile/migration helper; does not import proprietary/cached slide data.

## 11. Exceptions and recovery

Unknown/duplicate keys/fields/schema, bad limits, absent enterprise file or unapproved root reject; no receipt on race.

## 12. Configuration dependencies

Default LocalAppData/PowerPointIntelligence/Library/catalog.piidx; poll 60 seconds / watch 32 roots; interval5..86400 and watch0..128; enterprise <= 1,000 roots / 4 M characters.

## 13. Security considerations

No DTD/network resolver; index/config owned ancestors reject links; legacy text cannot grant new root authority. Managed flag is metadata, not OS/RBAC enforcement.

## 14. Unit and integration test definitions

- `tests/library/piidx/PiidxWorkspaceTests.cs:1`

Inspected definitions only. See the parent current test report for execution; historical passes do not establish current acceptance.

## 15. Known limitations

No central policy server/identity roles; cached old text is deliberately ignored.

## 16. Technical debt

Configuration/receipt backup and version migration must be operationally maintained.

## 17. Extension points

Add fields through whitelist/schema/tests; preserve selected-input versus owned-output policy.
