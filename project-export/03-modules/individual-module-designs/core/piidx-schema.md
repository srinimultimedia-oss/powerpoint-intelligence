# Versioned independent PIIDX schema

**VERIFIED source review.** This design describes inspected code, not current runtime/Office execution.

## 1. Purpose

Creates and verifies exact authoritative schema, constraints, indexes and search deletion trigger.

## 2. Implementation status

Production format1/engine0.1.0; foreign/future/nonempty unversioned DB not adopted.

## 3. Source locations

- `src/Library.Core/PiidxSchema.cs:1`

## 4. Internal components

NativeSqlite, fixed DDL, FTS5 unicode61 remove_diacritics2. For empty schema create DDL in BEGIN IMMEDIATE, singleton metadata, application_id/user_version, COMMIT, validate. Existing schema must match column order and sqlite_master SQL exactly. Deep validation runs integrity_check and foreign_key_check.

## 5. Classes and interfaces

- `src/Library.Core/PiidxSchema.cs:9` — `internal static class PiidxSchema`

## 6. Important methods and functions

Declaration locator below is lexical and may include overload/call-like lines; semantic workflow is described separately.

- `src/Library.Core/PiidxSchema.cs:42` — `internal static void Initialize(SqliteDatabase db)`
- `src/Library.Core/PiidxSchema.cs:59` — `internal static void Validate(SqliteDatabase db,bool deep=true)`

## 7. Inputs and outputs

SqliteDatabase -> initialized/validated index or preserved rejection.

## 8. Dependencies

NativeSqlite, fixed DDL, FTS5 unicode61 remove_diacritics2.

## 9. Workflow

For empty schema create DDL in BEGIN IMMEDIATE, singleton metadata, application_id/user_version, COMMIT, validate. Existing schema must match column order and sqlite_master SQL exactly. Deep validation runs integrity_check and foreign_key_check.

## 10. Integration points

SqliteDatabase -> initialized/validated index or preserved rejection. Production format1/engine0.1.0; foreign/future/nonempty unversioned DB not adopted.

## 11. Exceptions and recovery

InvalidDataException for missing/weakened/foreign/future definitions; init rolls back on any failure.

## 12. Configuration dependencies

FormatId PowerPointIntelligence.PIIDX, appId0x50494958, userVersion1; ten logical entities including FTS virtual table, three indexes and one trigger.

## 13. Security considerations

Exact DDL prevents silently weakened constraints; query constants do not interpolate caller table names. FTS shadow tables are SQLite implementation detail.

## 14. Unit and integration test definitions

- `tests/library/piidx/PiidxSchemaTests.cs:1`

Inspected definitions only. See the parent current test report for execution; historical passes do not establish current acceptance.

## 15. Known limitations

No general migration runner; additional unrelated tables not categorically forbidden when required definitions match.

## 16. Technical debt

Exact SQL equality is intentionally strict and must change only with versioned migration.

## 17. Extension points

Introduce explicit backup/migration/rollback fixtures rather than editing schema version constants alone.
