# PIIDX API reference — implemented foundation

These types are internal C# implementation components, not a public SDK contract. They are not yet included in the VSTO project or connected to the Library UI.

- SqliteDatabase.Open(path, readOnly): opens an OS SQLite connection, sets foreign keys/trusted-schema policy and busy timeout. Caller owns disposal. The adapter accepts ordinary file paths or :memory:; the higher-level PIIDX reader/writer still needs to enforce index format and lifecycle policy.
- SqliteDatabase.Execute(sql, params values): executes one authored statement with bound typed values. A second statement, incorrect parameter count, unsupported types or nonfinite numbers are rejected. Does not accept raw user query syntax as an application contract.
- SqliteDatabase.Query(sql, params values): streams object-array rows, retaining a disposable native statement during enumeration. Consume/dispose the enumerator before disposing the connection. Text/blob values are limited to 16 MiB.
- PiidxSchema.Initialize(db): initializes only an empty unversioned database, transactionally. Existing versioned files are validated; foreign nonempty databases are preserved and rejected.
- PiidxSchema.Validate(db): checks application/schema/header identity, required columns, exact authored schema definitions, SQLite integrity and foreign keys. It does not yet verify higher-level source freshness, logical content digests, PNG validity or job recovery.
- SqliteFailure.Code: numeric SQLite result. Error messages exclude SQL and bound content.

Reproduction: scripts/Test-PiidxSqliteCapability.ps1 and scripts/Test-PiidxFoundation.ps1. The latter records source hashes and runtime details under artifacts/piidx/schema-*/result.json. Files in these directories are test fixtures, not a completed sample presentation index. The negative tests intentionally create incompatible schemas.

Reader/writer/search/refresh/insertion service contracts will be documented here when implemented. No stub API is represented as operational.
## Implemented source layer

- PresentationSource.Discover(root, recursive, cancellation): streams supported source paths under explicit file/directory limits.
- PresentationSource.Fingerprint(path, cancellation): streams SHA256 without logging content.
- PresentationSource.Read(path, cancellation): returns typed IndexedPresentation/IndexedSlide/IndexedObject metadata, checks source freshness, and preserves original files. Native slide IDs are separate from positions; object keys are separate from potentially repeated native object IDs. See PIIDX_SOURCE_EXTRACTION.md for scope, bounds and remaining gaps.
- scripts/Test-PiidxSourceReader.ps1: 24 x64 assertions including all 62 local packages / 497 slides. This is extraction evidence, not a persisted-index or Office test.
