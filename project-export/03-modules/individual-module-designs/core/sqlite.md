# Windows native SQLite bridge

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

Parameterized single-statement bridge to OS SQLite with deterministic resource cleanup.

## 2. Implementation status

Production in-process source, Windows winsqlite3 capability required; no managed database server.

## 3. Source locations

- `src/Library.Core/NativeSqlite.cs:1`

## 4. Internal components

System32 winsqlite3.dll P/Invoke, SafeHandle, strictUTF8, CancellationToken. Open flags read-only1 or read-write/create6 plus FULLMUTEX; foreign_keys ON/trusted_schema OFF, busy5s. Pin SQL, prepare one statement, reject trailing SQL/parameter mismatch, bind transient buffers, step rows and release statement/connection handles. Cancellation interrupts native query.

## 5. Classes and interfaces

- `src/Library.Core/NativeSqlite.cs:11` — `internal sealed class SqliteFailure : IOException`
- `src/Library.Core/NativeSqlite.cs:16` — `internal sealed class SqliteHandle : SafeHandleZeroOrMinusOneIsInvalid`
- `src/Library.Core/NativeSqlite.cs:21` — `internal sealed class StatementHandle : SafeHandleZeroOrMinusOneIsInvalid`
- `src/Library.Core/NativeSqlite.cs:27` — `internal static class SqliteNative`
- `src/Library.Core/NativeSqlite.cs:52` — `internal sealed class SqliteDatabase : IDisposable`
- `src/Library.Core/NativeSqlite.cs:59` — `sealed class CancellationScope : IDisposable`

## 6. Important methods and functions

Declaration locator below is lexical and may include overload/call-like lines; semantic workflow is described separately.

- `src/Library.Core/NativeSqlite.cs:14` — `internal SqliteFailure(int code):base("PIIDX SQLite operation failed (code="+code+"). Query text and bound content are omitted.")`
- `src/Library.Core/NativeSqlite.cs:18` — `internal SqliteHandle(IntPtr value):base(true)`
- `src/Library.Core/NativeSqlite.cs:23` — `internal StatementHandle(IntPtr value):base(true)`
- `src/Library.Core/NativeSqlite.cs:30` — `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_open_v2(byte[] path,out IntPtr db,int flags,IntPtr vfs);`
- `src/Library.Core/NativeSqlite.cs:31` — `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_close_v2(IntPtr db);`
- `src/Library.Core/NativeSqlite.cs:32` — `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_busy_timeout(SqliteHandle db,int ms);`
- `src/Library.Core/NativeSqlite.cs:33` — `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_prepare_v2(SqliteHandle db,IntPtr sql,int length,out IntPtr statement,out IntPtr tail);`
- `src/Library.Core/NativeSqlite.cs:34` — `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_finalize(IntPtr statement);`
- `src/Library.Core/NativeSqlite.cs:35` — `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_step(StatementHandle statement);`
- `src/Library.Core/NativeSqlite.cs:36` — `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern void sqlite3_interrupt(SqliteHandle db);`
- `src/Library.Core/NativeSqlite.cs:37` — `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_get_autocommit(SqliteHandle db);`
- `src/Library.Core/NativeSqlite.cs:38` — `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_bind_parameter_count(StatementHandle statement);`
- `src/Library.Core/NativeSqlite.cs:39` — `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_bind_null(StatementHandle statement,int index);`
- `src/Library.Core/NativeSqlite.cs:40` — `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_bind_int64(StatementHandle statement,int index,long value);`
- `src/Library.Core/NativeSqlite.cs:41` — `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_bind_double(StatementHandle statement,int index,double value);`
- `src/Library.Core/NativeSqlite.cs:42` — `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_bind_text(StatementHandle statement,int index,byte[] value,int size,IntPtr destroy);`
- `src/Library.Core/NativeSqlite.cs:43` — `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_bind_blob(StatementHandle statement,int index,byte[] value,int size,IntPtr destroy);`
- `src/Library.Core/NativeSqlite.cs:44` — `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_column_count(StatementHandle statement);`
- `src/Library.Core/NativeSqlite.cs:45` — `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_column_type(StatementHandle statement,int index);`
- `src/Library.Core/NativeSqlite.cs:46` — `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern long sqlite3_column_int64(StatementHandle statement,int index);`
- `src/Library.Core/NativeSqlite.cs:47` — `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern double sqlite3_column_double(StatementHandle statement,int index);`
- `src/Library.Core/NativeSqlite.cs:48` — `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern IntPtr sqlite3_column_blob(StatementHandle statement,int index);`
- `src/Library.Core/NativeSqlite.cs:49` — `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern IntPtr sqlite3_column_text(StatementHandle statement,int index);`
- `src/Library.Core/NativeSqlite.cs:50` — `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_column_bytes(StatementHandle statement,int index);`
- `src/Library.Core/NativeSqlite.cs:62` — `internal CancellationScope(SqliteDatabase owner,CancellationTokenRegistration registration)`
- `src/Library.Core/NativeSqlite.cs:63` — `public void Dispose()`
- `src/Library.Core/NativeSqlite.cs:65` — `internal IDisposable ObserveCancellation(CancellationToken token)`
- `src/Library.Core/NativeSqlite.cs:72` — `internal void RollbackIfActive()`
- `src/Library.Core/NativeSqlite.cs:75` — `internal static SqliteDatabase Open(string path,bool readOnly)`
- `src/Library.Core/NativeSqlite.cs:127` — `internal void Execute(string sql,params object[] values)`
- `src/Library.Core/NativeSqlite.cs:135` — `internal IEnumerable<object[]> Query(string sql,params object[] values)`
- `src/Library.Core/NativeSqlite.cs:158` — `public void Dispose()`

## 7. Inputs and outputs

Index path/readOnly+SQL+bound values -> rows or completion. Supported values null/string/blob/int64/int/bool/double.

## 8. Dependencies

System32 winsqlite3.dll P/Invoke, SafeHandle, strictUTF8, CancellationToken.

## 9. Workflow

Open flags read-only1 or read-write/create6 plus FULLMUTEX; foreign_keys ON/trusted_schema OFF, busy5s. Pin SQL, prepare one statement, reject trailing SQL/parameter mismatch, bind transient buffers, step rows and release statement/connection handles. Cancellation interrupts native query.

## 10. Integration points

Index path/readOnly+SQL+bound values -> rows or completion. Supported values null/string/blob/int64/int/bool/double. Production in-process source, Windows winsqlite3 capability required; no managed database server.

## 11. Exceptions and recovery

SqliteFailure retains code and omits query/content; interrupt maps to OperationCanceledException when active token cancelled. Unsupported values/nonfinite doubles rejected.

## 12. Configuration dependencies

Maximum bound/read value16MiB; SQLite busy timeout5000ms; no DLL search in content directories.

## 13. Security considerations

Parameterized SQL and single-statement boundary; System32-only DLL imports. Local index is not encrypted and relies on OS access rights.

## 14. Unit and integration test definitions

- `tests/library/piidx/PiidxSchemaTests.cs:1`
- `scripts/Test-PiidxSqliteCapability.ps1:1`

Inspected definitions only. See the parent current test report for execution; historical passes do not establish current acceptance.

## 15. Known limitations

OS FTS5/provider variability, no native fallback. Large result sets may be materialized by callers.

## 16. Technical debt

No native SQLite version pin independent of Windows; architecture/provider acceptance needed.

## 17. Extension points

Add binding types with limits/Unicode/blob/rollback tests; retain safe handles and native cancellation.
