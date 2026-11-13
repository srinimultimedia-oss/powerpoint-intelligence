# Native and event contract declarations

VERIFIED source declarations; signatures are references, not a copied native implementation. No native API was executed by this inventory. Runtime ownership/thread rules are described in internal-interfaces.md and PowerPoint integration.

| Source | Declaration |
|---|---|
| `apps/powerpoint-vsto/ContextPropertiesView.cs:23` | `internal event Action<string,string,object> Command;` |
| `apps/powerpoint-vsto/ContextPropertiesView.cs:24` | `internal event Action<string> Dismissed;` |
| `apps/powerpoint-vsto/ContextSelectionBridge.cs:26` | `[DllImport("user32.dll")]private static extern bool GetWindowRect(IntPtr hwnd,out Rect rect);` |
| `apps/powerpoint-vsto/ContextSelectionBridge.cs:27` | `[DllImport("user32.dll")]private static extern IntPtr GetAncestor(IntPtr hwnd,uint flags);` |
| `apps/powerpoint-vsto/ContextSelectionBridge.cs:28` | `[DllImport("user32.dll")]internal static extern IntPtr GetForegroundWindow();` |
| `apps/powerpoint-vsto/Contracts.cs:7` | `public interface IPowerPointHost` |
| `apps/powerpoint-vsto/DemoApi.cs:12` | `public interface IDemoAutomation` |
| `apps/powerpoint-vsto/PowerPointHost.cs:18` | `[System.Runtime.InteropServices.DllImport("user32.dll",ExactSpelling=true)]` |
| `apps/powerpoint-vsto/PowerPointHost.cs:19` | `private static extern uint GetWindowThreadProcessId(IntPtr window,out uint processId);` |
| `apps/powerpoint-vsto/PowerPointHost.cs:42` | `diagnostic("Office event failed: " + ex.GetType().Name + "; HRESULT=0x" + ex.HResult.ToString("X8", CultureInfo.InvariantCulture));` |
| `src/Library.Core/NativeSqlite.cs:30` | `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_open_v2(byte[] path,out IntPtr db,int flags,IntPtr vfs);` |
| `src/Library.Core/NativeSqlite.cs:31` | `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_close_v2(IntPtr db);` |
| `src/Library.Core/NativeSqlite.cs:32` | `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_busy_timeout(SqliteHandle db,int ms);` |
| `src/Library.Core/NativeSqlite.cs:33` | `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_prepare_v2(SqliteHandle db,IntPtr sql,int length,out IntPtr statement,out IntPtr tail);` |
| `src/Library.Core/NativeSqlite.cs:34` | `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_finalize(IntPtr statement);` |
| `src/Library.Core/NativeSqlite.cs:35` | `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_step(StatementHandle statement);` |
| `src/Library.Core/NativeSqlite.cs:36` | `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern void sqlite3_interrupt(SqliteHandle db);` |
| `src/Library.Core/NativeSqlite.cs:37` | `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_get_autocommit(SqliteHandle db);` |
| `src/Library.Core/NativeSqlite.cs:38` | `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_bind_parameter_count(StatementHandle statement);` |
| `src/Library.Core/NativeSqlite.cs:39` | `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_bind_null(StatementHandle statement,int index);` |
| `src/Library.Core/NativeSqlite.cs:40` | `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_bind_int64(StatementHandle statement,int index,long value);` |
| `src/Library.Core/NativeSqlite.cs:41` | `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_bind_double(StatementHandle statement,int index,double value);` |
| `src/Library.Core/NativeSqlite.cs:42` | `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_bind_text(StatementHandle statement,int index,byte[] value,int size,IntPtr destroy);` |
| `src/Library.Core/NativeSqlite.cs:43` | `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_bind_blob(StatementHandle statement,int index,byte[] value,int size,IntPtr destroy);` |
| `src/Library.Core/NativeSqlite.cs:44` | `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_column_count(StatementHandle statement);` |
| `src/Library.Core/NativeSqlite.cs:45` | `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_column_type(StatementHandle statement,int index);` |
| `src/Library.Core/NativeSqlite.cs:46` | `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern long sqlite3_column_int64(StatementHandle statement,int index);` |
| `src/Library.Core/NativeSqlite.cs:47` | `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern double sqlite3_column_double(StatementHandle statement,int index);` |
| `src/Library.Core/NativeSqlite.cs:48` | `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern IntPtr sqlite3_column_blob(StatementHandle statement,int index);` |
| `src/Library.Core/NativeSqlite.cs:49` | `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern IntPtr sqlite3_column_text(StatementHandle statement,int index);` |
| `src/Library.Core/NativeSqlite.cs:50` | `[DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_column_bytes(StatementHandle statement,int index);` |
| `src/Chart.Renderers/NativeRenderingCore.cs:136` | `public interface IEditableSceneSurface` |
