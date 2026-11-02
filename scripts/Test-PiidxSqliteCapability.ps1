$ErrorActionPreference='Stop'
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class PpiSqliteCapabilityProbe {
 [DllImport("winsqlite3.dll", CallingConvention=CallingConvention.Cdecl)] public static extern int sqlite3_open([MarshalAs(UnmanagedType.LPUTF8Str)] string name,out IntPtr db);
 [DllImport("winsqlite3.dll", CallingConvention=CallingConvention.Cdecl)] public static extern int sqlite3_exec(IntPtr db,[MarshalAs(UnmanagedType.LPUTF8Str)] string sql,IntPtr cb,IntPtr arg,out IntPtr error);
 [DllImport("winsqlite3.dll", CallingConvention=CallingConvention.Cdecl)] public static extern IntPtr sqlite3_libversion();
 [DllImport("winsqlite3.dll", CallingConvention=CallingConvention.Cdecl)] public static extern int sqlite3_close(IntPtr db);
 [DllImport("winsqlite3.dll", CallingConvention=CallingConvention.Cdecl)] public static extern void sqlite3_free(IntPtr p);
}
"@
$db=[IntPtr]::Zero;$err=[IntPtr]::Zero
try {
 $rc=[PpiSqliteCapabilityProbe]::sqlite3_open(':memory:',[ref]$db)
 if($rc){throw "sqlite3_open failed: $rc"}
 $sql="CREATE VIRTUAL TABLE capability_fts USING fts5(title,body); INSERT INTO capability_fts VALUES('Unicode résumé','library indexing'); SELECT rowid FROM capability_fts WHERE capability_fts MATCH 'library'; PRAGMA integrity_check;"
 $rc=[PpiSqliteCapabilityProbe]::sqlite3_exec($db,$sql,[IntPtr]::Zero,[IntPtr]::Zero,[ref]$err)
 if($rc){throw "FTS5 capability failed: $rc $([Runtime.InteropServices.Marshal]::PtrToStringAnsi($err))"}
 [pscustomobject]@{Status='PASS';OS=[Environment]::OSVersion.VersionString;Process64=[Environment]::Is64BitProcess;SQLite=[Runtime.InteropServices.Marshal]::PtrToStringAnsi([PpiSqliteCapabilityProbe]::sqlite3_libversion());Provider='Windows inbox winsqlite3.dll';Evidence='In-memory FTS5 create/insert/search SQL and integrity pragma executed successfully. This is not a PIIDX or Office integration test.'} | ConvertTo-Json
}finally{if($err -ne [IntPtr]::Zero){[PpiSqliteCapabilityProbe]::sqlite3_free($err)};if($db -ne [IntPtr]::Zero){$null=[PpiSqliteCapabilityProbe]::sqlite3_close($db)}}