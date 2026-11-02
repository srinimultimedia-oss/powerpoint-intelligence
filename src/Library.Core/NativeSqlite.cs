using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Microsoft.Win32.SafeHandles;

namespace PowerPointIntelligence.Library
{
    internal sealed class SqliteFailure : IOException
    {
        internal readonly int Code;
        internal SqliteFailure(int code):base("PIIDX SQLite operation failed (code="+code+"). Query text and bound content are omitted."){Code=code;}
    }
    internal sealed class SqliteHandle : SafeHandleZeroOrMinusOneIsInvalid
    {
        internal SqliteHandle(IntPtr value):base(true){SetHandle(value);}
        protected override bool ReleaseHandle(){return SqliteNative.sqlite3_close_v2(handle)==0;}
    }
    internal sealed class StatementHandle : SafeHandleZeroOrMinusOneIsInvalid
    {
        internal StatementHandle(IntPtr value):base(true){SetHandle(value);}
        protected override bool ReleaseHandle(){SqliteNative.sqlite3_finalize(handle);return true;}
    }
    // Restrict loading to the OS provider: never probe a library folder for executable code.
    internal static class SqliteNative
    {
        const string Dll="winsqlite3.dll";
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_open_v2(byte[] path,out IntPtr db,int flags,IntPtr vfs);
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_close_v2(IntPtr db);
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_busy_timeout(SqliteHandle db,int ms);
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_prepare_v2(SqliteHandle db,IntPtr sql,int length,out IntPtr statement,out IntPtr tail);
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_finalize(IntPtr statement);
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_step(StatementHandle statement);
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern void sqlite3_interrupt(SqliteHandle db);
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_get_autocommit(SqliteHandle db);
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_bind_parameter_count(StatementHandle statement);
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_bind_null(StatementHandle statement,int index);
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_bind_int64(StatementHandle statement,int index,long value);
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_bind_double(StatementHandle statement,int index,double value);
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_bind_text(StatementHandle statement,int index,byte[] value,int size,IntPtr destroy);
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_bind_blob(StatementHandle statement,int index,byte[] value,int size,IntPtr destroy);
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_column_count(StatementHandle statement);
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_column_type(StatementHandle statement,int index);
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern long sqlite3_column_int64(StatementHandle statement,int index);
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern double sqlite3_column_double(StatementHandle statement,int index);
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern IntPtr sqlite3_column_blob(StatementHandle statement,int index);
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern IntPtr sqlite3_column_text(StatementHandle statement,int index);
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32),DllImport(Dll,CallingConvention=CallingConvention.Cdecl)] internal static extern int sqlite3_column_bytes(StatementHandle statement,int index);
    }
    internal sealed class SqliteDatabase : IDisposable
    {
        const int MaximumValueBytes=16*1024*1024;
        static readonly UTF8Encoding Utf8=new UTF8Encoding(false,true);
        SqliteHandle handle;
        CancellationToken activeCancellation;
        bool observingCancellation;
        sealed class CancellationScope : IDisposable
        {
            readonly SqliteDatabase owner;readonly CancellationTokenRegistration registration;
            internal CancellationScope(SqliteDatabase owner,CancellationTokenRegistration registration){this.owner=owner;this.registration=registration;}
            public void Dispose(){registration.Dispose();owner.activeCancellation=CancellationToken.None;owner.observingCancellation=false;}
        }
        internal IDisposable ObserveCancellation(CancellationToken token)
        {
            token.ThrowIfCancellationRequested();if(observingCancellation)throw new InvalidOperationException("A connection already observes cancellation.");
            activeCancellation=token;observingCancellation=true;
            var registration=token.Register(()=>{var value=handle;if(value!=null&&!value.IsClosed)try{SqliteNative.sqlite3_interrupt(value);}catch(ObjectDisposedException){}});
            return new CancellationScope(this,registration);
        }
        internal void RollbackIfActive(){if(handle!=null&&!handle.IsClosed&&SqliteNative.sqlite3_get_autocommit(handle)==0)Execute("ROLLBACK");}
        Exception Failure(int code){return code==9&&activeCancellation.IsCancellationRequested?(Exception)new OperationCanceledException(activeCancellation):new SqliteFailure(code);}
        SqliteDatabase(SqliteHandle handle){this.handle=handle;}
        internal static SqliteDatabase Open(string path,bool readOnly)
        {
            if(String.IsNullOrWhiteSpace(path))throw new ArgumentException("Index path is required.");
            if(path!=":memory:")path=Path.GetFullPath(path);
            IntPtr raw;
            int rc=SqliteNative.sqlite3_open_v2(Bytes(path),out raw,(readOnly?1:6)|0x10000,IntPtr.Zero);
            var value=new SqliteHandle(raw);
            if(rc!=0){value.Dispose();throw new SqliteFailure(rc);}
            var db=new SqliteDatabase(value);
            try{db.Check(SqliteNative.sqlite3_busy_timeout(value,5000));db.Execute("PRAGMA foreign_keys=ON");db.Execute("PRAGMA trusted_schema=OFF");return db;}
            catch{db.Dispose();throw;}
        }
        static byte[] Bytes(string value){return Utf8.GetBytes(value+"\0");}
        void Check(int rc){if(rc!=0)throw Failure(rc);}
        StatementHandle Prepare(string sql,object[] values)
        {
            if(handle==null||handle.IsClosed)throw new ObjectDisposedException("SqliteDatabase");
            if(String.IsNullOrWhiteSpace(sql)||sql.IndexOf('\0')>=0)throw new ArgumentException("SQL statement is invalid.");
            IntPtr raw=IntPtr.Zero,tail;var data=Bytes(sql);
            var pinned=GCHandle.Alloc(data,GCHandleType.Pinned);
            try
            {
                IntPtr start=pinned.AddrOfPinnedObject();
                Check(SqliteNative.sqlite3_prepare_v2(handle,start,data.Length,out raw,out tail));
                int used=IntPtr.Size==4?unchecked(tail.ToInt32()-start.ToInt32()):checked((int)(tail.ToInt64()-start.ToInt64()));
                if(used<0||used>=data.Length)throw new InvalidDataException("SQLite returned an invalid statement boundary.");
                string remainder=Utf8.GetString(data,used,data.Length-used-1);
                if(!String.IsNullOrWhiteSpace(remainder))throw new ArgumentException("Only one SQL statement is permitted per call.");
            }
            catch{if(raw!=IntPtr.Zero)SqliteNative.sqlite3_finalize(raw);throw;}
            finally{pinned.Free();}
            var statement=new StatementHandle(raw);
            try
            {
                if(statement.IsInvalid)throw new ArgumentException("SQL statement is empty.");
                if(SqliteNative.sqlite3_bind_parameter_count(statement)!=values.Length)throw new ArgumentException("SQL parameter count does not match.");
                for(int i=0;i<values.Length;i++)Bind(statement,i+1,values[i]);
                return statement;
            }
            catch{statement.Dispose();throw;}
        }
        void Bind(StatementHandle s,int index,object value)
        {
            int rc;
            if(value==null)rc=SqliteNative.sqlite3_bind_null(s,index);
            else if(value is string){byte[] bytes=Bytes((string)value);if(bytes.Length>MaximumValueBytes)throw new ArgumentException("Bound text exceeds limits.");rc=SqliteNative.sqlite3_bind_text(s,index,bytes,bytes.Length-1,new IntPtr(-1));}
            else if(value is byte[]){var bytes=(byte[])value;if(bytes.Length>MaximumValueBytes)throw new ArgumentException("Bound blob exceeds limits.");rc=SqliteNative.sqlite3_bind_blob(s,index,bytes.Length==0?new byte[1]:bytes,bytes.Length,new IntPtr(-1));}
            else if(value is long||value is int||value is bool)rc=SqliteNative.sqlite3_bind_int64(s,index,Convert.ToInt64(value));
            else if(value is double){double number=(double)value;if(Double.IsNaN(number)||Double.IsInfinity(number))throw new ArgumentException("Nonfinite numbers are not persisted.");rc=SqliteNative.sqlite3_bind_double(s,index,number);}
            else throw new ArgumentException("Unsupported SQLite bound value type.");
            Check(rc);
        }
        internal void Execute(string sql,params object[] values)
        {
            using(var statement=Prepare(sql,values))
            {
                int rc;while((rc=SqliteNative.sqlite3_step(statement))==100){}
                if(rc!=101)throw Failure(rc);
            }
        }
        internal IEnumerable<object[]> Query(string sql,params object[] values)
        {
            using(var statement=Prepare(sql,values))
            {
                activeCancellation.ThrowIfCancellationRequested();
                int rc;while((rc=SqliteNative.sqlite3_step(statement))==100)
                {
                    activeCancellation.ThrowIfCancellationRequested();
                    var row=new object[SqliteNative.sqlite3_column_count(statement)];
                    for(int i=0;i<row.Length;i++)
                    {
                        int type=SqliteNative.sqlite3_column_type(statement,i);
                        if(type==5)row[i]=null;
                        else if(type==1)row[i]=SqliteNative.sqlite3_column_int64(statement,i);
                        else if(type==2)row[i]=SqliteNative.sqlite3_column_double(statement,i);
                        else if(type==3||type==4){IntPtr data=type==3?SqliteNative.sqlite3_column_text(statement,i):SqliteNative.sqlite3_column_blob(statement,i);int length=SqliteNative.sqlite3_column_bytes(statement,i);if(length<0||length>MaximumValueBytes)throw new InvalidDataException("Indexed value exceeds limits.");byte[] bytes=new byte[length];if(length>0)Marshal.Copy(data,bytes,0,length);row[i]=type==3?(object)Utf8.GetString(bytes):bytes;}
                        else throw new InvalidDataException("Unsupported SQLite storage type.");
                    }
                    yield return row;
                }
                if(rc!=101)throw Failure(rc);
            }
        }
        public void Dispose(){if(handle!=null){handle.Dispose();handle=null;}}
    }
}