using PowerPointIntelligence.Security;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace PowerPointIntelligence.Library
{
    internal sealed partial class PiidxCatalog
    {
        internal readonly string IndexPath;
        static readonly CultureInfo Invariant=CultureInfo.InvariantCulture;
        PiidxCatalog(string path){IndexPath=PathPolicy.Configured(path,Environment.CurrentDirectory);PathPolicy.RejectLinkedAncestors(IndexPath);if(!String.Equals(Path.GetExtension(path),".piidx",StringComparison.OrdinalIgnoreCase))throw new ArgumentException("A PIIDX filename is required.");}
        internal static PiidxCatalog Open(string path)
        {
            var catalog=new PiidxCatalog(path);if(!File.Exists(catalog.IndexPath))throw new FileNotFoundException("PIIDX index is missing.",catalog.IndexPath);
            using(var db=catalog.Connect(true,true)){db.Execute("BEGIN");try{VerifyContent(db);db.Execute("COMMIT");}catch{db.RollbackIfActive();throw;}}
            return catalog;
        }
        internal static PiidxCatalog OpenOrCreate(string path)
        {
            if(File.Exists(path))return Open(path);
            var catalog=new PiidxCatalog(path);Directory.CreateDirectory(Path.GetDirectoryName(catalog.IndexPath));
            using(catalog.WriterLock(CancellationToken.None))
            using(var db=SqliteDatabase.Open(catalog.IndexPath,false)){bool existing=db.Query("SELECT name FROM sqlite_master WHERE name NOT LIKE 'sqlite_%'").Any();PiidxSchema.Initialize(db);if(existing)VerifyContent(db);else catalog.UpdateDigest(db,false);}
            return Open(catalog.IndexPath);
        }
        SqliteDatabase Connect(bool readOnly,bool deep=false)
        {
            if(!File.Exists(IndexPath))throw new FileNotFoundException("PIIDX index is missing.",IndexPath);
            if((File.GetAttributes(IndexPath)&FileAttributes.ReparsePoint)!=0)throw new InvalidDataException("PIIDX index cannot be a reparse point.");
            var db=SqliteDatabase.Open(IndexPath,readOnly);try{PiidxSchema.Validate(db,deep);return db;}catch{db.Dispose();throw;}
        }
        FileStream WriterLock(CancellationToken cancellation)
        {
            string path=PathPolicy.Within(Path.GetDirectoryName(IndexPath),Path.GetFileName(IndexPath)+".writer.lock");PathPolicy.RejectLinkedAncestors(path);DateTime deadline=DateTime.UtcNow.AddSeconds(5);
            while(true)
            {
                cancellation.ThrowIfCancellationRequested();
                if(File.Exists(path)&&(File.GetAttributes(path)&FileAttributes.ReparsePoint)!=0)throw new InvalidDataException("Index writer lock cannot be a reparse point.");
                try{return new FileStream(path,FileMode.OpenOrCreate,FileAccess.ReadWrite,FileShare.None);}
                catch(IOException ex){int code=ex.HResult&65535;if((code!=32&&code!=33)||DateTime.UtcNow>=deadline)throw;if(cancellation.WaitHandle.WaitOne(50))cancellation.ThrowIfCancellationRequested();}
            }
        }
        static string Utc(){return DateTime.UtcNow.ToString("o",Invariant);}
        static DateTime Date(string value){DateTime parsed;if(!DateTime.TryParseExact(value,"o",Invariant,DateTimeStyles.RoundtripKind,out parsed))throw new InvalidDataException("Index timestamp is invalid.");return parsed.ToUniversalTime();}
        static void Id(string value){Guid parsed;if(!Guid.TryParseExact(value,"D",out parsed)||parsed==Guid.Empty)throw new ArgumentException("A canonical nonempty catalog ID is required.");}
        static string Root(string path){string full=PathPolicy.Configured(path,Environment.CurrentDirectory);return full.Length==Path.GetPathRoot(full).Length?full:full.TrimEnd(Path.DirectorySeparatorChar,Path.AltDirectorySeparatorChar);}
        static string Relative(string root,string file){string prefix=root.TrimEnd(Path.DirectorySeparatorChar,Path.AltDirectorySeparatorChar)+Path.DirectorySeparatorChar;file=Path.GetFullPath(file);if(!file.StartsWith(prefix,StringComparison.OrdinalIgnoreCase))throw new InvalidDataException("Source escaped its registered library.");return file.Substring(prefix.Length).Replace('\\','/');}
        internal IReadOnlyList<LibraryRegistration> Libraries()
        {
            using(var db=Connect(true))return db.Query("SELECT library_id,root_path,name,recursive,enabled,last_refresh_utc FROM LIBRARIES ORDER BY name,library_id").Select(r=>new LibraryRegistration{Id=(string)r[0],Root=(string)r[1],Name=(string)r[2],Recursive=(long)r[3]!=0,Enabled=(long)r[4]!=0,LastRefreshUtc=r[5]==null?(DateTime?)null:Date((string)r[5])}).ToList();
        }
        internal string Register(string root,string name,bool recursive)
        {
            root=Root(root);if(!Directory.Exists(root))throw new DirectoryNotFoundException("Library root is unavailable.");
            if((File.GetAttributes(root)&FileAttributes.ReparsePoint)!=0)throw new InvalidDataException("Register a direct library directory.");
            if(String.IsNullOrWhiteSpace(name)||name.Length>200)throw new ArgumentException("Library name is required and must not exceed 200 characters.");
            using(WriterLock(CancellationToken.None))using(var db=Connect(false))
            {
                VerifyContent(db);db.Execute("BEGIN IMMEDIATE");
                try
                {
                    var existing=db.Query("SELECT library_id,root_path FROM LIBRARIES").FirstOrDefault(r=>String.Equals((string)r[1],root,StringComparison.OrdinalIgnoreCase));
                    string id=existing==null?Guid.NewGuid().ToString("D"):(string)existing[0];
                    if(existing==null){if((long)db.Query("SELECT count(*) FROM LIBRARIES").Single()[0]>=1000)throw new InvalidDataException("Registered-library limit exceeded.");db.Execute("INSERT INTO LIBRARIES(library_id,root_path,name,recursive,registered_utc) VALUES(?,?,?,?,?)",id,root,name,recursive,Utc());}
                    else db.Execute("UPDATE LIBRARIES SET name=?,recursive=?,enabled=1 WHERE library_id=?",name,recursive,id);
                    UpdateDigest(db,false);db.Execute("COMMIT");return id;
                }
                catch{db.RollbackIfActive();throw;}
            }
        }
        internal void Remove(string libraryId)
        {
            Id(libraryId);using(WriterLock(CancellationToken.None))using(var db=Connect(false))
            {
                VerifyContent(db);db.Execute("BEGIN IMMEDIATE");try{db.Execute("DELETE FROM LIBRARIES WHERE library_id=?",libraryId);UpdateDigest(db,false);db.Execute("COMMIT");}catch{db.RollbackIfActive();throw;}
            }
        }
        internal void VerifyIntegrity(){using(var db=Connect(true,true)){db.Execute("BEGIN");try{VerifyContent(db);db.Execute("COMMIT");}catch{db.RollbackIfActive();throw;}}}
        static readonly string[] DigestQueries={
            "SELECT * FROM LIBRARIES ORDER BY library_id",
            "SELECT * FROM PRESENTATIONS ORDER BY presentation_id",
            "SELECT * FROM SLIDES ORDER BY slide_id",
            "SELECT * FROM SLIDE_OBJECTS ORDER BY slide_id,object_id",
            "SELECT * FROM SLIDE_TAGS ORDER BY slide_id,normalized_tag",
            "SELECT slide_id,title,text,keywords FROM SLIDE_SEARCH ORDER BY slide_id"
        };
        static string Digest(SqliteDatabase db,IEnumerable<string> queries)
        {
            using(var hash=SHA256.Create())using(var crypto=new CryptoStream(Stream.Null,hash,CryptoStreamMode.Write))
            {
                using(var writer=new BinaryWriter(crypto,new UTF8Encoding(false,true),true))
                {
                    foreach(string query in queries)
                    {
                        writer.Write(query);foreach(var row in db.Query(query)){writer.Write((byte)1);writer.Write(row.Length);foreach(var value in row){if(value==null)writer.Write((byte)0);else if(value is string){writer.Write((byte)1);writer.Write((string)value);}else if(value is long){writer.Write((byte)2);writer.Write((long)value);}else if(value is double){writer.Write((byte)3);writer.Write((double)value);}else throw new InvalidDataException("Unsupported catalog digest value.");}}writer.Write((byte)0);
                    }
                    writer.Flush();
                }
                crypto.FlushFinalBlock();return BitConverter.ToString(hash.Hash).Replace("-","");
            }
        }
        static void VerifyContent(SqliteDatabase db)
        {
            var row=db.Query("SELECT integrity_digest,source_fingerprint,library_id FROM INDEX_METADATA WHERE singleton=1").Single();
            var libraries=db.Query("SELECT library_id FROM LIBRARIES ORDER BY library_id").ToArray();
            string expectedLibrary=libraries.Length==1?(string)libraries[0][0]:null;
            if(!Object.Equals(row[2],expectedLibrary))throw new InvalidDataException("PIIDX header library identity is invalid.");
            if((long)db.Query("SELECT count(*) FROM SLIDES").Single()[0]!=(long)db.Query("SELECT count(*) FROM SLIDE_SEARCH").Single()[0]||db.Query("SELECT 1 FROM SLIDE_SEARCH f LEFT JOIN SLIDES s ON s.slide_id=f.slide_id WHERE s.slide_id IS NULL OR f.title<>s.title OR f.text<>s.text OR f.keywords<>s.keywords LIMIT 1").Any()||db.Query("SELECT slide_id FROM SLIDE_SEARCH GROUP BY slide_id HAVING count(*)<>1 LIMIT 1").Any())throw new InvalidDataException("PIIDX search content is inconsistent with slide metadata.");
            if(row[0]==null||!String.Equals((string)row[0],Digest(db,DigestQueries),StringComparison.Ordinal)||!String.Equals((string)row[1],Digest(db,new[]{"SELECT presentation_id,source_path,content_hash,index_status FROM PRESENTATIONS ORDER BY presentation_id"}),StringComparison.Ordinal))throw new InvalidDataException("PIIDX logical content fingerprint failed; previous index is not replaced.");
        }
        void UpdateDigest(SqliteDatabase db,bool success)
        {
            var libraries=db.Query("SELECT library_id FROM LIBRARIES ORDER BY library_id").ToArray();db.Execute("UPDATE INDEX_METADATA SET library_id=? WHERE singleton=1",libraries.Length==1?libraries[0][0]:null);
            string source=Digest(db,new[]{"SELECT presentation_id,source_path,content_hash,index_status FROM PRESENTATIONS ORDER BY presentation_id"});
            if(success)db.Execute("UPDATE INDEX_METADATA SET source_fingerprint=?,integrity_digest=?,last_successful_update=? WHERE singleton=1",source,Digest(db,DigestQueries),Utc());
            else db.Execute("UPDATE INDEX_METADATA SET source_fingerprint=?,integrity_digest=? WHERE singleton=1",source,Digest(db,DigestQueries));
        }
        static string SlideId(string presentationId,uint nativeId)
        {
            using(var hash=SHA256.Create()){var bytes=hash.ComputeHash(Encoding.UTF8.GetBytes("PIIDX/1/"+presentationId+"/"+nativeId.ToString(Invariant)));var id=new byte[16];Array.Copy(bytes,id,16);return new Guid(id).ToString("D");}
        }
        static string Json(string text)
        {
            if(text==null)return "null";var b=new StringBuilder("\"");foreach(char c in text){if(c=='"'||c=='\\')b.Append('\\').Append(c);else if(c<32)b.Append("\\u").Append(((int)c).ToString("x4",Invariant));else b.Append(c);}return b.Append('"').ToString();
        }
        static void StorePresentation(SqliteDatabase db,string libraryId,string presentationId,string relative,IndexedPresentation source,bool exists)
        {
            if(exists){db.Execute("DELETE FROM SLIDES WHERE presentation_id=?",presentationId);db.Execute("UPDATE PRESENTATIONS SET source_path=?,relative_path=?,file_type=?,file_size=?,modified_utc=?,content_hash=?,index_status='indexed',slide_count=? WHERE presentation_id=?",source.SourcePath,relative,source.FileType,source.Size,source.ModifiedUtc.ToString("o",Invariant),source.ContentHash,source.Slides.Count,presentationId);}
            else db.Execute("INSERT INTO PRESENTATIONS VALUES(?,?,?,?,?,?,?,?,?,?)",presentationId,libraryId,source.SourcePath,relative,source.FileType,source.Size,source.ModifiedUtc.ToString("o",Invariant),source.ContentHash,"indexed",source.Slides.Count);
            string category=(Path.GetDirectoryName(relative)??"").Replace('\\','/');
            string indexedKeywords=source.Keywords+" "+relative;
            foreach(var slide in source.Slides)
            {
                string id=SlideId(presentationId,slide.NativeId);
                db.Execute("INSERT INTO SLIDES VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",id,presentationId,slide.Position,slide.NativeId.ToString(Invariant),slide.PartUri,slide.Title,slide.Text,indexedKeywords,category,source.Author,slide.Layout,slide.Master,source.Width,source.Height,slide.ContentFingerprint);
                db.Execute("INSERT INTO SLIDE_SEARCH(slide_id,title,text,keywords) VALUES(?,?,?,?)",id,slide.Title,slide.Text,indexedKeywords);
                foreach(string tag in slide.Tags.Distinct(StringComparer.OrdinalIgnoreCase))db.Execute("INSERT INTO SLIDE_TAGS VALUES(?,?,?)",id,tag,tag.ToUpperInvariant());
                foreach(var shape in slide.Objects)
                {
                    string bounds="{\"schema\":1,\"space\":\"parent\",\"present\":"+(shape.HasTransform?"true":"false")+",\"x\":"+shape.X.ToString(Invariant)+",\"y\":"+shape.Y.ToString(Invariant)+",\"width\":"+shape.Width.ToString(Invariant)+",\"height\":"+shape.Height.ToString(Invariant)+"}";
                    string meta="{\"schema\":1,\"nativeId\":"+Json(shape.NativeId)+",\"parentKey\":"+Json(shape.ParentKey)+",\"rotation\":"+shape.Rotation.ToString(Invariant)+",\"flipH\":"+(shape.FlipH?"true":"false")+",\"flipV\":"+(shape.FlipV?"true":"false")+",\"xmlFingerprint\":"+Json(shape.XmlFingerprint)+"}";
                    db.Execute("INSERT INTO SLIDE_OBJECTS VALUES(?,?,?,?,?,?,?)",id,shape.Key,shape.Kind,shape.Name,shape.Text,bounds,meta);
                }
            }
        }
        internal LibraryRefreshResult Refresh(string libraryId,CancellationToken cancellation,IProgress<LibraryIndexProgress> progress=null,bool force=false)
        {
            Id(libraryId);cancellation.ThrowIfCancellationRequested();
            using(WriterLock(cancellation))using(var db=Connect(false))using(db.ObserveCancellation(cancellation))
            {
                VerifyContent(db);
                var library=db.Query("SELECT root_path,recursive FROM LIBRARIES WHERE library_id=?",libraryId).SingleOrDefault();if(library==null)throw new KeyNotFoundException("Library registration is missing.");
                // The exclusive OS file handle proves no other compliant writer owns this catalog.
                db.Execute("UPDATE INDEX_JOBS SET state='interrupted',completed_utc=? WHERE state='running'",Utc());
                string job=Guid.NewGuid().ToString("D"),activeSource=null;db.Execute("INSERT INTO INDEX_JOBS(job_id,library_id,started_utc,state) VALUES(?,?,?,'running')",job,libraryId,Utc());
                var result=new LibraryRefreshResult{JobId=job};bool transaction=false,committed=false;
                try
                {
                    string root=(string)library[0];var paths=PresentationSource.Discover(root,(long)library[1]!=0,cancellation).OrderBy(p=>p,StringComparer.OrdinalIgnoreCase).ToArray();
                    var fileNames=new HashSet<string>(paths.Select(p=>Relative(root,p)),StringComparer.OrdinalIgnoreCase);
                    var old=db.Query("SELECT presentation_id,relative_path,content_hash,index_status FROM PRESENTATIONS WHERE library_id=?",libraryId).ToArray();
                    var previous=old.ToDictionary(r=>(string)r[1],StringComparer.OrdinalIgnoreCase);var seen=new HashSet<string>(StringComparer.Ordinal);var captured=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
                    db.Execute("BEGIN IMMEDIATE");transaction=true;
                    foreach(string file in paths)
                    {
                        cancellation.ThrowIfCancellationRequested();activeSource=file;string relative=Relative(root,file),fingerprint=PresentationSource.Fingerprint(file,cancellation);object[] existing;
                        previous.TryGetValue(relative,out existing);
                        if(existing==null){var renamed=old.Where(r=>!seen.Contains((string)r[0])&&!fileNames.Contains((string)r[1])&&(string)r[2]==fingerprint).ToArray();if(renamed.Length==1)existing=renamed[0];}
                        string presentationId=existing==null?Guid.NewGuid().ToString("D"):(string)existing[0];seen.Add(presentationId);result.FilesSeen++;
                        bool unchanged=!force&&existing!=null&&(string)existing[2]==fingerprint&&(string)existing[3]=="indexed";
                        if(unchanged)
                        {
                            db.Execute("UPDATE SLIDES SET category=? WHERE presentation_id=?",(Path.GetDirectoryName(relative)??"").Replace('\\','/'),presentationId);
                            var info=new FileInfo(file);db.Execute("UPDATE PRESENTATIONS SET source_path=?,relative_path=?,file_size=?,modified_utc=? WHERE presentation_id=?",file,relative,info.Length,info.LastWriteTimeUtc.ToString("o",Invariant),presentationId);
                        }
                        else
                        {
                            var source=PresentationSource.Read(file,cancellation);if(source.ContentHash!=fingerprint)throw new IOException("Source changed during refresh; previous snapshot retained.");
                            StorePresentation(db,libraryId,presentationId,relative,source,existing!=null);result.FilesUpdated++;result.SlidesIndexed+=source.Slides.Count;
                        }
                        captured[file]=fingerprint;
                        if(progress!=null)progress.Report(new LibraryIndexProgress{LibraryId=libraryId,Source=file,FilesSeen=result.FilesSeen,FilesUpdated=result.FilesUpdated,SlidesIndexed=result.SlidesIndexed});
                    }
                    foreach(var record in old)if(!seen.Contains((string)record[0])){db.Execute("UPDATE PRESENTATIONS SET index_status='missing' WHERE presentation_id=?",record[0]);result.MissingFiles++;}
                    foreach(var source in captured){cancellation.ThrowIfCancellationRequested();activeSource=source.Key;if(PresentationSource.Fingerprint(source.Key,cancellation)!=source.Value)throw new IOException("Source changed after capture; previous snapshot retained.");}
                    cancellation.ThrowIfCancellationRequested();db.Execute("UPDATE LIBRARIES SET last_refresh_utc=? WHERE library_id=?",Utc(),libraryId);UpdateDigest(db,true);
                    db.Execute("INSERT INTO SLIDE_SEARCH(SLIDE_SEARCH) VALUES('integrity-check')");
                    db.Execute("UPDATE INDEX_JOBS SET state='completed',completed_utc=?,files_seen=?,files_updated=?,slides_indexed=? WHERE job_id=?",Utc(),result.FilesSeen,result.FilesUpdated,result.SlidesIndexed,job);
                    db.Execute("COMMIT");transaction=false;committed=true;return result;
                }
                catch(Exception error)
                {
                    if(transaction)db.RollbackIfActive();
                    if(!committed){db.Execute("UPDATE INDEX_JOBS SET state=?,completed_utc=? WHERE job_id=?",error is OperationCanceledException?"cancelled":"failed",Utc(),job);db.Execute("INSERT INTO INDEX_ERRORS(job_id,source_path,stage,error_code,message,recorded_utc) VALUES(?,?,?,?,?,?)",job,activeSource,"refresh",error.GetType().Name,"Library refresh did not commit. Previous source snapshot retained.",Utc());}
                    throw;
                }
            }
        }
    }
}