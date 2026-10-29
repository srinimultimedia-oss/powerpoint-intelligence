using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
namespace PowerPointIntelligence.Library
{
    internal static class PiidxSchema
    {
        internal const string FormatId="PowerPointIntelligence.PIIDX";
        internal const int Version=1,ApplicationId=0x50494958;
        internal const string EngineVersion="0.1.0";
        internal static readonly string[] Definitions={
            "CREATE TABLE INDEX_METADATA(singleton INTEGER PRIMARY KEY CHECK(singleton=1),format_id TEXT NOT NULL,schema_version INTEGER NOT NULL,engine_version TEXT NOT NULL,created_utc TEXT NOT NULL,last_successful_update TEXT,library_id TEXT,source_fingerprint TEXT,integrity_digest TEXT)",
            "CREATE TABLE LIBRARIES(library_id TEXT PRIMARY KEY NOT NULL,root_path TEXT NOT NULL UNIQUE,name TEXT NOT NULL,recursive INTEGER NOT NULL CHECK(recursive IN(0,1)),registered_utc TEXT NOT NULL,last_refresh_utc TEXT,enabled INTEGER NOT NULL DEFAULT 1 CHECK(enabled IN(0,1)))",
            "CREATE TABLE PRESENTATIONS(presentation_id TEXT PRIMARY KEY NOT NULL,library_id TEXT NOT NULL REFERENCES LIBRARIES(library_id) ON DELETE CASCADE,source_path TEXT NOT NULL,relative_path TEXT NOT NULL,file_type TEXT NOT NULL CHECK(file_type IN('.pptx','.potx')),file_size INTEGER NOT NULL CHECK(file_size>=0),modified_utc TEXT NOT NULL,content_hash TEXT NOT NULL,index_status TEXT NOT NULL,slide_count INTEGER NOT NULL CHECK(slide_count>=0),UNIQUE(library_id,relative_path))",
            "CREATE TABLE SLIDES(slide_id TEXT PRIMARY KEY NOT NULL,presentation_id TEXT NOT NULL REFERENCES PRESENTATIONS(presentation_id) ON DELETE CASCADE,position INTEGER NOT NULL CHECK(position>0),native_slide_id TEXT NOT NULL,part_uri TEXT NOT NULL,title TEXT NOT NULL,text TEXT NOT NULL,keywords TEXT NOT NULL,category TEXT NOT NULL,author TEXT NOT NULL,layout TEXT NOT NULL,master TEXT NOT NULL,width REAL NOT NULL CHECK(width>0),height REAL NOT NULL CHECK(height>0),content_fingerprint TEXT NOT NULL,UNIQUE(presentation_id,native_slide_id),UNIQUE(presentation_id,position))",
            "CREATE TABLE SLIDE_OBJECTS(slide_id TEXT NOT NULL REFERENCES SLIDES(slide_id) ON DELETE CASCADE,object_id TEXT NOT NULL,kind TEXT NOT NULL,name TEXT NOT NULL,text TEXT NOT NULL,bounds_json TEXT NOT NULL,metadata_json TEXT NOT NULL,PRIMARY KEY(slide_id,object_id))",
            "CREATE TABLE SLIDE_TAGS(slide_id TEXT NOT NULL REFERENCES SLIDES(slide_id) ON DELETE CASCADE,tag TEXT NOT NULL,normalized_tag TEXT NOT NULL,PRIMARY KEY(slide_id,normalized_tag))",
            "CREATE TABLE SLIDE_PREVIEWS(slide_id TEXT NOT NULL REFERENCES SLIDES(slide_id) ON DELETE CASCADE,width INTEGER NOT NULL CHECK(width>0 AND width<=4096),height INTEGER NOT NULL CHECK(height>0 AND height<=4096),mime TEXT NOT NULL CHECK(mime='image/png'),image BLOB NOT NULL,content_hash TEXT NOT NULL,source_fingerprint TEXT NOT NULL,rendered_utc TEXT NOT NULL,PRIMARY KEY(slide_id,width,height))",
            "CREATE VIRTUAL TABLE SLIDE_SEARCH USING fts5(slide_id UNINDEXED,title,text,keywords,tokenize='unicode61 remove_diacritics 2')",
            "CREATE TABLE INDEX_JOBS(job_id TEXT PRIMARY KEY NOT NULL,library_id TEXT REFERENCES LIBRARIES(library_id) ON DELETE SET NULL,started_utc TEXT NOT NULL,completed_utc TEXT,state TEXT NOT NULL CHECK(state IN('running','completed','cancelled','failed','interrupted')),files_seen INTEGER NOT NULL DEFAULT 0,files_updated INTEGER NOT NULL DEFAULT 0,slides_indexed INTEGER NOT NULL DEFAULT 0)",
            "CREATE TABLE INDEX_ERRORS(error_id INTEGER PRIMARY KEY,job_id TEXT REFERENCES INDEX_JOBS(job_id) ON DELETE CASCADE,source_path TEXT,stage TEXT NOT NULL,error_code TEXT NOT NULL,message TEXT NOT NULL,recorded_utc TEXT NOT NULL)",
            "CREATE INDEX IX_PRESENTATIONS_LIBRARY ON PRESENTATIONS(library_id,index_status)",
            "CREATE INDEX IX_SLIDES_CATEGORY ON SLIDES(category)",
            "CREATE INDEX IX_SLIDE_TAGS_NORMALIZED ON SLIDE_TAGS(normalized_tag,slide_id)",
            "CREATE TRIGGER SLIDES_DELETE_SEARCH AFTER DELETE ON SLIDES BEGIN DELETE FROM SLIDE_SEARCH WHERE slide_id=OLD.slide_id; END"
        };
        static readonly Dictionary<string,string[]> RequiredColumns=new Dictionary<string,string[]>{
            {"INDEX_METADATA",new[]{"singleton","format_id","schema_version","engine_version","created_utc","last_successful_update","library_id","source_fingerprint","integrity_digest"}},
            {"LIBRARIES",new[]{"library_id","root_path","name","recursive","registered_utc","last_refresh_utc","enabled"}},
            {"PRESENTATIONS",new[]{"presentation_id","library_id","source_path","relative_path","file_type","file_size","modified_utc","content_hash","index_status","slide_count"}},
            {"SLIDES",new[]{"slide_id","presentation_id","position","native_slide_id","part_uri","title","text","keywords","category","author","layout","master","width","height","content_fingerprint"}},
            {"SLIDE_OBJECTS",new[]{"slide_id","object_id","kind","name","text","bounds_json","metadata_json"}},
            {"SLIDE_TAGS",new[]{"slide_id","tag","normalized_tag"}},
            {"SLIDE_PREVIEWS",new[]{"slide_id","width","height","mime","image","content_hash","source_fingerprint","rendered_utc"}},
            {"SLIDE_SEARCH",new[]{"slide_id","title","text","keywords"}},
            {"INDEX_JOBS",new[]{"job_id","library_id","started_utc","completed_utc","state","files_seen","files_updated","slides_indexed"}},
            {"INDEX_ERRORS",new[]{"error_id","job_id","source_path","stage","error_code","message","recorded_utc"}}
        };
        internal static void Initialize(SqliteDatabase db)
        {
            long current=(long)db.Query("PRAGMA user_version").Single()[0];
            if(current!=0){Validate(db);return;}
            if(db.Query("SELECT name FROM sqlite_master WHERE name NOT LIKE 'sqlite_%'").Any())throw new InvalidDataException("Unversioned nonempty database is not a PIIDX index; preserved.");
            db.Execute("BEGIN IMMEDIATE");
            try
            {
                foreach(string sql in Definitions)db.Execute(sql);
                db.Execute("INSERT INTO INDEX_METADATA(singleton,format_id,schema_version,engine_version,created_utc) VALUES(1,?,?,?,?)",FormatId,Version,EngineVersion,DateTime.UtcNow.ToString("o",CultureInfo.InvariantCulture));
                db.Execute("PRAGMA application_id="+ApplicationId.ToString(CultureInfo.InvariantCulture));
                db.Execute("PRAGMA user_version="+Version.ToString(CultureInfo.InvariantCulture));
                db.Execute("COMMIT");
            }
            catch{db.Execute("ROLLBACK");throw;}
            Validate(db);
        }
        internal static void Validate(SqliteDatabase db)
        {
            if((long)db.Query("PRAGMA application_id").Single()[0]!=ApplicationId||(long)db.Query("PRAGMA user_version").Single()[0]!=Version)throw new InvalidDataException("Unknown PIIDX format/schema version; no automatic downgrade or overwrite.");
            foreach(var table in RequiredColumns)
            {
                string[] columns=db.Query("PRAGMA table_info('"+table.Key+"')").Select(row=>(string)row[1]).ToArray();
                if(!table.Value.SequenceEqual(columns))throw new InvalidDataException("PIIDX schema columns are missing or incompatible: "+table.Key);
            }
            foreach(string expected in Definitions)
            {
                var name=Regex.Match(expected,@"^CREATE (?:VIRTUAL )?(?:TABLE|INDEX|TRIGGER) (\w+)").Groups[1].Value;
                var definition=db.Query("SELECT sql FROM sqlite_master WHERE name=?",name).SingleOrDefault();
                if(definition==null||!String.Equals((string)definition[0],expected,StringComparison.Ordinal))throw new InvalidDataException("PIIDX schema definition is incompatible: "+name);
            }
            var headers=db.Query("SELECT format_id,schema_version FROM INDEX_METADATA").ToArray();
            if(headers.Length!=1||(string)headers[0][0]!=FormatId||(long)headers[0][1]!=Version)throw new InvalidDataException("PIIDX format metadata is invalid.");
            if(db.Query("PRAGMA integrity_check").Any(row=>!String.Equals(row[0],"ok")))throw new InvalidDataException("PIIDX SQLite integrity verification failed.");
            if(db.Query("PRAGMA foreign_key_check").Any())throw new InvalidDataException("PIIDX foreign-key verification failed.");
        }
    }
}