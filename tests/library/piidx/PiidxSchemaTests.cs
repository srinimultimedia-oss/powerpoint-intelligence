using System;
using System.IO;
using System.Linq;
using PowerPointIntelligence.Library;
public static class PiidxSchemaTests
{
    static int count;
    static void Check(bool value,string name){if(!value)throw new Exception("FAIL "+name);count++;}
    static void Throws<T>(Action action,string name) where T:Exception {try{action();}catch(T){count++;return;}throw new Exception("FAIL "+name);}
    public static string Run(string folder)
    {
        Directory.CreateDirectory(folder);
        string path=Path.Combine(folder,"schema-smoke.piidx");
        if(File.Exists(path))throw new IOException("Use a fresh test directory; existing database preserved.");
        using(var db=SqliteDatabase.Open(path,false))
        {
            PiidxSchema.Initialize(db);
            PiidxSchema.Validate(db);
            Check(db.Query("PRAGMA integrity_check").Single()[0].Equals("ok"),"integrity");
            Check(db.Query("SELECT name FROM sqlite_master WHERE name IN ('INDEX_METADATA','LIBRARIES','PRESENTATIONS','SLIDES','SLIDE_OBJECTS','SLIDE_TAGS','SLIDE_PREVIEWS','SLIDE_SEARCH','INDEX_JOBS','INDEX_ERRORS')").Count()==10,"ten required entities");
            db.Execute("CREATE TABLE PROBE(k TEXT PRIMARY KEY,v BLOB,n INTEGER)");
            string unicode="Résumé 日本語 العربية ' OR 1=1; --";
            db.Execute("INSERT INTO PROBE VALUES(?,?,?)",unicode,new byte[]{0,1,255},9223372036854775806L);
            var row=db.Query("SELECT k,v,n FROM PROBE WHERE k=?",unicode).Single();
            Check((string)row[0]==unicode,"unicode parameter binding");
            Check(((byte[])row[1]).SequenceEqual(new byte[]{0,1,255}),"blob binding");
            Check((long)row[2]==9223372036854775806L,"int64 binding");
            Check(db.Query("SELECT k FROM PROBE WHERE k=?","' OR 1=1 --").Count()==0,"SQL injection remains data");
            Throws<ArgumentException>(()=>db.Execute("DELETE FROM PROBE; DROP TABLE PROBE"),"multiple statements rejected before mutation");
            Check(db.Query("SELECT k FROM PROBE").Count()==1,"multi-statement rejection preserves data");
            db.Execute("BEGIN IMMEDIATE");db.Execute("DELETE FROM PROBE");db.Execute("ROLLBACK");
            Check(db.Query("SELECT k FROM PROBE").Count()==1,"rollback retains data");
            db.Execute("INSERT INTO SLIDE_SEARCH(slide_id,title,text,keywords) VALUES(?,?,?,?)","id","Quarterly résumé","forecast revenue","finance");
            Check(db.Query("SELECT slide_id FROM SLIDE_SEARCH WHERE SLIDE_SEARCH MATCH ?","forecast").Count()==1,"FTS5 content query");
            Check(db.Query("SELECT slide_id FROM SLIDE_SEARCH WHERE SLIDE_SEARCH MATCH ?","\"résumé\"").Count()==1,"FTS5 Unicode query");
            db.Execute("INSERT INTO LIBRARIES(library_id,root_path,name,recursive,registered_utc) VALUES('library','C:/fixture','fixture',1,'2026-10-29T00:00:00Z')");
            db.Execute("INSERT INTO PRESENTATIONS VALUES('presentation','library','C:/fixture/test.pptx','test.pptx','.pptx',1,'2026-10-29T00:00:00Z','hash','indexed',1)");
            db.Execute("INSERT INTO SLIDES VALUES('id','presentation',1,'256','ppt/slides/slide1.xml','title','text','keywords','category','author','layout','master',720.0,405.0,'hash')");
            db.Execute("DELETE FROM SLIDES WHERE slide_id='id'");
            Check(!db.Query("SELECT slide_id FROM SLIDE_SEARCH WHERE slide_id='id'").Any(),"slide deletion removes FTS entry");
            Throws<SqliteFailure>(()=>db.Execute("INSERT INTO PROBE(k) VALUES(?)",unicode),"constraint failure reported");
            Check(db.Query("SELECT k FROM PROBE").Count()==1,"connection usable after statement failure");
            Throws<ArgumentException>(()=>db.Execute("SELECT ?",Double.NaN),"nonfinite rejected");
            Throws<ArgumentException>(()=>db.Execute("SELECT ?",new object()),"unsupported binding rejected");
            Throws<ArgumentException>(()=>db.Execute("SELECT ?"),"parameter count validated");
            db.Execute("PRAGMA user_version=99");
            Throws<InvalidDataException>(()=>PiidxSchema.Validate(db),"future schema rejected");
            db.Execute("PRAGMA user_version=1");
            db.Execute("UPDATE INDEX_METADATA SET format_id='foreign'");
            Throws<InvalidDataException>(()=>PiidxSchema.Validate(db),"foreign format rejected");
            db.Execute("UPDATE INDEX_METADATA SET format_id=?",PiidxSchema.FormatId);
            db.Execute("DROP TABLE SLIDE_TAGS");
            db.Execute("CREATE TABLE SLIDE_TAGS(slide_id TEXT,tag TEXT,normalized_tag TEXT)");
            Throws<InvalidDataException>(()=>PiidxSchema.Validate(db),"weakened schema constraints rejected");
            db.Execute("DROP TABLE SLIDE_TAGS");
            Throws<InvalidDataException>(()=>PiidxSchema.Validate(db),"missing entity rejected");
        }
        using(var db=SqliteDatabase.Open(path,true))
        {
            Check(db.Query("SELECT k FROM PROBE").Count()==1,"reopen retains committed values");
            Throws<SqliteFailure>(()=>db.Execute("DELETE FROM PROBE"),"read-only enforced");
        }
        var disposed=SqliteDatabase.Open(":memory:",false);disposed.Dispose();
        Throws<ObjectDisposedException>(()=>disposed.Query("SELECT 1").ToList(),"disposed connection rejected");
        string foreign=Path.Combine(folder,"foreign.piidx");
        using(var db=SqliteDatabase.Open(foreign,false)){db.Execute("CREATE TABLE unrelated(x)");Throws<InvalidDataException>(()=>PiidxSchema.Initialize(db),"foreign database not adopted");Check(db.Query("SELECT name FROM sqlite_master WHERE name='unrelated'").Count()==1,"foreign data preserved");}
        return count+" PIIDX database foundation assertions passed; process="+(Environment.Is64BitProcess?"x64":"x86")+"; tests do not establish presentation indexing or Office integration.";
    }
}