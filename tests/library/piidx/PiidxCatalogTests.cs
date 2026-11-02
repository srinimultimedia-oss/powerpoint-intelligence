using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using PowerPointIntelligence.Library;
public static class PiidxCatalogTests
{
    static int count;
    static void Check(bool value,string name){if(!value)throw new Exception("FAIL "+name);count++;}
    static void Throws<T>(Action action,string name) where T:Exception{try{action();}catch(T){count++;return;}throw new Exception("FAIL "+name);}
    sealed class ImmediateProgress : IProgress<LibraryIndexProgress>{readonly Action<LibraryIndexProgress> callback;internal ImmediateProgress(Action<LibraryIndexProgress> callback){this.callback=callback;}public void Report(LibraryIndexProgress value){callback(value);}}
    static LibrarySearchPage All(PiidxCatalog catalog,string library=null){return catalog.Search(new LibrarySearchRequest{LibraryId=library,Limit=1000},CancellationToken.None);}
    public static string Run(string folder)
    {
        count=0;
        Directory.CreateDirectory(folder);string root=Path.Combine(folder,"sources"),finance=Path.Combine(root,"Finance");Directory.CreateDirectory(finance);
        string pptx=Path.Combine(root,"日本語.pptx"),potx=Path.Combine(finance,"report.potx");PiidxSourceTests.Fixture(pptx);PiidxSourceTests.Fixture(potx);
        string path=Path.Combine(folder,"catalog.piidx");var catalog=PiidxCatalog.OpenOrCreate(path);catalog.VerifyIntegrity();Check(All(catalog).Total==0,"empty catalog opens");
        string library=catalog.Register(root,"Local library",true);Check(catalog.Register(root.ToUpperInvariant(),"Local library",true)==library,"root registration reused case-insensitively");
        var initial=catalog.Refresh(library,CancellationToken.None);Check(initial.FilesSeen==2&&initial.FilesUpdated==2&&initial.SlidesIndexed==4,"PPTX/POTX persisted indexing");
        var original=All(catalog);Check(original.Total==4&&original.Slides.Count==4,"browse persisted slides");
        catalog=PiidxCatalog.Open(path);Check(catalog.Libraries().Single().Id==library&&All(catalog).Total==4,"catalog reopen and registration persist");
        var title=new LibrarySearchRequest{Query="Résumé 日本語 العربية",Mode=LibrarySearchMode.ExactText};Check(catalog.Search(title,CancellationToken.None).Total==2,"exact text search");
        Check(catalog.Search(new LibrarySearchRequest{Query="Résumé 日本語",Mode=LibrarySearchMode.Phrase},CancellationToken.None).Total==2,"phrase search");
        Check(catalog.Search(new LibrarySearchRequest{Query="résu",Mode=LibrarySearchMode.Prefix},CancellationToken.None).Total==2,"prefix search");
        Check(catalog.Search(new LibrarySearchRequest{Query="Résumé العربية"},CancellationToken.None).Total==2,"multi-keyword Unicode search");
        Check(catalog.Search(new LibrarySearchRequest{Category="Finance"},CancellationToken.None).Total==2,"category filter");
        Check(catalog.Search(new LibrarySearchRequest{Tag="category=finance"},CancellationToken.None).Total==2,"tag filter");
        var sourceSlide=original.Slides.First(s=>s.SourcePath==pptx&&s.NativeId=="256");
        Check(catalog.Search(new LibrarySearchRequest{PresentationId=sourceSlide.PresentationId},CancellationToken.None).Total==2,"presentation filter");
        var firstPage=catalog.Search(new LibrarySearchRequest{Limit=1},CancellationToken.None);var secondPage=catalog.Search(new LibrarySearchRequest{Limit=1,Offset=1},CancellationToken.None);
        Check(firstPage.Total==4&&firstPage.Slides.Count==1&&firstPage.Slides[0].SlideId!=secondPage.Slides[0].SlideId,"deterministic pagination");
        foreach(string query in new[]{"' OR 1=1 --","'; DROP TABLE SLIDES; --","*","\"",":()", "AND OR NOT", "\"résumé\" "})catalog.Search(new LibrarySearchRequest{Query=query},CancellationToken.None);
        Check(All(catalog).Total==4,"special characters cannot execute SQL");
        catalog.VerifySource(sourceSlide,CancellationToken.None);Check(true,"source freshness verified");
        var incremental=catalog.Refresh(library,CancellationToken.None);Check(incremental.FilesUpdated==0&&incremental.FilesSeen==2,"unchanged sources skip extraction");
        Check(All(catalog).Slides.Select(s=>s.SlideId).OrderBy(s=>s).SequenceEqual(original.Slides.Select(s=>s.SlideId).OrderBy(s=>s)),"incremental IDs retained");
        string renamed=Path.Combine(root,"renamed.potx");File.Move(potx,renamed);catalog.Refresh(library,CancellationToken.None);
        var renamedReference=All(catalog).Slides.First(s=>s.SourcePath==renamed&&s.NativeId=="256");var beforeRename=original.Slides.First(s=>s.SourcePath==potx&&s.NativeId=="256");
        Check(renamedReference.SlideId==beforeRename.SlideId&&renamedReference.PresentationId==beforeRename.PresentationId,"unique-hash rename retains IDs");
        Check(catalog.Search(new LibrarySearchRequest{Category="Finance"},CancellationToken.None).Total==0,"renamed source category updates");
        File.Delete(pptx);PiidxSourceTests.Fixture(pptx,true);Throws<InvalidOperationException>(()=>catalog.VerifySource(sourceSlide,CancellationToken.None),"stale source rejected before insertion");
        catalog.Refresh(library,CancellationToken.None);var reordered=All(catalog).Slides.Single(s=>s.SourcePath==pptx&&s.NativeId=="256");
        Check(reordered.SlideId==sourceSlide.SlideId&&reordered.Position==2,"slide IDs survive native reordering");
        byte[] valid=File.ReadAllBytes(pptx);File.WriteAllText(pptx,"not a presentation");
        Throws<InvalidDataException>(()=>catalog.Refresh(library,CancellationToken.None),"unreadable source fails refresh");
        Check(All(catalog).Total==4,"failed refresh retains previous snapshot");File.WriteAllBytes(pptx,valid);catalog.VerifyIntegrity();
        using(var db=SqliteDatabase.Open(path,true)){Check((long)db.Query("SELECT count(*) FROM INDEX_JOBS WHERE state='failed'").Single()[0]==1,"failed job recorded");Check((long)db.Query("SELECT count(*) FROM INDEX_ERRORS").Single()[0]>=1,"sanitized error recorded");}
        var cancel=new CancellationTokenSource();
        Throws<OperationCanceledException>(()=>catalog.Refresh(library,cancel.Token,new ImmediateProgress(p=>cancel.Cancel()),true),"refresh cancellation");
        Check(All(catalog).Total==4,"cancelled refresh rolls back");catalog.VerifyIntegrity();
        bool readDuringRefresh=false;catalog.Refresh(library,CancellationToken.None,new ImmediateProgress(p=>{readDuringRefresh=All(catalog).Total==4;}),true);Check(readDuringRefresh,"concurrent reader sees committed snapshot");
        bool changed=false;
        try{Throws<IOException>(()=>catalog.Refresh(library,CancellationToken.None,new ImmediateProgress(p=>{if(!changed&&p.Source==pptx){File.AppendAllText(pptx,"changed after capture");changed=true;}}),true),"source mutation during refresh rejected");}
        finally{File.WriteAllBytes(pptx,valid);}
        Check(All(catalog).Total==4,"source-race failure retains snapshot");
        string outside=Path.Combine(folder,"outside.potx");File.Move(renamed,outside);catalog.Refresh(library,CancellationToken.None);
        Check(All(catalog).Total==2,"missing source excluded without fabricating data");Throws<FileNotFoundException>(()=>catalog.Slide(renamedReference.SlideId,CancellationToken.None),"missing source reference rejected");
        File.Move(outside,renamed);catalog.Refresh(library,CancellationToken.None);Check(All(catalog).Slides.Any(s=>s.SlideId==renamedReference.SlideId),"restored source retains identity");
        string rootB=Path.Combine(folder,"other-library");Directory.CreateDirectory(rootB);PiidxSourceTests.Fixture(Path.Combine(rootB,"second.pptx"));
        string libraryB=catalog.Register(rootB,"Other",true);catalog.Refresh(libraryB,CancellationToken.None);
        Check(All(catalog).Total==6&&All(catalog,libraryB).Total==2&&All(catalog,library).Total==4,"multiple registered libraries remain isolated");
        catalog.Remove(libraryB);Check(All(catalog).Total==4&&catalog.Libraries().Count==1,"removing one library preserves another");
        string staleJob=Guid.NewGuid().ToString("D");using(var db=SqliteDatabase.Open(path,false))db.Execute("INSERT INTO INDEX_JOBS(job_id,library_id,started_utc,state) VALUES(?,?,?,'running')",staleJob,library,DateTime.UtcNow.ToString("o"));
        using(var held=new FileStream(path+".writer.lock",FileMode.Open,FileAccess.ReadWrite,FileShare.None))using(var waiting=new CancellationTokenSource())
        {
            waiting.CancelAfter(30);Throws<OperationCanceledException>(()=>catalog.Refresh(library,waiting.Token),"exclusive writer wait is cancellable");
            using(var db=SqliteDatabase.Open(path,true))Check((string)db.Query("SELECT state FROM INDEX_JOBS WHERE job_id=?",staleJob).Single()[0]=="running","live writer handle prevents recovery from stale markers alone");
        }
        catalog.Refresh(library,CancellationToken.None);using(var db=SqliteDatabase.Open(path,true))Check((string)db.Query("SELECT state FROM INDEX_JOBS WHERE job_id=?",staleJob).Single()[0]=="interrupted","exclusive writer ownership permits interrupted-job recovery");
        string damaged=Path.Combine(folder,"damaged.piidx");byte[] damage=new byte[]{1,2,3,4};File.WriteAllBytes(damaged,damage);
        Throws<IOException>(()=>PiidxCatalog.OpenOrCreate(damaged),"corrupt file is rejected without recreation");Check(File.ReadAllBytes(damaged).SequenceEqual(damage),"corrupt source index preserved byte-for-byte");
        using(var sourceHandle=new FileStream(pptx,FileMode.Open,FileAccess.Read,FileShare.Read))catalog.Refresh(library,CancellationToken.None,null,true);
        Check(All(catalog).Total==4,"indexing needs no source write access");
        using(var db=SqliteDatabase.Open(path,false))db.Execute("UPDATE SLIDES SET title='tampered' WHERE slide_id=?",sourceSlide.SlideId);
        Throws<InvalidDataException>(()=>PiidxCatalog.Open(path),"logical tampering rejected");
        string sampleRoot=Path.Combine(folder,"sample-source");Directory.CreateDirectory(sampleRoot);PiidxSourceTests.Fixture(Path.Combine(sampleRoot,"sample.pptx"));
        var sample=PiidxCatalog.OpenOrCreate(Path.Combine(folder,"sample-library.piidx"));string sampleId=sample.Register(sampleRoot,"Sample",true);sample.Refresh(sampleId,CancellationToken.None);sample.VerifyIntegrity();
        Check(All(sample).Total==2,"test-presentation PIIDX sample generated");
        sample.Remove(sampleId);Check(sample.Libraries().Count==0&&All(sample).Total==0,"removal cascades catalog and search only");Check(Directory.GetFiles(sampleRoot).Length==1,"removal preserves source presentations");sample.VerifyIntegrity();
        sampleId=sample.Register(sampleRoot,"Sample",true);sample.Refresh(sampleId,CancellationToken.None);Check(All(sample).Total==2,"delivered sample remains populated");
        return count+" persisted PIIDX catalog assertions passed; no Office automation executed.";
    }
}