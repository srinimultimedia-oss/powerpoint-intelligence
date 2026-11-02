using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using PowerPointIntelligence.Library;
public static class PiidxCorpusTests
{
    public static string Run(string output,string sources)
    {
        Directory.CreateDirectory(output);var catalog=PiidxCatalog.OpenOrCreate(Path.Combine(output,"local-library.piidx"));string id=catalog.Register(sources,"Local template corpus",true);
        var time=Stopwatch.StartNew();var first=catalog.Refresh(id,CancellationToken.None);long initial=time.ElapsedMilliseconds;
        if(first.FilesSeen!=62||first.FilesUpdated!=62||first.SlidesIndexed!=497)throw new Exception("Corpus persisted counts do not match source inventory.");
        var page=catalog.Search(new LibrarySearchRequest{Limit=24},CancellationToken.None);if(page.Total!=497||page.Slides.Count!=24)throw new Exception("Corpus browsing/paging failed.");
        if(catalog.Search(new LibrarySearchRequest{Query="cube"},CancellationToken.None).Total<1)throw new Exception("Indexed content search failed.");
        catalog.VerifyIntegrity();time.Restart();var second=catalog.Refresh(id,CancellationToken.None);long incremental=time.ElapsedMilliseconds;
        if(second.FilesUpdated!=0||second.SlidesIndexed!=0)throw new Exception("Unchanged corpus was reparsed.");
        var reader=Task.Run(()=>{for(int i=0;i<12;i++){var next=catalog.Search(new LibrarySearchRequest{Query="pyramid",Limit=5},CancellationToken.None);if(next.Total<1||next.Slides.Count<1)throw new Exception("Concurrent search lost committed results.");}});
        catalog.Refresh(id,CancellationToken.None);reader.GetAwaiter().GetResult();
        using(var db=SqliteDatabase.Open(catalog.IndexPath,true))
        {
            long objects=(long)db.Query("SELECT count(*) FROM SLIDE_OBJECTS").Single()[0],tags=(long)db.Query("SELECT count(*) FROM SLIDE_TAGS").Single()[0];
            if(objects<1||tags<1)throw new Exception("Object/tag persistence is empty.");
            return "7 corpus assertions passed; packages=62; slides=497; objects="+objects+"; tags="+tags+"; initialMs="+initial+"; incrementalMs="+incremental+"; databaseBytes="+new FileInfo(catalog.IndexPath).Length+". No native Office rendering/insertion was executed.";
        }
    }
}