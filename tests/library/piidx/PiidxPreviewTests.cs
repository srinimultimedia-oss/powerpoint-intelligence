using System;
using System.IO;
using System.Linq;
using System.Threading;
using PowerPointIntelligence.Library;
public static class PiidxPreviewTests
{
    static int count;
    static void Check(bool value,string name){if(!value)throw new Exception("FAIL "+name);count++;}
    static void Throws<T>(Action action,string name)where T:Exception{try{action();}catch(T){count++;return;}throw new Exception("FAIL "+name);}
    public static string Run(string output)
    {
        count=0;
        Directory.CreateDirectory(output);var png=Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYGD4DwABBAEAX+XDSwAAAABJRU5ErkJggg==");
        PiidxPng.Validate(png,1,1);Check(true,"independent PNG fixture validates");
        Throws<InvalidDataException>(()=>PiidxPng.Validate(png,2,1),"metadata dimensions must match PNG");
        Throws<InvalidDataException>(()=>PiidxPng.Validate(png.Take(30).ToArray(),1,1),"truncated PNG rejected");
        var bad=(byte[])png.Clone();bad[40]^=1;Throws<InvalidDataException>(()=>PiidxPng.Validate(bad,1,1),"PNG chunk checksum enforced");
        Throws<InvalidDataException>(()=>PiidxPng.Validate(png.Concat(new byte[]{0}).ToArray(),1,1),"PNG trailing bytes rejected");
        string root=Path.Combine(output,"source");Directory.CreateDirectory(root);string source=Path.Combine(root,"source.pptx");PiidxSourceTests.Fixture(source);
        var catalog=PiidxCatalog.OpenOrCreate(Path.Combine(output,"preview.piidx"));string library=catalog.Register(root,"Preview test",true);catalog.Refresh(library,CancellationToken.None);
        var slide=catalog.Search(new LibrarySearchRequest(),CancellationToken.None).Slides.First();
        Check(catalog.Preview(slide.SlideId,1,1,CancellationToken.None)==null,"missing cache returns no fabricated preview");
        catalog.PutPreview(slide,png,1,1,CancellationToken.None);Check(catalog.Preview(slide.SlideId,1,1,CancellationToken.None).SequenceEqual(png),"PNG bytes persist");
        catalog=PiidxCatalog.Open(catalog.IndexPath);Check(catalog.Preview(slide.SlideId,1,1,CancellationToken.None).SequenceEqual(png),"PNG cache survives reopen");
        var returned=catalog.Preview(slide.SlideId,1,1,CancellationToken.None);returned[0]=0;Check(catalog.Preview(slide.SlideId,1,1,CancellationToken.None)[0]==137,"cache returns independent byte storage");
        using(var db=SqliteDatabase.Open(catalog.IndexPath,false))db.Execute("UPDATE SLIDE_PREVIEWS SET image=?",bad);
        Throws<InvalidDataException>(()=>catalog.Preview(slide.SlideId,1,1,CancellationToken.None),"stored PNG tampering rejected before decode");
        Check(catalog.Search(new LibrarySearchRequest(),CancellationToken.None).Total==2,"bad derivative cache does not discard indexed sources");
        catalog.PutPreview(slide,png,1,1,CancellationToken.None);catalog.VerifyIntegrity();
        var cancelled=new CancellationToken(true);Throws<OperationCanceledException>(()=>catalog.PutPreview(slide,png,1,1,cancelled),"cache write cancellation");
        catalog.Refresh(library,CancellationToken.None,null,true);Check(catalog.Preview(slide.SlideId,1,1,CancellationToken.None)==null,"rebuild invalidates derivative previews");
        File.Delete(source);PiidxSourceTests.Fixture(source,true);catalog.Refresh(library,CancellationToken.None);
        Throws<InvalidOperationException>(()=>catalog.PutPreview(slide,png,1,1,CancellationToken.None),"stale rendered selection cannot populate current cache");
        return count+" preview-cache assertions passed; PNG fixture is synthetic, not evidence of native slide rendering.";
    }
}