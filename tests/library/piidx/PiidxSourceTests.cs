using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using PowerPointIntelligence.Library;
public static class PiidxSourceTests
{
    static int count;
    static void Check(bool condition,string name){if(!condition)throw new Exception("FAIL "+name);count++;}
    static void Throws<T>(Action action,string name) where T:Exception{try{action();}catch(T){count++;return;}throw new Exception("FAIL "+name);}
    static void Part(ZipArchive zip,string name,string xml){using(var output=new StreamWriter(zip.CreateEntry(name).Open(),new UTF8Encoding(false)))output.Write(xml);}
    internal static void Fixture(string file,bool reverse=false,bool duplicate=false,bool external=false,bool dtd=false,bool strict=false)
    {
        string p=strict?"http://purl.oclc.org/ooxml/presentationml/main":"http://schemas.openxmlformats.org/presentationml/2006/main";
        string a=strict?"http://purl.oclc.org/ooxml/drawingml/main":"http://schemas.openxmlformats.org/drawingml/2006/main";
        string r=strict?"http://purl.oclc.org/ooxml/officeDocument/relationships":"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        using(var stream=new FileStream(file,FileMode.CreateNew))using(var zip=new ZipArchive(stream,ZipArchiveMode.Create))
        {
            string ids=reverse?"<p:sldId id='257' r:id='r2'/><p:sldId id='256' r:id='r1'/>":"<p:sldId id='256' r:id='r1'/><p:sldId id='"+(duplicate?"256":"257")+"' r:id='r2'/>";
            Part(zip,"[Content_Types].xml","<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override PartName='/ppt/presentation.xml' ContentType='application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'/></Types>");
            Part(zip,"ppt/presentation.xml",(dtd?"<!DOCTYPE presentation [<!ENTITY data SYSTEM 'file:///not-read'>]>":"")+"<p:presentation xmlns:p='"+p+"' xmlns:r='"+r+"'><p:sldIdLst>"+ids+"</p:sldIdLst><p:sldSz cx='9144000' cy='5143500'/></p:presentation>");
            Part(zip,"ppt/_rels/presentation.xml.rels","<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='"+r+"/slide' Target='"+(external?"https://invalid.example/slide.xml' TargetMode='External":"slides/slide1.xml")+"'/><Relationship Id='r2' Type='"+r+"/slide' Target='slides/slide2.xml'/></Relationships>");
            Part(zip,"ppt/slides/slide1.xml","<p:sld xmlns:p='"+p+"' xmlns:a='"+a+"'><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id='2' name='Title'/><p:cNvSpPr/><p:nvPr><p:ph type='title'/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm rot='1800000' flipH='1'><a:off x='12700' y='25400'/><a:ext cx='127000' cy='254000'/></a:xfrm></p:spPr><p:txBody><a:p><a:r><a:t>Résumé 日本語 العربية</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>");
            Part(zip,"ppt/slides/slide2.xml","<p:sld xmlns:p='"+p+"'><p:cSld><p:spTree/></p:cSld></p:sld>");
            Part(zip,"ppt/slides/_rels/slide1.xml.rels","<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='tags' Type='"+r+"/tags' Target='../tags/tag1.xml'/><Relationship Id='tags2' Type='"+r+"/tags' Target='../tags/tag2.xml'/></Relationships>");
            Part(zip,"ppt/tags/tag1.xml","<p:tagLst xmlns:p='"+p+"'><p:tag name='CATEGORY' val='Finance'/></p:tagLst>");
            Part(zip,"ppt/tags/tag2.xml","<p:tagLst xmlns:p='"+p+"'><p:tag name='LONG_METADATA' val='"+new string('X',32768)+"'/></p:tagLst>");
            Part(zip,"docProps/core.xml","<cp:coreProperties xmlns:cp='http://schemas.openxmlformats.org/package/2006/metadata/core-properties' xmlns:dc='http://purl.org/dc/elements/1.1/'><dc:creator>Test author</dc:creator><cp:keywords>finance; strategy</cp:keywords></cp:coreProperties>");
        }
    }
    public static string Run(string folder,string originals)
    {
        count=0;
        Directory.CreateDirectory(folder);string source=Path.Combine(folder,"日本語.pptx");Fixture(source);
        string hash=PresentationSource.Fingerprint(source,CancellationToken.None);
        var first=PresentationSource.Read(source,CancellationToken.None);
        Check(first.Slides.Count==2,"PPTX enumeration");
        Check(first.Width==720&&first.Height==405,"native dimensions");
        Check(first.Author=="Test author"&&first.Keywords=="finance; strategy","core metadata");
        Check(first.Slides[0].Title=="Résumé 日本語 العربية","Unicode title and text");
        Check(first.Slides[0].Tags.Contains("CATEGORY=Finance"),"source user-defined tags");
        Check(first.Slides[0].Tags.Count==2&&first.Slides[0].Tags.Any(t=>t=="LONG_METADATA="+new string('X',32768)),"multiple parts and 32K tag values preserved");
        Check(first.Slides[1].Text==""&&first.Slides[1].Objects.Count==0,"empty slide");
        Check(first.Slides[0].NativeId==256&&first.Slides[0].Position==1,"native identity separate from position");
        var shape=first.Slides[0].Objects.Single();
        Check(shape.NativeId=="2"&&shape.Name=="Title"&&shape.Kind=="sp","object metadata");
        Check(shape.HasTransform&&shape.X==12700&&shape.Width==127000&&shape.Rotation==1800000&&shape.FlipH,"raw parent transform");
        Check(hash==PresentationSource.Fingerprint(source,CancellationToken.None),"source bytes unchanged");
        string moved=Path.Combine(folder,"reordered.potx");Fixture(moved,true);
        var second=PresentationSource.Read(moved,CancellationToken.None);
        Check(second.FileType==".potx"&&second.Slides.Count==2,"POTX source supported");
        Check(second.Slides[1].NativeId==first.Slides[0].NativeId&&second.Slides[1].Position==2,"identity survives reordering");
        string bad=Path.Combine(folder,"duplicate.pptx");Fixture(bad,false,true);
        Throws<InvalidDataException>(()=>PresentationSource.Read(bad,CancellationToken.None),"duplicate slide IDs rejected");
        string remote=Path.Combine(folder,"external.pptx");Fixture(remote,false,false,true);
        Throws<InvalidDataException>(()=>PresentationSource.Read(remote,CancellationToken.None),"external slide not fetched");
        string dtd=Path.Combine(folder,"dtd.pptx");Fixture(dtd,false,false,false,true);
        Throws<System.Xml.XmlException>(()=>PresentationSource.Read(dtd,CancellationToken.None),"DTD prohibited");
        string strict=Path.Combine(folder,"strict.pptx");Fixture(strict,false,false,false,false,true);
        Check(PresentationSource.Read(strict,CancellationToken.None).Slides[0].Title==first.Slides[0].Title,"strict OOXML supported");
        var cancelled=new CancellationToken(true);
        Throws<OperationCanceledException>(()=>PresentationSource.Read(source,cancelled),"read cancellation");
        Throws<OperationCanceledException>(()=>PresentationSource.Discover(folder,true,cancelled).ToArray(),"discovery cancellation");
        File.WriteAllText(Path.Combine(folder,"irrelevant.txt"),"not a source");
        Check(PresentationSource.Discover(folder,false,CancellationToken.None).Count()==6,"supported file discovery");
        int packages=0,slides=0;bool uniqueObjectKeys=true,repeatedNativeIds=false;
        if(!String.IsNullOrEmpty(originals))foreach(string file in PresentationSource.Discover(originals,true,CancellationToken.None))
        {
            IndexedPresentation presentation;try{presentation=PresentationSource.Read(file,CancellationToken.None);}catch(Exception ex){throw new Exception("Source package failed: "+Path.GetFileName(file)+"; "+ex.Message,ex);}packages++;slides+=presentation.Slides.Count;foreach(var slide in presentation.Slides){uniqueObjectKeys &= slide.Objects.Select(o=>o.Key).Distinct().Count()==slide.Objects.Count;repeatedNativeIds |= slide.Objects.Select(o=>o.NativeId).Distinct().Count()!=slide.Objects.Count;}
        }
        if(!String.IsNullOrEmpty(originals)){Check(packages==62,"62 local source packages extracted");Check(slides==497,"497 local source slides extracted");Check(uniqueObjectKeys,"object keys remain unique");Check(repeatedNativeIds,"real repeated native IDs preserved without loss");}
        return count+" independent presentation-source assertions passed; originals="+packages+" packages/"+slides+" slides. No PowerPoint rendering or insertion was executed.";
    }
}