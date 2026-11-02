using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.IO.Compression;
using System.Text;
using PowerPointIntelligence.AddIn;
internal static class LibraryBrowserTests
{
    private static int passed;
    private static void Check(bool value,string name){if(!value)throw new Exception("Library regression: "+name);passed++;}
    internal static void Run()
    {
        Check(TemplateCatalog.Search(null,"Owner").Any(),"Search includes editable field defaults");
        Check(TemplateCatalog.Search(null,"accountable leader").Any(t=>t.Category=="Organization"),"Multiword search includes field content");
        Check(TemplateCatalog.Search("Teams","research group").Count()==1,"Category and multiword query");
        Check(!TemplateCatalog.Search(null,"no-such-template-8473").Any(),"Empty search result");
        Check(TemplateCatalog.Search(null,"  RESEARCH   group  ").Select(t=>t.Id).SequenceEqual(TemplateCatalog.Search(null,"research group").Select(t=>t.Id)),"Whitespace and case normalization");
        int matches=TemplateCatalog.Search(null,"owner").Count();Check(TemplateCatalog.Categories.Sum(c=>TemplateCatalog.Search(c,"owner").Count())==matches,"Category counts partition matches");
        var first=TemplateCatalog.Items[0];var second=TemplateCatalog.Items[1];var store=new LibraryDrafts();var edit=store.Load(first);edit["title"]="A user draft";edit["subtitle"]="";store.Save(first,edit);edit["title"]="mutated outside store";
        Check(store.Load(first)["title"]=="A user draft","Save copies draft input");
        Check(store.Load(second)["title"]==second.Name,"Templates have isolated drafts");
        var copy=store.Load(first);copy["title"]="outside mutation";Check(store.Load(first)["title"]=="A user draft","Load returns copy");
        Check(store.Load(first)["subtitle"]=="","Empty fields retained");
        TemplateCatalog.Search(null,"process").ToList();first.Build("03");Check(store.Load(first)["title"]=="A user draft","Search and palette projection do not lose draft");
        bool rejected=false;try{store.Save(first,new Dictionary<string,string>{{"title",new String('x',2001)}});}catch(System.IO.InvalidDataException){rejected=true;}Check(rejected,"Field length bound");
        Check(store.Load(first)["title"]=="A user draft","Invalid draft leaves saved values intact");
        rejected=false;try{store.Save(first,new Dictionary<string,string>{{"unknown","value"}});}catch(System.IO.InvalidDataException){rejected=true;}Check(rejected,"Unknown fields rejected");
        foreach(var size in new[]{new[]{960f,540f},new[]{720f,540f}}){
            var widths=new[]{140f,190f,110f,170f,220f,100f};var heights=new[]{90f,75f,130f,60f,105f,120f};var plan=GridPlanner.Plan(size[0],size[1],widths,heights);
            Check(plan.Length==widths.Length,"Grid places every shape");
            for(int i=0;i<plan.Length;i++){Check(plan[i].X>=40&&plan[i].Y>=60&&plan[i].X+widths[i]<=size[0]-40&&plan[i].Y+heights[i]<=size[1]-40,"Grid remains inside canvas");for(int j=i+1;j<plan.Length;j++)Check(plan[i].X+widths[i]<=plan[j].X||plan[j].X+widths[j]<=plan[i].X||plan[i].Y+heights[i]<=plan[j].Y||plan[j].Y+heights[j]<=plan[i].Y,"Heterogeneous objects do not overlap");}
        }
        foreach(var widths in new[]{new[]{1000f},new[]{Single.NaN},new[]{-1f}}){rejected=false;try{GridPlanner.Plan(960,540,widths,new[]{100f});}catch(InvalidOperationException){rejected=true;}Check(rejected,"Impossible dimensions rejected before COM");}
        rejected=false;try{GridPlanner.Plan(960,540,Enumerable.Repeat(200f,100).ToArray(),Enumerable.Repeat(200f,100).ToArray());}catch(InvalidOperationException){rejected=true;}Check(rejected,"Overflowing many-row layout rejected");
        Check(TemplateTheme.Binding("#FFFFFF",TemplatePalettes.Get(),true).Item1==14,"Light text follows destination background color role");
        Check(TemplateTheme.Binding(TemplatePalettes.Get().Colors[0],TemplatePalettes.Get(),false).Item1==5,"Accent follows theme accent one");
        Check(TemplateTheme.Binding(TemplatePalettes.Get().Colors[0],TemplatePalettes.Get(),true).Item1==13,"Body text follows destination text role");
        Check(TemplateTheme.Binding(TemplatePalettes.Get().Colors[6],TemplatePalettes.Get(),false).Item2>0,"Muted destination accent tint");
        LocalAssets();
        LayoutRules();
        using(var rgb=new System.Drawing.Bitmap(3,2,System.Drawing.Imaging.PixelFormat.Format24bppRgb))using(var argb=new System.Drawing.Bitmap(3,2,System.Drawing.Imaging.PixelFormat.Format32bppArgb)){
            rgb.SetResolution(54,54);argb.SetResolution(96,96);for(int y=0;y<2;y++)for(int x=0;x<3;x++){var color=System.Drawing.Color.FromArgb(255,x*70,y*80,35);rgb.SetPixel(x,y,color);argb.SetPixel(x,y,color);}
            Check(TemplateLibraryTests.Pixels(rgb)==TemplateLibraryTests.Pixels(argb),"Exact pixel comparison ignores DPI and storage format");argb.SetPixel(1,1,System.Drawing.Color.Red);Check(TemplateLibraryTests.Pixels(rgb)!=TemplateLibraryTests.Pixels(argb),"Exact pixel comparison detects a one-pixel change");
        }
        Console.WriteLine("PASS "+passed+" library search/drafts and bounds-aware layout assertions. Native UI not tested.");
    }
    private static void Part(ZipArchive zip,string name,string text){using(var writer=new StreamWriter(zip.CreateEntry(name).Open(),new UTF8Encoding(false)))writer.Write(text);}
    private static void LocalAssets()
    {
        string folder=Path.Combine(Path.GetTempPath(),"ppi-library-fixtures-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(folder);string file=Path.Combine(folder,"Quarterly.pptx");
        using(var zip=ZipFile.Open(file,ZipArchiveMode.Create)){
            Part(zip,"ppt/presentation.xml","<p:presentation xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><p:sldIdLst><p:sldId id='256' r:id='rId2'/><p:sldId id='257' r:id='rId1'/></p:sldIdLst><p:sldSz cx='12192000' cy='6858000'/></p:presentation>");
            Part(zip,"ppt/_rels/presentation.xml.rels","<Relationships><Relationship Id='rId1' Target='slides/slide1.xml'/><Relationship Id='rId2' Target='slides/slide2.xml'/></Relationships>");
            Part(zip,"ppt/slides/slide1.xml","<sld><t>Second slide</t></sld>");Part(zip,"ppt/slides/slide2.xml","<sld><t>Revenue and margin</t></sld>");Part(zip,"docProps/core.xml","<core><creator>Research Team</creator></core>");
        }
        var slides=LocalSlideIndex.ReadFile(file);Check(slides.Count==2&&slides[0].Text=="Revenue and margin","Index follows presentation relationship order, not filenames");
        Check(slides[0].Width==960&&slides[0].Height==540,"EMU dimensions converted to points");
        Check(LocalSlideIndex.Filter(slides,"margin quarterly","research",null,"16:9",null,false).Count()==1,"Text, filename, author and aspect filters combine");
        Check(!LocalSlideIndex.Filter(slides,"",null,null,"4:3",null,false).Any(),"Aspect filter excludes other shapes");
        Check(!LocalSlideIndex.Filter(slides,"",null,null,"All",DateTime.UtcNow.AddDays(1),false).Any(),"Modified-date filter");
        Check(LocalSlideIndex.Filter(slides.Concat(slides),"",null,null,"All",null,true).Count()==2,"Exact repeated assets hidden");
        Check(LocalSlideIndex.Filter(slides.Concat(slides),"",null,null,"All",null,false).Count()==4,"Duplicate display selectable");
        Check(slides[0].Signature!=slides[1].Signature,"Distinct slides remain distinct");
        string cachedFolder;string cache=LocalLibraryCache.Encode(folder,slides);var cached=LocalLibraryCache.Decode(cache,out cachedFolder);Check(cachedFolder==folder&&cached.Count==2&&cached[0].Text==slides[0].Text,"Versioned local index round trip");
        bool invalidCache=false;try{LocalLibraryCache.Decode(cache.Replace("version=\"1\"","version=\"2\""),out cachedFolder);}catch(InvalidDataException){invalidCache=true;}Check(invalidCache,"Unknown cached index version rejected");
        invalidCache=false;try{LocalLibraryCache.Decode(cache.Replace(file,Path.Combine(Path.GetDirectoryName(folder),"outside.pptx")),out cachedFolder);}catch(InvalidDataException){invalidCache=true;}Check(invalidCache,"Cached files cannot escape selected folder");
        string bad=Path.Combine(folder,"bad.pptx");using(var zip=ZipFile.Open(bad,ZipArchiveMode.Create)){Part(zip,"ppt/presentation.xml","<!DOCTYPE p [<!ENTITY e SYSTEM 'file:///never-read'>]><p>&e;</p>");}
        bool rejected=false;try{LocalSlideIndex.ReadFile(bad);}catch(System.Xml.XmlException){rejected=true;}Check(rejected,"DTD and external entities rejected");
        rejected=false;try{LocalSlideIndex.ReadFile(Path.Combine(folder,"macro.pptm"));}catch(InvalidDataException){rejected=true;}Check(rejected,"Macro-enabled library file extensions rejected");
        using(var zip=ZipFile.Open(file,ZipArchiveMode.Update)){zip.GetEntry("ppt/slides/slide2.xml").Delete();Part(zip,"ppt/slides/slide2.xml","<sld><t>Changed revenue</t></sld>");}
        Check(LocalSlideIndex.ReadFile(file)[0].Signature!=slides[0].Signature,"Refresh detects changed source signature");
        string nested=Path.Combine(folder,"Diagrams","Cubes");Directory.CreateDirectory(nested);string nestedFile=Path.Combine(nested,"Cubes.potx");System.IO.File.Copy(file,nestedFile);
        Check(LocalSlideIndex.Discover(folder,true).Contains(nestedFile),"Recursive library discovers nested source collections");
        Check(!LocalSlideIndex.Discover(folder,false).Contains(nestedFile),"Nonrecursive mode excludes nested files");
        Check(LocalSlideIndex.IsWithin(folder,nestedFile)&&!LocalSlideIndex.IsWithin(folder,folder+"-other"+Path.DirectorySeparatorChar+"wrong.pptx"),"Recursive root boundary rejects sibling prefixes");
        string nestedCache=LocalLibraryCache.Encode(folder,LocalSlideIndex.ReadFile(nestedFile));Check(LocalLibraryCache.Decode(nestedCache,out cachedFolder).Count==2,"Nested original cache stays within selected root");
        bool wrongSize=false;try{TemplateSources.RequireSameSize(slides[0],720,540);}catch(InvalidOperationException){wrongSize=true;}Check(wrongSize,"Keep-original mode rejects dimension changes before Office access");TemplateSources.RequireSameSize(slides[0],960,540);Check(true,"Keep-original accepts identical source dimensions");

    }

    private static void LayoutRules()
    {
        string id=Guid.NewGuid().ToString("D");var rect=new System.Drawing.RectangleF(100,70,200,90);var own=new System.Drawing.SizeF(50,30);
        foreach(var pair in new[]{Tuple.Create("L",100f),Tuple.Create("R",250f),Tuple.Create("C",175f),Tuple.Create("T",70f),Tuple.Create("B",130f),Tuple.Create("M",100f)}){var rule=new LayoutRule {Leader=id,LeaderEdge=pair.Item1,OwnEdge=pair.Item1,Offset=0};Check(Math.Abs(LayoutRule.Decode(rule.Encode()).Position(rect,own)-pair.Item2)<.01,"Alignment role and wire round trip "+pair.Item1);}
        var flow=new LayoutRule {Leader=id,LeaderEdge="R",OwnEdge="L",Offset=16};Check(flow.Position(rect,own)==316,"Content growth preserves following gap");rect.Width=300;Check(flow.Position(rect,own)==416,"Following position responds to resized anchor");
        var edge=new LayoutRule {Leader="SLIDE",LeaderEdge="R",OwnEdge="R",Offset=-40};Check(edge.Position(new System.Drawing.RectangleF(0,0,960,540),own)==870,"Right edge lock keeps slide margin");Check(edge.Position(new System.Drawing.RectangleF(0,0,720,540),own)==630,"Edge lock responds to slide resize");
        foreach(string bad in new[]{"v2|SLIDE|L|L|0","v1|bad-id|L|L|0","v1|SLIDE|L|T|0","v1|SLIDE|R|R|NaN","v1|SLIDE|R|R|Infinity","v1|SLIDE|R|R|999999"}){bool rejected=false;try{LayoutRule.Decode(bad);}catch(InvalidDataException){rejected=true;}Check(rejected,"Invalid persisted relationship rejected");}
    }

}
