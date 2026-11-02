using System;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Collections.Generic;
internal static class TemplateFidelityTests
{
    internal static void Run()
    {
        int count=0;Action<bool,string> check=(ok,name)=>{if(!ok)throw new Exception("FAIL fidelity policy: "+name);count++;};
        using(var reference=new Bitmap(400,400)){
            using(var g=Graphics.FromImage(reference)){g.Clear(Color.White);g.FillRectangle(Brushes.Gray,200,20,20,360);}
            using(var a=new Bitmap(reference))check(TemplateFidelity.Compare(reference,a).Exact,"identical pixels exact");
            using(var a=new Bitmap(reference)){for(int i=0;i<8;i++)a.SetPixel(200,30+i,Color.FromArgb(138,138,138));var r=TemplateFidelity.Compare(reference,a);check(r.Accepted&&!r.Exact&&r.ChangedPixels==8&&r.MaxChannelDelta==10,"eight subtle edge changes accepted with explicit metrics");}
            using(var a=new Bitmap(reference)){for(int i=0;i<9;i++)a.SetPixel(200,30+i,Color.FromArgb(138,138,138));check(!TemplateFidelity.Compare(reference,a).Accepted,"ninth pixel rejected");}
            using(var a=new Bitmap(reference)){a.SetPixel(200,30,Color.FromArgb(139,139,139));check(!TemplateFidelity.Compare(reference,a).Accepted,"delta eleven rejected");}
            using(var a=new Bitmap(reference)){a.SetPixel(200,30,Color.FromArgb(254,128,128,128));check(!TemplateFidelity.Compare(reference,a).Accepted,"alpha change rejected");}
            using(var a=new Bitmap(reference)){a.SetPixel(20,30,Color.FromArgb(250,250,250));check(!TemplateFidelity.Compare(reference,a).Accepted,"flat interior edit rejected");}
            using(var a=new Bitmap(400,399))check(!TemplateFidelity.Compare(reference,a).Accepted,"dimension mismatch rejected");
            using(var a=new Bitmap(reference)){using(var g=Graphics.FromImage(a)){g.Clear(Color.FromArgb(254,254,254));}check(!TemplateFidelity.Compare(reference,a).Accepted,"global one-level recolor rejected");}
            using(var a=new Bitmap(reference)){using(var g=Graphics.FromImage(a)){g.Clear(Color.White);g.FillRectangle(Brushes.Gray,201,20,20,360);}check(!TemplateFidelity.Compare(reference,a).Accepted,"one-pixel translation rejected");}
            using(var a=new Bitmap(reference)){a.SetPixel(200,30,Color.Black);check(!TemplateFidelity.Compare(reference,a).Accepted,"small high-contrast content change rejected");}
        }
        using(var a=new Bitmap(100,100))using(var b=new Bitmap(100,100)){using(var g=Graphics.FromImage(a))g.Clear(Color.White);using(var g=Graphics.FromImage(b))g.Clear(Color.White);a.SetPixel(50,50,Color.Gray);b.SetPixel(50,50,Color.FromArgb(130,130,130));check(!TemplateFidelity.Compare(a,b).Accepted,"fraction limit enforced for small images");}
        count+=GeometryTests();
        Console.WriteLine(count+" bounded antialias policy assertions passed; native formatting is not inferred from pixels.");
    }
    private static int GeometryTests()
    {
        int count=0;Action<bool,string> check=(ok,name)=>{if(!ok)throw new Exception("FAIL persisted geometry: "+name);count++;};Action<Action,string> reject=(call,name)=>{bool failed=false;try{call();}catch(Exception error) when(error is InvalidDataException||error is XmlException){failed=true;}check(failed,name);};
        const string p="http://schemas.openxmlformats.org/presentationml/2006/main",a="http://schemas.openxmlformats.org/drawingml/2006/main";
        string xml="<p:sld xmlns:p='"+p+"' xmlns:a='"+a+"'><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:grpSp><p:nvGrpSpPr/><p:grpSpPr><a:xfrm rot='1'><a:off x='5' y='6'/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr/><p:spPr><a:xfrm flipH='1'/><a:custGeom><a:gdLst><a:gd name='g1' fmla='val 7'/></a:gdLst><a:pathLst><a:path w='100' h='100'><a:moveTo><a:pt x='1' y='2'/></a:moveTo><a:cubicBezTo><a:pt x='3' y='4'/><a:pt x='5' y='6'/><a:pt x='7' y='8'/></a:cubicBezTo><a:close/></a:path></a:pathLst></a:custGeom></p:spPr></p:sp></p:grpSp></p:spTree></p:cSld></p:sld>";
        string expected=TemplateFidelity.GeometryXml(xml);check(expected.Contains("cubicBezTo")&&expected.Contains("val 7"),"curves and guides retained without COM node getters");
        check(expected==TemplateFidelity.GeometryXml(xml.Replace("w='100' h='100'","h='100' w='100'")),"attribute ordering does not change semantic geometry");
        check(expected==TemplateFidelity.GeometryXml(xml.Replace("a:","d:").Replace("xmlns:a=","xmlns:d=")),"namespace prefixes do not change semantic geometry");
        check(expected!=TemplateFidelity.GeometryXml(xml.Replace("x='7' y='8'","x='9' y='8'")),"changed curve control point rejected exactly");
        check(expected!=TemplateFidelity.GeometryXml(xml.Replace("cubicBezTo","lnTo")),"path command change rejected exactly");
        check(expected!=TemplateFidelity.GeometryXml(xml.Replace("val 7","val 8")),"guide change rejected exactly");
        check(expected!=TemplateFidelity.GeometryXml(xml.Replace("rot='1'","rot='2'")),"group transform change rejected exactly");
        check(expected!=TemplateFidelity.GeometryXml(xml.Replace("flipH='1'","flipH='0'")),"flip change rejected exactly");
        reject(()=>TemplateFidelity.GeometryXml(xml.Replace("<p:sp>","<p:unknown>").Replace("</p:sp>","</p:unknown>")),"unsupported containers fail closed");
        reject(()=>TemplateFidelity.GeometryXml("<!DOCTYPE p:sld [<!ENTITY e SYSTEM 'file:///never-read'>]>"+xml),"DTD prohibited");
        reject(()=>TemplateFidelity.GeometryXml("<x/>"),"missing slide tree refused");
        string transform="<a:xfrm><a:off x='10' y='20'/><a:ext cx='30' cy='40'/></a:xfrm>";
        Func<string,string,string,string> part=(root,idxValue,xfrm)=>"<p:"+root+" xmlns:p='"+p+"' xmlns:a='"+a+"'><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:sp><p:nvSpPr><p:nvPr><p:ph type='title' idx='"+idxValue+"'/></p:nvPr></p:nvSpPr><p:spPr>"+xfrm+"</p:spPr></p:sp></p:spTree></p:cSld></p:"+root+">";
        string inheritedSlide=part("sld","7",""),explicitSlide=part("sld","7",transform),layout=part("sldLayout","7",""),master=part("sldMaster","91",transform),resolved=TemplateFidelity.GeometryXml(explicitSlide);
        check(TemplateFidelity.EffectiveGeometryXml(inheritedSlide,layout,master)==resolved,"master title inheritance equals identical explicit slide geometry");
        string layoutTransform=transform.Replace("x='10'","x='11'");check(TemplateFidelity.EffectiveGeometryXml(inheritedSlide,part("sldLayout","7",layoutTransform),master)==TemplateFidelity.GeometryXml(part("sld","7",layoutTransform)),"layout geometry overrides master");
        check(TemplateFidelity.EffectiveGeometryXml(explicitSlide,part("sldLayout","7",layoutTransform),master)==resolved,"explicit slide geometry overrides inherited values");
        check(TemplateFidelity.EffectiveGeometryXml(inheritedSlide,layout,master.Replace("x='10'","x='12'"))!=resolved,"changed inherited coordinate fails exact geometry");
        reject(()=>TemplateFidelity.EffectiveGeometryXml(inheritedSlide,part("sldLayout","8",""),master),"unmatched layout index fails closed");
        reject(()=>TemplateFidelity.EffectiveGeometryXml(inheritedSlide,layout,master.Replace("type='title'","type='body'")),"wrong master type fails closed");
        string duplicateLayout=layout.Replace("</p:spTree>",layout.Substring(layout.IndexOf("<p:sp>"),layout.IndexOf("</p:sp>")-layout.IndexOf("<p:sp>")+7)+"</p:spTree>");reject(()=>TemplateFidelity.EffectiveGeometryXml(inheritedSlide,duplicateLayout,master),"ambiguous layout index fails closed");
        reject(()=>TemplateFidelity.EffectiveGeometryXml(inheritedSlide,layout,part("sldMaster","91","")),"unresolved geometry is not replaced with a default box");
        string file=Path.Combine(Path.GetTempPath(),"ppi-geometry-contract-"+Guid.NewGuid().ToString("N")+".pptx");
        try{
            using(var archive=ZipFile.Open(file,ZipArchiveMode.Create)){
                Action<string,string> add=(name,body)=>{using(var stream=archive.CreateEntry(name).Open())using(var writer=new StreamWriter(stream,new UTF8Encoding(false)))writer.Write(body);};
                add("ppt/presentation.xml","<p:presentation xmlns:p='"+p+"' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><p:sldIdLst><p:sldId id='256' r:id='chosen'/></p:sldIdLst></p:presentation>");
                add("ppt/_rels/presentation.xml.rels","<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='chosen' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide' Target='slides/slide42.xml'/></Relationships>");
                add("ppt/slides/slide42.xml",xml);
            }
            check(TemplateFidelity.PackageGeometry(file,1)==expected,"slide resolved by relationship rather than filename/index assumption");reject(()=>TemplateFidelity.PackageGeometry(file,0),"invalid slide index refused");reject(()=>TemplateFidelity.PackageGeometry(file,2),"missing slide index refused");
            using(var archive=ZipFile.Open(file,ZipArchiveMode.Update)){
                archive.GetEntry("ppt/slides/slide42.xml").Delete();Action<string,string> add=(name,body)=>{using(var writer=new StreamWriter(archive.CreateEntry(name).Open()))writer.Write(body);};add("ppt/slides/slide42.xml",inheritedSlide);
                add("ppt/slides/_rels/slide42.xml.rels","<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='layout' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout' Target='../slideLayouts/slideLayout9.xml'/></Relationships>");add("ppt/slideLayouts/slideLayout9.xml",layout);
                add("ppt/slideLayouts/_rels/slideLayout9.xml.rels","<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='master' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster' Target='../slideMasters/slideMaster11.xml'/></Relationships>");add("ppt/slideMasters/slideMaster11.xml",master);
            }
            check(TemplateFidelity.PackageGeometry(file,1)==resolved,"package resolves slide-layout-master relationships");
            using(var archive=ZipFile.Open(file,ZipArchiveMode.Update)){archive.GetEntry("ppt/slideMasters/slideMaster11.xml").Delete();using(var writer=new StreamWriter(archive.CreateEntry("ppt/slideMasters/slideMaster11.xml").Open()))writer.Write(master.Replace("x='10'","x='12'"));}
            check(TemplateFidelity.PackageGeometry(file,1)!=resolved,"package inherited geometry changes remain strict failures");
            using(var archive=ZipFile.Open(file,ZipArchiveMode.Update)){var entry=archive.GetEntry("ppt/_rels/presentation.xml.rels");entry.Delete();using(var writer=new StreamWriter(archive.CreateEntry("ppt/_rels/presentation.xml.rels").Open()))writer.Write("<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='chosen' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide' TargetMode='External' Target='https://never-fetch.example/slide.xml'/></Relationships>");}
            reject(()=>TemplateFidelity.PackageGeometry(file,1),"external relationships refused without fetching");
        }finally{if(File.Exists(file))File.Delete(file);}
        return count;
    }
    internal static void Replay(string folder,string output)
    {
        Directory.CreateDirectory(output);var rows=new List<string>{"Index,Stage,SameDimensions,ExactPixels,WithinTolerance,ChangedPixels,MaxChannelDelta,AlphaChanges,NonEdgeChanges,AllowedPixels"};int passed=0;
        for(int i=0;i<16;i++)foreach(var stage in new[]{"inserted","source-repeat","reopened"}){var r=TemplateFidelity.Files(Path.Combine(folder,"original-"+i+"-source.png"),Path.Combine(folder,"original-"+i+"-"+stage+".png"));rows.Add(i+","+stage+","+r.Csv);if(!r.Accepted)throw new Exception("Saved raster evidence outside policy: "+i+" "+stage+" "+r.Csv);passed++;}
        File.WriteAllLines(Path.Combine(output,"raster-replay.csv"),rows);Console.WriteLine("PASS "+passed+" saved raster comparisons. This is an offline replay, not a new native Office/formatting pass.");
    }
}
