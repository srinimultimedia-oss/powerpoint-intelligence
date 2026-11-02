using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading;
using PowerPointIntelligence.Library;
public static class PiidxEdgeTests
{
    public static string Run(string output)
    {
        Directory.CreateDirectory(output);string root=Path.Combine(output,"large");Directory.CreateDirectory(root);string path=Path.Combine(root,"large.pptx");
        string p="http://schemas.openxmlformats.org/presentationml/2006/main",r="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        using(var file=new FileStream(path,FileMode.CreateNew))using(var zip=new ZipArchive(file,ZipArchiveMode.Create))
        {
            var ids=new StringBuilder();var rels=new StringBuilder("<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>");
            for(int i=0;i<1001;i++)
            {
                ids.Append("<p:sldId id='").Append(256+i).Append("' r:id='r").Append(i).Append("'/>");rels.Append("<Relationship Id='r").Append(i).Append("' Type='").Append(r).Append("/slide' Target='slides/s").Append(i).Append(".xml'/>");
                using(var writer=new StreamWriter(zip.CreateEntry("ppt/slides/s"+i+".xml").Open(),new UTF8Encoding(false)))writer.Write("<p:sld xmlns:p='"+p+"'><p:cSld><p:spTree/></p:cSld></p:sld>");
            }
            using(var writer=new StreamWriter(zip.CreateEntry("ppt/presentation.xml").Open(),new UTF8Encoding(false)))writer.Write("<p:presentation xmlns:p='"+p+"' xmlns:r='"+r+"'><p:sldIdLst>"+ids+"</p:sldIdLst><p:sldSz cx='9144000' cy='5143500'/></p:presentation>");
            using(var writer=new StreamWriter(zip.CreateEntry("ppt/_rels/presentation.xml.rels").Open(),new UTF8Encoding(false)))writer.Write(rels+"</Relationships>");
        }
        var catalog=PiidxCatalog.OpenOrCreate(Path.Combine(output,"large.piidx"));string id=catalog.Register(root,"Large",true);var first=catalog.Refresh(id,CancellationToken.None);
        if(first.SlidesIndexed!=1001)throw new Exception("Large source was not completely indexed.");
        var last=catalog.Search(new LibrarySearchRequest{Offset=1000,Limit=10},CancellationToken.None);if(last.Total!=1001||last.Slides.Count!=1||last.Slides[0].Position!=1001)throw new Exception("Large source pagination failed.");
        catalog.VerifyIntegrity();if(catalog.Refresh(id,CancellationToken.None).FilesUpdated!=0)throw new Exception("Large source reparsed unexpectedly.");
        string longRoot=Path.Combine(output,"long-path");while(longRoot.Length<275)longRoot=Path.Combine(longRoot,"directory-component");Directory.CreateDirectory(longRoot);
        PiidxSourceTests.Fixture(Path.Combine(longRoot,"unicode-日本語.pptx"));var longCatalog=PiidxCatalog.OpenOrCreate(Path.Combine(output,"long-path.piidx"));string longId=longCatalog.Register(longRoot,"Long path",true);longCatalog.Refresh(longId,CancellationToken.None);
        if(longCatalog.Search(new LibrarySearchRequest(),CancellationToken.None).Total!=2)throw new Exception("Long Unicode source path failed.");
        return "4 PIIDX edge assertions passed; large/long-path extraction does not establish native Office acceptance.";
    }
}