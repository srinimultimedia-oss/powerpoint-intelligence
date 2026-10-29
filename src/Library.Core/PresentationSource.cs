using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Xml;
using System.Xml.Linq;

namespace PowerPointIntelligence.Library
{
    internal sealed class IndexedObject
    {
        internal string Key,ParentKey,NativeId,ParentId,Kind,Name,Text,XmlFingerprint;
        // Coordinates remain in the native parent coordinate system; grouped positions are not falsely described as absolute.
        internal long X,Y,Width,Height,Rotation;
        internal bool HasTransform,FlipH,FlipV;
    }
    internal sealed class IndexedSlide
    {
        internal uint NativeId;
        internal int Position;
        internal string PartUri,Title,Text,Category,Layout,Master,ContentFingerprint;
        internal readonly List<IndexedObject> Objects=new List<IndexedObject>();
        internal readonly List<string> Tags=new List<string>();
    }
    internal sealed class IndexedPresentation
    {
        internal string SourcePath,FileType,ContentHash,Author,Keywords;
        internal long Size;
        internal DateTime ModifiedUtc;
        internal double Width,Height;
        internal readonly List<IndexedSlide> Slides=new List<IndexedSlide>();
    }
    internal static class PresentationSource
    {
        const long MaximumFileBytes=1024L*1024*1024;
        const int MaximumPartChars=40*1024*1024;
        static readonly XNamespace PackageRels="http://schemas.openxmlformats.org/package/2006/relationships";
        static readonly XNamespace Mc="http://schemas.openxmlformats.org/markup-compatibility/2006";
        static readonly HashSet<string> SupportedShapes=new HashSet<string>{"sp","pic","graphicFrame","cxnSp","grpSp","contentPart"};
        internal static string Fingerprint(string path,CancellationToken cancellation)
        {
            using(var hash=SHA256.Create())using(var file=new FileStream(path,FileMode.Open,FileAccess.Read,FileShare.ReadWrite|FileShare.Delete))
            {
                var buffer=new byte[65536];int read;
                while((read=file.Read(buffer,0,buffer.Length))>0){cancellation.ThrowIfCancellationRequested();hash.TransformBlock(buffer,0,read,null,0);}
                cancellation.ThrowIfCancellationRequested();hash.TransformFinalBlock(new byte[0],0,0);return Hex(hash.Hash);
            }
        }
        static string Hex(byte[] bytes){return BitConverter.ToString(bytes).Replace("-","");}
        static string HashText(string value){using(var hash=SHA256.Create())return Hex(hash.ComputeHash(Encoding.UTF8.GetBytes(value)));}
        internal static IEnumerable<string> Discover(string root,bool recursive,CancellationToken cancellation)
        {
            root=Path.GetFullPath(root);if(!Directory.Exists(root))throw new DirectoryNotFoundException("Library root is unavailable.");
            if((File.GetAttributes(root)&FileAttributes.ReparsePoint)!=0)throw new InvalidDataException("Library root must be a direct directory, not a reparse point.");
            var pending=new Stack<string>();pending.Push(root);int directories=0,files=0;
            while(pending.Count>0)
            {
                cancellation.ThrowIfCancellationRequested();string folder=pending.Pop();if(++directories>10000)throw new InvalidDataException("Library directory limit exceeded.");
                foreach(string file in Directory.EnumerateFiles(folder))
                {
                    cancellation.ThrowIfCancellationRequested();if(++files>100000)throw new InvalidDataException("Library file limit exceeded.");
                    if((File.GetAttributes(file)&FileAttributes.ReparsePoint)!=0)continue;
                    string ext=Path.GetExtension(file);if(String.Equals(ext,".pptx",StringComparison.OrdinalIgnoreCase)||String.Equals(ext,".potx",StringComparison.OrdinalIgnoreCase))yield return Path.GetFullPath(file);
                }
                if(recursive)foreach(string child in Directory.EnumerateDirectories(folder))if((File.GetAttributes(child)&FileAttributes.ReparsePoint)==0)pending.Push(child);
            }
        }
        static XDocument Xml(ZipArchive zip,string part,bool required)
        {
            var entry=zip.GetEntry(part);if(entry==null){if(required)throw new InvalidDataException("A required presentation part is missing.");return null;}
            if(entry.Length>MaximumPartChars)throw new InvalidDataException("Presentation XML part exceeds the configured bound.");
            using(var stream=entry.Open())using(var reader=XmlReader.Create(stream,new XmlReaderSettings{DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=MaximumPartChars}))return XDocument.Load(reader,LoadOptions.None);
        }
        static string RelationshipPart(string part){int slash=part.LastIndexOf('/');return part.Substring(0,slash+1)+"_rels/"+part.Substring(slash+1)+".rels";}
        static string Resolve(string from,string target)
        {
            if(String.IsNullOrWhiteSpace(target)||target.IndexOf('\\')>=0||target.IndexOf(':')>=0)throw new InvalidDataException("Unsupported package relationship target.");
            var uri=new Uri(new Uri("https://piidx.invalid/"+from),target);
            if(uri.Scheme!="https"||uri.Host!="piidx.invalid"||uri.Query.Length>0||uri.Fragment.Length>0)throw new InvalidDataException("External package reference is not indexed.");
            return Uri.UnescapeDataString(uri.AbsolutePath.TrimStart('/'));
        }
        static string Related(XDocument rels,string from,string suffix,string id,bool required)
        {
            var matches=rels==null?new XElement[0]:rels.Elements(PackageRels+"Relationships").Elements(PackageRels+"Relationship").Where(e=>((string)e.Attribute("Type")??"").EndsWith("/"+suffix,StringComparison.Ordinal)&&(id==null||(string)e.Attribute("Id")==id)).ToArray();
            if(matches.Length==0){if(required)throw new InvalidDataException("Required source relationship is missing.");return null;}
            if(matches.Length!=1||(string)matches[0].Attribute("TargetMode")=="External")throw new InvalidDataException("Ambiguous or external source relationship.");
            return Resolve(from,(string)matches[0].Attribute("Target"));
        }
        static string Text(XContainer node,XNamespace drawing)
        {
            var text=new StringBuilder();foreach(var element in node.Descendants(drawing+"t")){if(text.Length>0)text.Append(' ');text.Append(element.Value);if(text.Length>640000)throw new InvalidDataException("Slide text exceeds the configured bound.");}return text.ToString();
        }
        static long Number(XElement node,string name,bool required=false)
        {
            string text=node==null?null:(string)node.Attribute(name);long number;
            if(text==null&&!required)return 0;
            if(!Int64.TryParse(text,NumberStyles.Integer,CultureInfo.InvariantCulture,out number)||number < -1000000000000L||number >1000000000000L)throw new InvalidDataException("Invalid source geometry.");return number;
        }
        static string NativeObjectId(XElement shape,XNamespace p)
        {
            var nv=shape.Elements().FirstOrDefault(e=>e.Name.Namespace==p&&e.Name.LocalName.StartsWith("nv",StringComparison.Ordinal));
            var property=nv==null?null:nv.Descendants(p+"cNvPr").FirstOrDefault();
            uint id;if(property==null||!UInt32.TryParse((string)property.Attribute("id"),NumberStyles.None,CultureInfo.InvariantCulture,out id))throw new InvalidDataException("Source object identity is missing or invalid; kind="+shape.Name.LocalName+"; present="+(property!=null)+".");return id.ToString(CultureInfo.InvariantCulture);
        }
        static string ObjectKey(XElement shape)
        {
            return String.Join("/",shape.AncestorsAndSelf().Reverse().Select(e=>e.Name.LocalName+"["+(e.ElementsBeforeSelf(e.Name).Count()+1).ToString(CultureInfo.InvariantCulture)+"]"));
        }
        static string Name(XDocument xml,XNamespace p,string fallback)
        {
            var content=xml==null||xml.Root==null?null:xml.Root.Element(p+"cSld");return (content==null?null:(string)content.Attribute("name"))??fallback??"";
        }
        static void SelectCompatibilityBranches(XDocument slide)
        {
            foreach(var alternate in slide.Descendants(Mc+"AlternateContent").Reverse().ToArray())
            {
                var branch=alternate.Element(Mc+"Fallback")??alternate.Elements(Mc+"Choice").FirstOrDefault();
                if(branch==null)throw new InvalidDataException("Empty alternate-content branch.");alternate.ReplaceWith(branch.Nodes().ToArray());
            }
        }
        internal static IndexedPresentation Read(string path,CancellationToken cancellation)
        {
            cancellation.ThrowIfCancellationRequested();path=Path.GetFullPath(path);string type=Path.GetExtension(path).ToLowerInvariant();
            if(type!=".pptx"&&type!=".potx")throw new InvalidDataException("Library sources must be PPTX or POTX presentations.");
            var info=new FileInfo(path);if(info.Length>MaximumFileBytes)throw new InvalidDataException("Presentation exceeds the configured file bound.");
            string hash=Fingerprint(path,cancellation);
            var result=new IndexedPresentation{SourcePath=path,FileType=type,Size=info.Length,ModifiedUtc=info.LastWriteTimeUtc,ContentHash=hash,Author="",Keywords=""};
            using(var file=new FileStream(path,FileMode.Open,FileAccess.Read,FileShare.ReadWrite|FileShare.Delete))using(var zip=new ZipArchive(file,ZipArchiveMode.Read))
            {
                if(zip.Entries.Count>100000)throw new InvalidDataException("Package member limit exceeded.");
                var presentation=Xml(zip,"ppt/presentation.xml",true);
                if(presentation.Root==null||presentation.Root.Name.LocalName!="presentation")throw new InvalidDataException("Source is not a presentation.");
                XNamespace p=presentation.Root.Name.Namespace;
                bool strict=p=="http://purl.oclc.org/ooxml/presentationml/main";
                if(!strict&&p!="http://schemas.openxmlformats.org/presentationml/2006/main")throw new InvalidDataException("Unknown presentation namespace.");
                XNamespace drawing=strict?"http://purl.oclc.org/ooxml/drawingml/main":"http://schemas.openxmlformats.org/drawingml/2006/main";
                XNamespace relation=strict?"http://purl.oclc.org/ooxml/officeDocument/relationships":"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                var size=presentation.Root.Element(p+"sldSz");long width=Number(size,"cx",true),height=Number(size,"cy",true);
                if(width<=0||height<=0)throw new InvalidDataException("Slide dimensions are invalid.");result.Width=width/12700.0;result.Height=height/12700.0;
                var properties=Xml(zip,"docProps/core.xml",false);
                if(properties!=null){result.Author=properties.Descendants(XName.Get("creator","http://purl.org/dc/elements/1.1/")).Select(e=>e.Value).FirstOrDefault()??"";result.Keywords=properties.Descendants(XName.Get("keywords","http://schemas.openxmlformats.org/package/2006/metadata/core-properties")).Select(e=>e.Value).FirstOrDefault()??"";}
                var rels=Xml(zip,"ppt/_rels/presentation.xml.rels",true);var list=presentation.Root.Element(p+"sldIdLst");
                var ids=new HashSet<uint>();long totalText=0,totalTagText=0,totalObjectText=0;int totalObjects=0;
                foreach(var reference in list==null?Enumerable.Empty<XElement>():list.Elements(p+"sldId"))
                {
                    cancellation.ThrowIfCancellationRequested();uint native;
                    if(!UInt32.TryParse((string)reference.Attribute("id"),NumberStyles.None,CultureInfo.InvariantCulture,out native)||native==0||!ids.Add(native))throw new InvalidDataException("Native slide identity is invalid or duplicated.");
                    if(result.Slides.Count>=10000)throw new InvalidDataException("Presentation slide limit exceeded.");
                    string part=Related(rels,"ppt/presentation.xml","slide",(string)reference.Attribute(relation+"id"),true);
                    var xml=Xml(zip,part,true);if(xml.Root==null||xml.Root.Name!=p+"sld")throw new InvalidDataException("Invalid slide root.");
                    string fingerprint=HashText(hash+"|"+part+"|"+xml.ToString(SaveOptions.DisableFormatting));SelectCompatibilityBranches(xml);
                    var slide=new IndexedSlide{NativeId=native,Position=result.Slides.Count+1,PartUri=part,Text=Text(xml,drawing),Layout="",Master="",Category="",ContentFingerprint=fingerprint};
                    if((totalText+=slide.Text.Length)>64L*1024*1024)throw new InvalidDataException("Presentation text limit exceeded.");
                    var titleShape=xml.Descendants(p+"sp").FirstOrDefault(s=>s.Descendants(p+"ph").Any(ph=>(string)ph.Attribute("type")=="title"||(string)ph.Attribute("type")=="ctrTitle"));
                    slide.Title=titleShape==null?slide.Text.Substring(0,Math.Min(100,slide.Text.Length)):Text(titleShape,drawing);
                    var slideRels=Xml(zip,RelationshipPart(part),false);
                    var tagRelationships=slideRels==null?Enumerable.Empty<XElement>():slideRels.Elements(PackageRels+"Relationships").Elements(PackageRels+"Relationship").Where(e=>((string)e.Attribute("Type")??"").EndsWith("/tags",StringComparison.Ordinal));
                    foreach(var tagReference in tagRelationships)
                    {
                        if((string)tagReference.Attribute("TargetMode")=="External")throw new InvalidDataException("External source tags are not fetched.");
                        string tagsPart=Resolve(part,(string)tagReference.Attribute("Target"));
                        var tags=Xml(zip,tagsPart,true);
                        if(tags.Root==null||tags.Root.Name!=p+"tagLst")throw new InvalidDataException("Source tag part has an invalid root.");
                        foreach(var tag in tags.Root.Elements(p+"tag"))
                        {
                            string name=(string)tag.Attribute("name"),value=(string)tag.Attribute("val");
                            if(String.IsNullOrWhiteSpace(name)||value==null||name.Length>256||value.Length>65536||slide.Tags.Count>=512)throw new InvalidDataException("Source tags exceed supported bounds.");
                            if((totalTagText+=name.Length+value.Length+1)>32L*1024*1024)throw new InvalidDataException("Presentation tag text limit exceeded.");
                            slide.Tags.Add(name+"="+value);
                        }
                    }
                    string layout=Related(slideRels,part,"slideLayout",null,false);
                    if(layout!=null){var layoutXml=Xml(zip,layout,true);slide.Layout=Name(layoutXml,p,layoutXml.Root==null?"":(string)layoutXml.Root.Attribute("type"));string master=Related(Xml(zip,RelationshipPart(layout),false),layout,"slideMaster",null,false);if(master!=null)slide.Master=Name(Xml(zip,master,true),p,master);}
                    // Native object IDs can be zero or repeated in real source packages. Preserve them and use a scoped structural key.
                    foreach(var shape in xml.Descendants().Where(e=>e.Name.Namespace==p&&SupportedShapes.Contains(e.Name.LocalName)))
                    {
                        cancellation.ThrowIfCancellationRequested();if(slide.Objects.Count>=20000||++totalObjects>500000)throw new InvalidDataException("Presentation object limit exceeded.");
                        string id=NativeObjectId(shape,p);
                        var nv=shape.Elements().First(e=>e.Name.LocalName.StartsWith("nv",StringComparison.Ordinal)).Descendants(p+"cNvPr").First();
                        var parent=shape.Ancestors(p+"grpSp").FirstOrDefault();
                        var item=new IndexedObject{Key=ObjectKey(shape),ParentKey=parent==null?null:ObjectKey(parent),NativeId=id,ParentId=parent==null?null:NativeObjectId(parent,p),Kind=shape.Name.LocalName,Name=(string)nv.Attribute("name")??"",Text=Text(shape,drawing),XmlFingerprint=HashText(shape.ToString(SaveOptions.DisableFormatting))};
                        var transform=shape.Element(p+"xfrm")??shape.Elements().Where(e=>e.Name.LocalName=="spPr"||e.Name.LocalName=="grpSpPr").Elements(drawing+"xfrm").FirstOrDefault();
                        if(transform!=null){item.HasTransform=true;item.X=Number(transform.Element(drawing+"off"),"x");item.Y=Number(transform.Element(drawing+"off"),"y");item.Width=Number(transform.Element(drawing+"ext"),"cx");item.Height=Number(transform.Element(drawing+"ext"),"cy");item.Rotation=Number(transform,"rot");item.FlipH=(string)transform.Attribute("flipH")=="1"||(string)transform.Attribute("flipH")=="true";item.FlipV=(string)transform.Attribute("flipV")=="1"||(string)transform.Attribute("flipV")=="true";}
                        if((totalObjectText+=item.Text.Length)>128L*1024*1024)throw new InvalidDataException("Presentation object text limit exceeded.");
                        slide.Objects.Add(item);
                    }
                    result.Slides.Add(slide);
                }
            }
            info.Refresh();if(info.Length!=result.Size||info.LastWriteTimeUtc!=result.ModifiedUtc)throw new IOException("Source metadata changed while reading; previous index must be retained.");
            if(!String.Equals(hash,Fingerprint(path,cancellation),StringComparison.Ordinal))throw new IOException("Source changed while reading; previous index must be retained.");
            return result;
        }
    }
}