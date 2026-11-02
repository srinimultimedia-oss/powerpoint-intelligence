using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
namespace PowerPointIntelligence.AddIn
{
    internal static class LocalLibraryCache
    {
        internal static string Encode(string folder,IEnumerable<LocalSlideAsset> items)
        {
            var doc=new XmlDocument {XmlResolver=null};var root=doc.CreateElement("Library","urn:powerpoint-intelligence:local-library:v1");doc.AppendChild(root);root.SetAttribute("version","1");root.SetAttribute("folder",Path.GetFullPath(folder));int count=0;
            foreach(var a in items){if(++count>10000)throw new InvalidDataException("Index exceeds 10000 slides.");var node=doc.CreateElement("Slide",root.NamespaceURI);root.AppendChild(node);foreach(var pair in new[]{Tuple.Create("file",a.File),Tuple.Create("title",a.Title),Tuple.Create("author",a.Author),Tuple.Create("layout",a.Layout),Tuple.Create("master",a.Master),Tuple.Create("signature",a.Signature),Tuple.Create("number",a.Slide.ToString(CultureInfo.InvariantCulture)),Tuple.Create("width",a.Width.ToString("R",CultureInfo.InvariantCulture)),Tuple.Create("height",a.Height.ToString("R",CultureInfo.InvariantCulture)),Tuple.Create("modified",a.Modified.ToUniversalTime().ToString("o",CultureInfo.InvariantCulture))})node.SetAttribute(pair.Item1,pair.Item2??"");node.InnerText=a.Text??"";}
            if(doc.OuterXml.Length>30000000)throw new InvalidDataException("Index exceeds 30 million characters.");return doc.OuterXml;
        }
        internal static List<LocalSlideAsset> Decode(string xml,out string folder)
        {
            var doc=new XmlDocument {XmlResolver=null};using(var reader=XmlReader.Create(new StringReader(xml),new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=30000000}))doc.Load(reader);var root=doc.DocumentElement;if(root==null||root.LocalName!="Library"||root.NamespaceURI!="urn:powerpoint-intelligence:local-library:v1"||root.GetAttribute("version")!="1")throw new InvalidDataException("Unknown library index version.");folder=Path.GetFullPath(root.GetAttribute("folder"));var result=new List<LocalSlideAsset>();long text=0;
            foreach(XmlElement node in root.ChildNodes.OfType<XmlElement>()){
                int number;float width,height;DateTime modified;string file=Path.GetFullPath(node.GetAttribute("file"));if(node.LocalName!="Slide"||node.NamespaceURI!=root.NamespaceURI||!LocalSlideIndex.IsWithin(folder,file)||!new[]{".pptx",".potx"}.Contains(Path.GetExtension(file).ToLowerInvariant())||!Int32.TryParse(node.GetAttribute("number"),out number)||number<1||number>1000||!Single.TryParse(node.GetAttribute("width"),NumberStyles.Float,CultureInfo.InvariantCulture,out width)||!Single.TryParse(node.GetAttribute("height"),NumberStyles.Float,CultureInfo.InvariantCulture,out height)||Single.IsNaN(width+height)||Single.IsInfinity(width+height)||width<=0||height<=0||!DateTime.TryParseExact(node.GetAttribute("modified"),"o",CultureInfo.InvariantCulture,DateTimeStyles.RoundtripKind,out modified)||String.IsNullOrWhiteSpace(node.GetAttribute("signature")))throw new InvalidDataException("Invalid cached slide metadata.");if(result.Count>=10000||node.InnerText.Length>64000||(text+=node.InnerText.Length)>20000000)throw new InvalidDataException("Cached slide text exceeds limits.");result.Add(new LocalSlideAsset {File=file,Slide=number,Width=width,Height=height,Modified=modified.ToUniversalTime(),Signature=node.GetAttribute("signature"),Title=node.GetAttribute("title"),Author=node.GetAttribute("author"),Layout=node.GetAttribute("layout"),Master=node.GetAttribute("master"),Text=node.InnerText});
            }return result;
        }
    }
}
