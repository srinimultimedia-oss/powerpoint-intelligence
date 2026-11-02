using PowerPointIntelligence.Security;
using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using System.Security.Cryptography;
using System.Text;
using Office=Microsoft.Office.Core;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
internal sealed class TemplatePixelResult
{
    internal bool SameDimensions,Exact,Accepted;
    internal int ChangedPixels,MaxChannelDelta,AlphaChanges,NonEdgeChanges,AllowedPixels;
    internal string Csv {get{return SameDimensions+","+Exact+","+Accepted+","+ChangedPixels+","+MaxChannelDelta+","+AlphaChanges+","+NonEdgeChanges+","+AllowedPixels;}}
}
internal static class TemplateFidelity
{
    // Owner-approved policy: never blur, resize, ignore regions, or change alpha.
    internal const int MaxPixels=8,MaxDelta=10;
    internal const double MaxFraction=.00005; // 0.005%, also capped at eight pixels.
    private static int Delta(Color a,Color b){return Math.Max(Math.Abs(a.R-b.R),Math.Max(Math.Abs(a.G-b.G),Math.Abs(a.B-b.B)));}
    private static bool Edge(Bitmap image,int x,int y)
    {var center=image.GetPixel(x,y);for(int dy=-1;dy<=1;dy++)for(int dx=-1;dx<=1;dx++){int nx=x+dx,ny=y+dy;if(nx>=0&&ny>=0&&nx<image.Width&&ny<image.Height&&Delta(center,image.GetPixel(nx,ny))>=16)return true;}return false;}
    internal static TemplatePixelResult Compare(Bitmap expected,Bitmap actual)
    {
        if(expected==null||actual==null)throw new ArgumentNullException("images");
        var r=new TemplatePixelResult {SameDimensions=expected.Width==actual.Width&&expected.Height==actual.Height};if(!r.SameDimensions)return r;
        long total=(long)expected.Width*expected.Height;if(total>2000000)throw new InvalidDataException("Fidelity comparison is bounded to two million pixels.");r.AllowedPixels=(int)Math.Min(MaxPixels,Math.Floor(total*MaxFraction));
        for(int y=0;y<expected.Height;y++)for(int x=0;x<expected.Width;x++){
            var a=expected.GetPixel(x,y);var b=actual.GetPixel(x,y);if(a.ToArgb()==b.ToArgb())continue;r.ChangedPixels++;r.MaxChannelDelta=Math.Max(r.MaxChannelDelta,Delta(a,b));if(a.A!=b.A)r.AlphaChanges++;if(!Edge(expected,x,y))r.NonEdgeChanges++;
        }
        r.Exact=r.ChangedPixels==0;r.Accepted=r.Exact||(r.ChangedPixels<=r.AllowedPixels&&r.MaxChannelDelta<=MaxDelta&&r.AlphaChanges==0&&r.NonEdgeChanges==0);return r;
    }
    internal static TemplatePixelResult Files(string expected,string actual)
    {using(var a=new Bitmap(expected))using(var b=new Bitmap(actual))return Compare(a,b);}
    private const string P="http://schemas.openxmlformats.org/presentationml/2006/main",A="http://schemas.openxmlformats.org/drawingml/2006/main",R="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static XmlDocument ReadPart(ZipArchive zip,string name)
    {
        var entries=zip.Entries.Where(e=>e.FullName==name).ToArray();if(entries.Length!=1||entries[0].Length>32000000)throw new InvalidDataException("Missing, ambiguous or oversized geometry package part: "+name);
        var doc=new XmlDocument {XmlResolver=null};using(var stream=entries[0].Open())using(var reader=XmlReader.Create(stream,new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=32000000}))doc.Load(reader);return doc;
    }
    private static void Canonical(XmlElement element,StringBuilder value)
    {
        value.Append('<').Append('{').Append(element.NamespaceURI).Append('}').Append(element.LocalName);
        foreach(var a in element.Attributes.Cast<XmlAttribute>().Where(a=>a.NamespaceURI!="http://www.w3.org/2000/xmlns/").OrderBy(a=>a.NamespaceURI,StringComparer.Ordinal).ThenBy(a=>a.LocalName,StringComparer.Ordinal))value.Append(' ').Append('{').Append(a.NamespaceURI).Append('}').Append(a.LocalName).Append('=').Append(a.Value.Length).Append(':').Append(a.Value);
        value.Append('>');foreach(XmlNode child in element.ChildNodes){if(child is XmlElement)Canonical((XmlElement)child,value);else if(child.NodeType==XmlNodeType.Text||child.NodeType==XmlNodeType.CDATA)value.Append(child.Value.Length).Append(':').Append(child.Value);else if(child.NodeType!=XmlNodeType.Whitespace&&child.NodeType!=XmlNodeType.SignificantWhitespace)throw new InvalidDataException("Unsupported geometry XML node.");}value.Append("</>");
    }
    internal static string GeometryXml(string xml)
    {
        if(xml==null||xml.Length>32000000)throw new InvalidDataException("Missing/oversized slide geometry XML.");var doc=new XmlDocument {XmlResolver=null};using(var reader=XmlReader.Create(new StringReader(xml),new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=32000000}))doc.Load(reader);
        var ns=new XmlNamespaceManager(doc.NameTable);ns.AddNamespace("p",P);var tree=doc.SelectSingleNode("/p:sld/p:cSld/p:spTree",ns) as XmlElement;if(tree==null)throw new InvalidDataException("Slide geometry tree missing.");var result=new StringBuilder("drawingml-geometry-v2|");int count=0;
        foreach(XmlElement props in tree.ChildNodes.OfType<XmlElement>().Where(e=>e.NamespaceURI==P&&e.LocalName=="grpSpPr"))foreach(XmlElement transform in props.ChildNodes.OfType<XmlElement>().Where(e=>e.NamespaceURI==A&&e.LocalName=="xfrm"))Canonical(transform,result);
        Action<XmlElement,int> visit=null;visit=(parent,depth)=>{if(depth>8)throw new InvalidDataException("Geometry nesting exceeds eight levels.");foreach(XmlElement shape in parent.ChildNodes.OfType<XmlElement>()){
            if(shape.NamespaceURI==P&&(shape.LocalName=="nvGrpSpPr"||shape.LocalName=="grpSpPr"||shape.LocalName=="extLst"))continue;
            if(shape.NamespaceURI!=P||!new[]{"sp","grpSp","cxnSp","pic","graphicFrame"}.Contains(shape.LocalName))throw new InvalidDataException("Unsupported shape container in persisted geometry.");if(++count>2000)throw new InvalidDataException("Geometry exceeds 2000 shapes.");result.Append('[').Append(shape.LocalName).Append('|');
            foreach(XmlElement props in shape.ChildNodes.OfType<XmlElement>()){
                if(props.NamespaceURI==P&&(props.LocalName=="spPr"||props.LocalName=="grpSpPr"))foreach(XmlElement item in props.ChildNodes.OfType<XmlElement>())if(item.NamespaceURI==A&&new[]{"xfrm","prstGeom","custGeom"}.Contains(item.LocalName))Canonical(item,result);
                if(props.NamespaceURI==P&&props.LocalName=="xfrm")Canonical(props,result);
            }
            if(shape.LocalName=="grpSp")visit(shape,depth+1);result.Append(']');
        }};visit(tree,0);return result.ToString();
    }
    private static string RelatedPart(ZipArchive zip,string owner,string kind,string id=null)
    {
        int slash=owner.LastIndexOf('/');string folder=owner.Substring(0,slash),name=owner.Substring(slash+1);var rels=ReadPart(zip,folder+"/_rels/"+name+".rels");const string ns="http://schemas.openxmlformats.org/package/2006/relationships";
        if(rels.DocumentElement.LocalName!="Relationships"||rels.DocumentElement.NamespaceURI!=ns)throw new InvalidDataException("Unsupported relationship namespace.");
        var matches=rels.DocumentElement.ChildNodes.OfType<XmlElement>().Where(e=>e.NamespaceURI==ns&&e.LocalName=="Relationship"&&(id==null||e.GetAttribute("Id")==id)&&e.GetAttribute("Type").EndsWith("/"+kind,StringComparison.Ordinal)).ToArray();
        if(matches.Length!=1||matches[0].GetAttribute("TargetMode")=="External")throw new InvalidDataException("Missing, ambiguous or external "+kind+" relationship.");
        string target=matches[0].GetAttribute("Target");if(String.IsNullOrEmpty(target)||target.Contains(":")||target.Contains("\\")||target.Contains("?")||target.Contains("#"))throw new InvalidDataException("Unsafe package target.");
        var path=new System.Collections.Generic.List<string>();foreach(var item in (target.StartsWith("/",StringComparison.Ordinal)?target:folder+"/"+target).Split('/')){if(item==""||item==".")continue;if(item==".."){if(path.Count==0)throw new InvalidDataException("Escaping package target.");path.RemoveAt(path.Count-1);}else path.Add(item);}string part=String.Join("/",path),expected=kind=="slide"?"ppt/slides/":kind=="slideLayout"?"ppt/slideLayouts/":"ppt/slideMasters/";if(!part.StartsWith(expected,StringComparison.Ordinal))throw new InvalidDataException("Unexpected package target for "+kind+".");return part;
    }
    private static XmlDocument ReadXml(string xml)
    {if(xml==null||xml.Length>32000000)throw new InvalidDataException("Missing/oversized inheritance XML.");var doc=new XmlDocument {XmlResolver=null};using(var reader=XmlReader.Create(new StringReader(xml),new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=32000000}))doc.Load(reader);return doc;}
    private static XmlNamespaceManager Names(XmlDocument doc){var ns=new XmlNamespaceManager(doc.NameTable);ns.AddNamespace("p",P);ns.AddNamespace("a",A);return ns;}
    private static XmlElement Transform(XmlElement shape){return shape.SelectSingleNode("p:spPr/a:xfrm",Names(shape.OwnerDocument)) as XmlElement;}
    private static XmlElement Placeholder(XmlElement shape){return shape.SelectSingleNode("p:nvSpPr/p:nvPr/p:ph",Names(shape.OwnerDocument)) as XmlElement;}
    private static string Index(XmlElement ph){string value=ph.GetAttribute("idx");uint index;if(value=="")return "0";if(!UInt32.TryParse(value,System.Globalization.NumberStyles.None,CultureInfo.InvariantCulture,out index))throw new InvalidDataException("Invalid placeholder index.");return index.ToString(CultureInfo.InvariantCulture);}
    private static string Role(XmlElement ph){string role=ph.GetAttribute("type");return role==""?"obj":role=="ctrTitle"?"title":role;}
    private static XmlElement Match(XmlDocument doc,XmlElement ph,bool byType)
    {
        var candidates=doc.SelectNodes("/p:"+(byType?"sldMaster":"sldLayout")+"/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph]",Names(doc)).Cast<XmlElement>().Where(s=>byType?Role(Placeholder(s))==Role(ph):Index(Placeholder(s))==Index(ph)).ToArray();
        if(candidates.Length!=1)throw new InvalidDataException("Placeholder inheritance is missing or ambiguous ("+(byType?"type":"index")+").");return candidates[0];
    }
    private static bool NeedsInheritance(XmlDocument slide){return slide.SelectNodes("/p:sld/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph]",Names(slide)).Cast<XmlElement>().Any(s=>Transform(s)==null);}
    private static void Inherit(XmlDocument slide,Func<XmlDocument> layout,Func<XmlDocument> master)
    {
        foreach(var shape in slide.SelectNodes("/p:sld/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph]",Names(slide)).Cast<XmlElement>()){
            if(Transform(shape)!=null)continue;var inherited=Match(layout(),Placeholder(shape),false);var transform=Transform(inherited);if(transform==null)transform=Transform(Match(master(),Placeholder(inherited),true));if(transform==null)throw new InvalidDataException("Placeholder has no explicit transform anywhere in its inheritance chain.");
            var props=shape.SelectSingleNode("p:spPr",Names(slide)) as XmlElement;if(props==null){props=slide.CreateElement("p","spPr",P);shape.AppendChild(props);}props.PrependChild(slide.ImportNode(transform,true));
        }
    }
    internal static string EffectiveGeometryXml(string slideXml,string layoutXml,string masterXml)
    {var slide=ReadXml(slideXml);XmlDocument layout=null,master=null;Inherit(slide,()=>layout??(layout=ReadXml(layoutXml)),()=>master??(master=ReadXml(masterXml)));return GeometryXml(slide.OuterXml);}
    internal static string PackageGeometry(string file,int slideIndex)
    {
        if(slideIndex<1)throw new InvalidDataException("Invalid slide index.");using(var zip=ZipFile.OpenRead(file)){
            var presentation=ReadPart(zip,"ppt/presentation.xml");var ns=Names(presentation);var ids=presentation.SelectNodes("/p:presentation/p:sldIdLst/p:sldId",ns);if(slideIndex>ids.Count)throw new InvalidDataException("Slide index outside package.");string id=((XmlElement)ids[slideIndex-1]).GetAttribute("id",R);string part=RelatedPart(zip,"ppt/presentation.xml","slide",id);var slide=ReadPart(zip,part);
            if(NeedsInheritance(slide)){XmlDocument layout=null,master=null;string layoutPart=null;Func<XmlDocument> getLayout=()=>layout??(layout=ReadPart(zip,layoutPart=RelatedPart(zip,part,"slideLayout")));Func<XmlDocument> getMaster=()=>{getLayout();return master??(master=ReadPart(zip,RelatedPart(zip,layoutPart,"slideMaster")));};Inherit(slide,getLayout,getMaster);}
            return GeometryXml(slide.OuterXml);
        }
    }
    private static string PersistedGeometry(PowerPoint.Slide slide,string diagnosticFile)
    {
        string file=Path.Combine(Path.GetTempPath(),"ppi-fidelity-geometry-"+Guid.NewGuid().ToString("N")+".pptx");
        try{var deck=(PowerPoint.Presentation)slide.Parent;deck.SaveCopyAs(file,PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,Office.MsoTriState.msoFalse);string geometry=PackageGeometry(file,slide.SlideIndex);if(diagnosticFile!=null)File.WriteAllText(PathPolicy.TestPath(diagnosticFile),geometry);return geometry;}
        finally{if(File.Exists(file))File.Delete(file);}
    }
    // Native values are compared exactly. Rendering tolerance never applies here.
    internal static string Formatting(PowerPoint.Slide slide,string diagnosticFile=null)
    {
        var text=new StringBuilder("native-template-format-v3|");int visited=0,characters=0;
        Action<object> append=value=>text.Append(Convert.ToString(value,CultureInfo.InvariantCulture)).Append('|');
        append(PersistedGeometry(slide,diagnosticFile==null?null:diagnosticFile+".geometry.txt"));
        Action<PowerPoint.Shape,int> visit=null;visit=(shape,depth)=>{
            if(++visited>2000||depth>8)throw new InvalidDataException("Native formatting fingerprint exceeds its shape bound.");append(shape.Type);append(shape.Visible);append(shape.HorizontalFlip);append(shape.VerticalFlip);
            if(shape.Type==Office.MsoShapeType.msoAutoShape){append(shape.AutoShapeType);append(shape.Adjustments.Count);for(int i=1;i<=shape.Adjustments.Count;i++)append(shape.Adjustments[i]);}
            append(shape.Fill.Visible);append(shape.Fill.Type);
            if(shape.Fill.Visible==Office.MsoTriState.msoTrue){append(shape.Fill.ForeColor.RGB);append(shape.Fill.BackColor.RGB);append(shape.Fill.Transparency);if(shape.Fill.Type==Office.MsoFillType.msoFillGradient){append(shape.Fill.GradientStyle);append(shape.Fill.GradientVariant);append(shape.Fill.GradientDegree);append(shape.Fill.GradientStops.Count);foreach(Office.GradientStop stop in shape.Fill.GradientStops){append(stop.Position);append(stop.Color.RGB);append(stop.Transparency);}}}
            append(shape.Line.Visible);if(shape.Line.Visible==Office.MsoTriState.msoTrue){append(shape.Line.ForeColor.RGB);append(shape.Line.Transparency);append(shape.Line.Weight);append(shape.Line.DashStyle);append(shape.Line.Style);append(shape.Line.BeginArrowheadStyle);append(shape.Line.EndArrowheadStyle);}
            if(shape.HasTextFrame==Office.MsoTriState.msoTrue){var frame=shape.TextFrame;append(frame.MarginLeft);append(frame.MarginRight);append(frame.MarginTop);append(frame.MarginBottom);append(frame.VerticalAnchor);append(frame.Orientation);append(frame.WordWrap);append(frame.AutoSize);var range=frame.TextRange;
                if((characters+=range.Length)>50000)throw new InvalidDataException("Native text fingerprint exceeds 50,000 characters.");
                for(int i=1;i<=range.Length;i++){var c=range.Characters(i,1);var f=c.Font;append(f.Name);append(f.Size);append(f.Bold);append(f.Italic);append(f.Underline);append(f.Shadow);append(f.Emboss);append(f.BaselineOffset);append(f.Color.RGB);var p=c.ParagraphFormat;append(p.Alignment);append(p.SpaceBefore);append(p.SpaceAfter);append(p.SpaceWithin);append(p.Bullet.Visible);}
            }
            if(shape.Type==Office.MsoShapeType.msoPicture||shape.Type==Office.MsoShapeType.msoLinkedPicture){var p=shape.PictureFormat;append(p.CropLeft);append(p.CropTop);append(p.CropRight);append(p.CropBottom);append(p.Brightness);append(p.Contrast);append(p.ColorType);}
            if(shape.Type==Office.MsoShapeType.msoGroup)foreach(PowerPoint.Shape child in shape.GroupItems)visit(child,depth+1);
        };foreach(PowerPoint.Shape shape in slide.Shapes)visit(shape,0);
        if(diagnosticFile!=null)File.WriteAllText(PathPolicy.TestPath(diagnosticFile),text.ToString());
        using(var sha=SHA256.Create())return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(text.ToString()))).Replace("-","");
    }
}
