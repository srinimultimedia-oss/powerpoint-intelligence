using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Windows.Forms;
using Office=Microsoft.Office.Core;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
namespace PowerPointIntelligence.AddIn
{
    internal sealed class ConnectorEndpoint
    {
        internal string Kind="site",Anchor="center";
        internal int ShapeId,Site,Series=1,Point=1;
        internal float X,Y;internal bool Rendered;
        internal ConnectorEndpoint Copy(){return (ConnectorEndpoint)MemberwiseClone();}
        internal void Validate(){if(!new[]{"site","point","free"}.Contains(Kind)||!new[]{"center","top","right","bottom","left"}.Contains(Anchor)||Kind!="free"&&ShapeId<1||Site<0||Site>1024||Series<1||Series>100||Point<1||Point>10000||ShapeId<0||!ConnectorSpec.Finite(X)||!ConnectorSpec.Finite(Y)||X<0||Y<0||X>10000||Y>10000||Kind=="free"&&(ShapeId!=0||Site!=0))throw new InvalidDataException("Invalid connector endpoint.");}
    }
    internal sealed class ConnectorSpec
    {
        internal ConnectorEndpoint Begin=new ConnectorEndpoint(),End=new ConnectorEndpoint();
        internal string Route="elbow",Dash="solid",BeginArrow="none",EndArrow="triangle";
        internal int Color=0x7E6552;
        internal float Weight=1.5f;
        internal bool AutoRoute=true;
        internal static bool Finite(float x){return !Single.IsNaN(x)&&!Single.IsInfinity(x);}
        internal void Validate(){if(Begin==null||End==null)throw new InvalidDataException("Two endpoints required.");Begin.Validate();End.Validate();if(!new[]{"straight","elbow","curve"}.Contains(Route)||!new[]{"solid","dash","dot","dashdot"}.Contains(Dash)||!new[]{"none","triangle","open","stealth"}.Contains(BeginArrow)||!new[]{"none","triangle","open","stealth"}.Contains(EndArrow)||Color<0||Color>0xFFFFFF||!Finite(Weight)||Weight<.25||Weight>12)throw new InvalidDataException("Invalid connector styling.");if(Begin.Kind=="free"&&End.Kind=="free"&&Math.Abs(Begin.X-End.X)<.01&&Math.Abs(Begin.Y-End.Y)<.01)throw new InvalidDataException("Choose different free endpoint coordinates.");if(Begin.Kind==End.Kind&&Begin.ShapeId==End.ShapeId&&(Begin.Kind=="site"&&Begin.Site==End.Site||Begin.Kind=="point"&&Begin.Series==End.Series&&Begin.Point==End.Point&&Begin.Anchor==End.Anchor))throw new InvalidDataException("Choose different connection points.");}
        internal ConnectorSpec Copy(){var s=(ConnectorSpec)MemberwiseClone();s.Begin=Begin.Copy();s.End=End.Copy();return s;}
        internal static PointF AnchorAt(RectangleF bounds,string anchor)
        {if(!Finite(bounds.Left)||!Finite(bounds.Top)||!Finite(bounds.Width)||!Finite(bounds.Height)||bounds.Width<0||bounds.Height<0||!Finite(bounds.Right)||!Finite(bounds.Bottom))throw new InvalidDataException("Native point bounds unavailable.");switch(anchor){case "center":return new PointF(bounds.Left+bounds.Width/2,bounds.Top+bounds.Height/2);case "top":return new PointF(bounds.Left+bounds.Width/2,bounds.Top);case "right":return new PointF(bounds.Right,bounds.Top+bounds.Height/2);case "bottom":return new PointF(bounds.Left+bounds.Width/2,bounds.Bottom);case "left":return new PointF(bounds.Left,bounds.Top+bounds.Height/2);default:throw new InvalidDataException("Unknown endpoint anchor.");}}
    }
    internal static class ConnectorCodec
    {
        internal const string Namespace="urn:powerpoint-intelligence:connector:v1";
        private static readonly XNamespace Ns=Namespace;
        internal static string Encode(ConnectorSpec s)
        {if(s==null)throw new InvalidDataException("Missing connector settings.");s.Validate();Func<string,ConnectorEndpoint,XElement> endpoint=(role,e)=>new XElement(Ns+role,new XAttribute("kind",e.Kind),new XAttribute("shape",e.ShapeId),new XAttribute("site",e.Site),new XAttribute("series",e.Series),new XAttribute("point",e.Point),new XAttribute("anchor",e.Anchor),new XAttribute("x",e.X.ToString("R",CultureInfo.InvariantCulture)),new XAttribute("y",e.Y.ToString("R",CultureInfo.InvariantCulture)),new XAttribute("rendered",e.Rendered?"1":"0"));return new XElement(Ns+"connector",new XAttribute("version","2"),new XAttribute("route",s.Route),new XAttribute("dash",s.Dash),new XAttribute("beginArrow",s.BeginArrow),new XAttribute("endArrow",s.EndArrow),new XAttribute("color",s.Color),new XAttribute("weight",s.Weight.ToString("R",CultureInfo.InvariantCulture)),new XAttribute("auto",s.AutoRoute?"1":"0"),endpoint("begin",s.Begin),endpoint("end",s.End)).ToString(SaveOptions.DisableFormatting);}
        private static void Shape(XElement node,string name,params string[] attrs){if(node==null||node.Name!=Ns+name||node.Attributes().Any(a=>!a.IsNamespaceDeclaration&&(a.Name.NamespaceName!=""||!attrs.Contains(a.Name.LocalName)))||attrs.Any(a=>node.Attribute(a)==null)||node.Nodes().Any(n=>!(n is XElement)&&!(n is XText&&String.IsNullOrWhiteSpace(((XText)n).Value))))throw new InvalidDataException("Unknown connector metadata fields.");}
        internal static ConnectorSpec Decode(string xml)
        {
            if(String.IsNullOrWhiteSpace(xml)||xml.Length>16384)throw new InvalidDataException("Missing/excessive connector metadata.");try{XDocument doc;using(var reader=XmlReader.Create(new StringReader(xml),new XmlReaderSettings{DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=16384}))doc=XDocument.Load(reader);if(doc.Nodes().Any(n=>!(n is XElement)))throw new InvalidDataException("Unexpected connector document nodes.");var root=doc.Root;Shape(root,"connector","version","route","dash","beginArrow","endArrow","color","weight","auto");string version=(string)root.Attribute("version");if(!new[]{"1","2"}.Contains(version)||root.Elements().Count()!=2||!new[]{"0","1"}.Contains((string)root.Attribute("auto")))throw new InvalidDataException("Unsupported connector version.");Func<string,ConnectorEndpoint> endpoint=role=>{var e=root.Element(Ns+role);Shape(e,role,version=="1"?new[]{"kind","shape","site","series","point","anchor","x","y"}:new[]{"kind","shape","site","series","point","anchor","x","y","rendered"});if(e.HasElements||version=="2"&&!new[]{"0","1"}.Contains((string)e.Attribute("rendered")))throw new InvalidDataException("Nested endpoint fields.");return new ConnectorEndpoint {Kind=(string)e.Attribute("kind"),ShapeId=Int32.Parse((string)e.Attribute("shape"),NumberStyles.None,CultureInfo.InvariantCulture),Site=Int32.Parse((string)e.Attribute("site"),NumberStyles.None,CultureInfo.InvariantCulture),Series=Int32.Parse((string)e.Attribute("series"),NumberStyles.None,CultureInfo.InvariantCulture),Point=Int32.Parse((string)e.Attribute("point"),NumberStyles.None,CultureInfo.InvariantCulture),Anchor=(string)e.Attribute("anchor"),X=Single.Parse((string)e.Attribute("x"),CultureInfo.InvariantCulture),Y=Single.Parse((string)e.Attribute("y"),CultureInfo.InvariantCulture),Rendered=version=="2"&&(string)e.Attribute("rendered")=="1"};};var s=new ConnectorSpec {Begin=endpoint("begin"),End=endpoint("end"),Route=(string)root.Attribute("route"),Dash=(string)root.Attribute("dash"),BeginArrow=(string)root.Attribute("beginArrow"),EndArrow=(string)root.Attribute("endArrow"),Color=Int32.Parse((string)root.Attribute("color"),CultureInfo.InvariantCulture),Weight=Single.Parse((string)root.Attribute("weight"),CultureInfo.InvariantCulture),AutoRoute=(string)root.Attribute("auto")=="1"};s.Validate();return s;}catch(Exception error) when(error is XmlException||error is FormatException||error is OverflowException||error is ArgumentException){throw new InvalidDataException("Malformed connector metadata.",error);}
        }
    }
    internal static class ConnectorBindingPlan
    {
        internal static ConnectorEndpoint Reconcile(ConnectorEndpoint prior,bool connected,int nativeShapeId,int nativeSite,PointF position)
        {
            prior.Validate();if(!ConnectorSpec.Finite(position.X)||!ConnectorSpec.Finite(position.Y)||position.X<0||position.Y<0||position.X>10000||position.Y>10000)throw new InvalidDataException("Native connector endpoint position is invalid.");
            var result=prior.Copy();
            if(connected){if(nativeShapeId<1||nativeSite<1||nativeSite>1024)throw new InvalidDataException("Native connector attachment is invalid.");result.Kind="site";result.ShapeId=nativeShapeId;result.Site=nativeSite;}
            else if(prior.Kind=="site"||prior.Kind=="point"&&(!prior.Rendered||Math.Abs(prior.X-position.X)>.75f||Math.Abs(prior.Y-position.Y)>.75f)){result.Kind="free";result.ShapeId=0;result.Site=0;}
            bool unchangedBinding=prior.Kind==result.Kind&&prior.ShapeId==result.ShapeId&&prior.Site==result.Site;
            if(!prior.Rendered||!unchangedBinding||Math.Abs(prior.X-position.X)>.1f||Math.Abs(prior.Y-position.Y)>.1f){result.X=position.X;result.Y=position.Y;}
            result.Rendered=true;result.Validate();return result;
        }
    }
    internal static class ConnectorGeometryPlan
    {
        internal static PointF Position(RectangleF bounds,float rotation,bool flipH,bool flipV,bool begin)
        {
            if(new[]{bounds.X,bounds.Y,bounds.Width,bounds.Height,rotation}.Any(x=>!ConnectorSpec.Finite(x)||Math.Abs(x)>1000000)||bounds.Width<0||bounds.Height<0)throw new InvalidDataException("Invalid native connector transform.");
            double angle=(rotation%360)*Math.PI/180,cos=Math.Cos(angle),sin=Math.Sin(angle);if(Math.Abs(cos)<1e-12)cos=0;if(Math.Abs(sin)<1e-12)sin=0;
            double dx=(flipH?1:-1)*bounds.Width/2.0,dy=(flipV?1:-1)*bounds.Height/2.0;if(!begin){dx=-dx;dy=-dy;}
            return new PointF((float)(bounds.X+bounds.Width/2.0+dx*cos-dy*sin),(float)(bounds.Y+bounds.Height/2.0+dx*sin+dy*cos));
        }
    }
    internal static class UniversalConnectors
    {
        private static PowerPoint.Slide Active(PowerPoint.Application app){if(app==null||app.Presentations.Count==0||app.ActiveWindow.ViewType!=PowerPoint.PpViewType.ppViewNormal)throw new InvalidOperationException("Select slide objects in Normal view.");var slide=app.ActiveWindow.View.Slide as PowerPoint.Slide;if(slide==null)throw new InvalidOperationException("Select a slide in Normal view.");return slide;}
        internal static bool IsManaged(PowerPoint.Shape shape){return shape!=null&&shape.Tags["PPI_CONNECTOR"]=="1";}
        internal static ConnectorSpec Read(PowerPoint.Shape shape){var parts=XmlPartUpdates.AttachedParts(shape,ConnectorCodec.Namespace);if(parts.Length!=1)throw new InvalidDataException("Connector metadata missing/ambiguous; canonicalMatches="+parts.Length+".");return ConnectorCodec.Decode(parts[0].XML);}
        private static PowerPoint.Shape Find(PowerPoint.Slide slide,int id){var found=slide.Shapes.Cast<PowerPoint.Shape>().Where(s=>s.Id==id).ToArray();if(found.Length!=1)throw new InvalidDataException("Connector target is missing or ambiguous.");return found[0];}
        private static PointF Locate(PowerPoint.Slide slide,ConnectorEndpoint e)
        {
            if(e.Kind=="free")return new PointF(e.X,e.Y);var shape=Find(slide,e.ShapeId);if(e.Kind=="site"){if(shape.ConnectionSiteCount<1||e.Site>shape.ConnectionSiteCount)throw new InvalidDataException("Selected native connection site is unavailable.");return ConnectorSpec.AnchorAt(new RectangleF(shape.Left,shape.Top,shape.Width,shape.Height),"center");}
            if(shape.HasChart!=Office.MsoTriState.msoTrue||Math.Abs(shape.Rotation)>.01||shape.HorizontalFlip==Office.MsoTriState.msoTrue||shape.VerticalFlip==Office.MsoTriState.msoTrue)throw new InvalidOperationException("Chart point anchors need an unrotated, unflipped native chart.");var chart=shape.Chart;var series=(PowerPoint.SeriesCollection)chart.SeriesCollection();if(e.Series>series.Count)throw new InvalidDataException("Series no longer exists.");var points=(PowerPoint.Points)series.Item(e.Series).Points();if(e.Point>points.Count)throw new InvalidDataException("Point no longer exists.");dynamic point=points.Item(e.Point);float cw=(float)chart.ChartArea.Width,ch=(float)chart.ChartArea.Height;if(!ConnectorSpec.Finite(cw)||!ConnectorSpec.Finite(ch)||cw<=0||ch<=0)throw new InvalidDataException("Native chart dimensions are unavailable.");float sx=shape.Width/cw,sy=shape.Height/ch;var bounds=new RectangleF(shape.Left+(float)point.Left*sx,shape.Top+(float)point.Top*sy,(float)point.Width*sx,(float)point.Height*sy);return ConnectorSpec.AnchorAt(bounds,e.Anchor);
        }
        private static Office.MsoConnectorType Route(string route){return route=="straight"?Office.MsoConnectorType.msoConnectorStraight:route=="curve"?Office.MsoConnectorType.msoConnectorCurve:Office.MsoConnectorType.msoConnectorElbow;}
        private static Office.MsoArrowheadStyle Arrow(string value){return value=="none"?Office.MsoArrowheadStyle.msoArrowheadNone:value=="open"?Office.MsoArrowheadStyle.msoArrowheadOpen:value=="stealth"?Office.MsoArrowheadStyle.msoArrowheadStealth:Office.MsoArrowheadStyle.msoArrowheadTriangle;}
        private static void Style(PowerPoint.Shape line,ConnectorSpec s){line.ConnectorFormat.Type=Route(s.Route);line.Line.ForeColor.RGB=s.Color;line.Line.Weight=s.Weight;line.Line.DashStyle=s.Dash=="dash"?Office.MsoLineDashStyle.msoLineDash:s.Dash=="dot"?Office.MsoLineDashStyle.msoLineRoundDot:s.Dash=="dashdot"?Office.MsoLineDashStyle.msoLineDashDot:Office.MsoLineDashStyle.msoLineSolid;line.Line.BeginArrowheadStyle=Arrow(s.BeginArrow);line.Line.EndArrowheadStyle=Arrow(s.EndArrow);}
        private static void Geometry(PowerPoint.Shape line,PowerPoint.Slide slide,ConnectorSpec s,PointF a,PointF b)
        {
            var c=line.ConnectorFormat;if(c.BeginConnected==Office.MsoTriState.msoTrue)c.BeginDisconnect();if(c.EndConnected==Office.MsoTriState.msoTrue)c.EndDisconnect();line.Rotation=0;if(line.HorizontalFlip==Office.MsoTriState.msoTrue)line.Flip(Office.MsoFlipCmd.msoFlipHorizontal);if(line.VerticalFlip==Office.MsoTriState.msoTrue)line.Flip(Office.MsoFlipCmd.msoFlipVertical);line.Left=Math.Min(a.X,b.X);line.Top=Math.Min(a.Y,b.Y);line.Width=Math.Abs(a.X-b.X);line.Height=Math.Abs(a.Y-b.Y);if(a.X>b.X)line.Flip(Office.MsoFlipCmd.msoFlipHorizontal);if(a.Y>b.Y)line.Flip(Office.MsoFlipCmd.msoFlipVertical);
            if(s.Begin.Kind=="site")c.BeginConnect(Find(slide,s.Begin.ShapeId),Math.Max(1,s.Begin.Site));if(s.End.Kind=="site")c.EndConnect(Find(slide,s.End.ShapeId),Math.Max(1,s.End.Site));if(s.AutoRoute&&s.Begin.Kind=="site"&&s.End.Kind=="site"&&s.Begin.Site==0&&s.End.Site==0)line.RerouteConnections();
        }
        private static void Bounds(PowerPoint.Presentation deck,PointF point){if(!ConnectorSpec.Finite(point.X)||!ConnectorSpec.Finite(point.Y)||point.X<0||point.Y<0||point.X>deck.PageSetup.SlideWidth||point.Y>deck.PageSetup.SlideHeight)throw new InvalidDataException("Connector anchor is outside the slide or unavailable.");}
        private static PointF Position(PowerPoint.Shape line,bool begin)
        {
            if(line.Connector!=Office.MsoTriState.msoTrue)throw new InvalidOperationException("Select a native connector.");
            return ConnectorGeometryPlan.Position(new RectangleF(line.Left,line.Top,line.Width,line.Height),line.Rotation,line.HorizontalFlip==Office.MsoTriState.msoTrue,line.VerticalFlip==Office.MsoTriState.msoTrue,begin);
        }
        private static ConnectorSpec Live(PowerPoint.Shape line,ConnectorSpec prior)
        {
            var c=line.ConnectorFormat;var result=prior.Copy();bool begin=c.BeginConnected==Office.MsoTriState.msoTrue,end=c.EndConnected==Office.MsoTriState.msoTrue;
            result.Begin=ConnectorBindingPlan.Reconcile(prior.Begin,begin,begin?c.BeginConnectedShape.Id:0,begin?c.BeginConnectionSite:0,Position(line,true));
            result.End=ConnectorBindingPlan.Reconcile(prior.End,end,end?c.EndConnectedShape.Id:0,end?c.EndConnectionSite:0,Position(line,false));result.Validate();return result;
        }
        private static void Snapshot(PowerPoint.Shape line,ConnectorSpec spec)
        {
            var a=Position(line,true);var b=Position(line,false);spec.Begin.X=a.X;spec.Begin.Y=a.Y;spec.End.X=b.X;spec.End.Y=b.Y;spec.Begin.Rendered=spec.End.Rendered=true;
        }
        private static Office.CustomXMLPart Part(PowerPoint.Shape line){Read(line);return XmlPartUpdates.AttachedParts(line,ConnectorCodec.Namespace).Single();}
        private sealed class Saved
        {
            internal PowerPoint.Shape Line;internal ConnectorSpec LiveSpec;internal PointF A,B;internal Office.CustomXMLPart Part;internal string Xml;
            internal RectangleF NativeBounds;internal float Rotation;internal bool FlipH,FlipV;
            internal Office.MsoConnectorType Type;internal Office.MsoLineDashStyle Dash;internal Office.MsoArrowheadStyle BeginArrow,EndArrow;internal int Rgb;internal Office.MsoThemeColorIndex Theme;internal float Weight,Tint;internal float[] Adjustments;
            internal static Saved Capture(PowerPoint.Shape line){if(line.Adjustments.Count>8)throw new InvalidOperationException("Connector adjustment geometry exceeds its safe restoration bound.");var part=UniversalConnectors.Part(line);var spec=Live(line,Read(line));return new Saved {Line=line,Part=part,Xml=part.XML,NativeBounds=new RectangleF(line.Left,line.Top,line.Width,line.Height),Rotation=line.Rotation,FlipH=line.HorizontalFlip==Office.MsoTriState.msoTrue,FlipV=line.VerticalFlip==Office.MsoTriState.msoTrue,LiveSpec=spec,A=Position(line,true),B=Position(line,false),Type=line.ConnectorFormat.Type,Dash=line.Line.DashStyle,BeginArrow=line.Line.BeginArrowheadStyle,EndArrow=line.Line.EndArrowheadStyle,Rgb=line.Line.ForeColor.RGB,Theme=line.Line.ForeColor.ObjectThemeColor,Tint=line.Line.ForeColor.TintAndShade,Weight=line.Line.Weight,Adjustments=Enumerable.Range(1,line.Adjustments.Count).Select(i=>line.Adjustments[i]).ToArray()};}
            internal void Restore(PowerPoint.Slide slide){Line.ConnectorFormat.Type=Type;Geometry(Line,slide,LiveSpec,A,B);Line.Rotation=0;if(Line.HorizontalFlip==Office.MsoTriState.msoTrue)Line.Flip(Office.MsoFlipCmd.msoFlipHorizontal);if(Line.VerticalFlip==Office.MsoTriState.msoTrue)Line.Flip(Office.MsoFlipCmd.msoFlipVertical);Line.Width=NativeBounds.Width;Line.Height=NativeBounds.Height;Line.Left=NativeBounds.Left;Line.Top=NativeBounds.Top;if(FlipH)Line.Flip(Office.MsoFlipCmd.msoFlipHorizontal);if(FlipV)Line.Flip(Office.MsoFlipCmd.msoFlipVertical);Line.Rotation=Rotation;if(Math.Abs(Line.Rotation-Rotation)>.01||Math.Abs(Line.Left-NativeBounds.Left)>.1||Math.Abs(Line.Top-NativeBounds.Top)>.1||Math.Abs(Line.Width-NativeBounds.Width)>.1||Math.Abs(Line.Height-NativeBounds.Height)>.1||(Line.HorizontalFlip==Office.MsoTriState.msoTrue)!=FlipH||(Line.VerticalFlip==Office.MsoTriState.msoTrue)!=FlipV)throw new InvalidOperationException("Native connector transform restoration could not be verified. Use Undo.");if(Line.Adjustments.Count!=Adjustments.Length)throw new InvalidOperationException("Native connector adjustments could not be restored.");for(int i=0;i<Adjustments.Length;i++)Line.Adjustments[i+1]=Adjustments[i];Line.Line.DashStyle=Dash;Line.Line.BeginArrowheadStyle=BeginArrow;Line.Line.EndArrowheadStyle=EndArrow;Line.Line.Weight=Weight;if(Theme==Office.MsoThemeColorIndex.msoNotThemeColor)Line.Line.ForeColor.RGB=Rgb;else Line.Line.ForeColor.ObjectThemeColor=Theme;Line.Line.ForeColor.TintAndShade=Tint;Part=XmlPartUpdates.Shape(Line,UniversalConnectors.Part(Line),Xml);}
        }
        internal static PowerPoint.Shape Create(PowerPoint.Application app,ConnectorSpec spec)
        {
            spec.Validate();spec=spec.Copy();var slide=Active(app);var deck=app.ActivePresentation;var a=Locate(slide,spec.Begin);var b=Locate(slide,spec.End);Bounds(deck,a);Bounds(deck,b);app.StartNewUndoEntry();PowerPoint.Shape line=null;try{line=slide.Shapes.AddConnector(Route(spec.Route),a.X,a.Y,b.X,b.Y);Style(line,spec);Geometry(line,slide,spec,a,b);Snapshot(line,spec);Bounds(deck,new PointF(spec.Begin.X,spec.Begin.Y));Bounds(deck,new PointF(spec.End.X,spec.End.Y));var part=line.CustomerData.Add();if(!part.LoadXML(ConnectorCodec.Encode(spec)))throw new InvalidDataException("Connector metadata could not be saved.");line.Tags.Add("PPI_CONNECTOR","1");line.Name="NexWave connector "+line.Id;line.Select();return line;}catch(Exception error){if(line!=null)try{line.Delete();}catch{throw new InvalidOperationException("Connector creation failed; cleanup incomplete. Use Undo.",error);}throw;}
        }
        private static PowerPoint.Shape Selected(PowerPoint.Application app){var s=app.ActiveWindow.Selection;if(s.Type!=PowerPoint.PpSelectionType.ppSelectionShapes||s.ShapeRange.Count!=1||!IsManaged(s.ShapeRange[1]))throw new InvalidOperationException("Select one managed universal connector.");return s.ShapeRange[1];}
        internal static void Update(PowerPoint.Application app,ConnectorSpec spec)
        {
            var slide=Active(app);var line=Selected(app);spec.Validate();spec=spec.Copy();var a=Locate(slide,spec.Begin);var b=Locate(slide,spec.End);Bounds(app.ActivePresentation,a);Bounds(app.ActivePresentation,b);var saved=Saved.Capture(line);app.StartNewUndoEntry();try{Style(line,spec);Geometry(line,slide,spec,a,b);Snapshot(line,spec);Bounds(app.ActivePresentation,new PointF(spec.Begin.X,spec.Begin.Y));Bounds(app.ActivePresentation,new PointF(spec.End.X,spec.End.Y));saved.Part=XmlPartUpdates.Shape(line,saved.Part,ConnectorCodec.Encode(spec));}catch(Exception error){try{saved.Restore(slide);}catch(Exception restore){throw new InvalidOperationException("Connector update failed; restoration incomplete. Use Undo.",new AggregateException(error,restore));}throw;}
        }
        internal static void Reflow(PowerPoint.Application app,Action<string> checkpoint=null)
        {
            if(app==null||app.Presentations.Count==0||app.ActiveWindow.ViewType!=PowerPoint.PpViewType.ppViewNormal)return;var slide=Active(app);if(checkpoint!=null)checkpoint("before-enumeration");var lines=slide.Shapes.Cast<PowerPoint.Shape>().Where(IsManaged).Take(101).ToArray();if(checkpoint!=null)checkpoint("after-enumeration");if(lines.Length>100)throw new InvalidOperationException("Connector refresh is limited to 100 per slide.");
            var changes=new List<Tuple<Saved,ConnectorSpec,PointF,PointF,bool>>();foreach(var line in lines){var prior=Read(line);if(checkpoint!=null)checkpoint("after-metadata-read");var live=Live(line,prior);if(checkpoint!=null)checkpoint("after-native-binding-read");var start=Position(line,true);var end=Position(line,false);var a=live.Begin.Kind=="point"?Locate(slide,live.Begin):start;var b=live.End.Kind=="point"?Locate(slide,live.End):end;Bounds(app.ActivePresentation,a);Bounds(app.ActivePresentation,b);bool move=Math.Abs(a.X-start.X)>.1f||Math.Abs(a.Y-start.Y)>.1f||Math.Abs(b.X-end.X)>.1f||Math.Abs(b.Y-end.Y)>.1f;
                // Avoid materializing line color/theme/adjustment restoration state on a no-op.
                if(move||ConnectorCodec.Encode(live)!=ConnectorCodec.Encode(prior))changes.Add(Tuple.Create(Saved.Capture(line),live,a,b,move));}
            if(checkpoint!=null)checkpoint("planned-changes-"+changes.Count);var applied=new List<Saved>();try{foreach(var change in changes){var saved=change.Item1;applied.Add(saved);if(change.Item5){Geometry(saved.Line,slide,change.Item2,change.Item3,change.Item4);Snapshot(saved.Line,change.Item2);}string xml=ConnectorCodec.Encode(change.Item2);if(xml!=ConnectorCodec.Encode(ConnectorCodec.Decode(saved.Xml)))saved.Part=XmlPartUpdates.Shape(saved.Line,saved.Part,xml);}}
            catch(Exception original){bool restored=true;for(int i=applied.Count-1;i>=0;i--)try{applied[i].Restore(slide);}catch{restored=false;}if(!restored)throw new InvalidOperationException("Connector refresh failed and restoration was incomplete. Use Undo.",original);throw;}
        }
        internal static void Insert(PowerPoint.Application app){Show(app,false);}
        internal static void Edit(PowerPoint.Application app){Show(app,true);}
        private sealed class EndpointEditor
        {
            internal readonly GroupBox Box;private readonly ComboBox target,kind,anchor;private readonly NumericUpDown site,series,point,freeX,freeY;private readonly PowerPoint.Shape[] shapes;
            internal EndpointEditor(string title,PowerPoint.Shape[] choices,ConnectorEndpoint prior,int x)
            {
                shapes=choices;Box=new GroupBox {Text=title,Left=x,Top=15,Width=285,Height=310};target=new ComboBox {Left=10,Top=24,Width=262,DropDownStyle=ComboBoxStyle.DropDownList};foreach(var shape in shapes)target.Items.Add(shape.Name+" ["+shape.Id+"]");target.SelectedIndex=shapes.Length==0?-1:Math.Max(0,Array.FindIndex(shapes,s=>s.Id==prior.ShapeId));kind=new ComboBox {Left=10,Top=61,Width=262,DropDownStyle=ComboBoxStyle.DropDownList};kind.Items.AddRange(new object[]{"Native shape connection site","Native chart data point","Free slide coordinates"});kind.SelectedIndex=prior.Kind=="free"?2:prior.Kind=="point"?1:0;
                site=new NumericUpDown {Left=160,Top=99,Width=105,Minimum=0,Maximum=1024,Value=prior.Site};series=new NumericUpDown {Left=160,Top=133,Width=105,Minimum=1,Maximum=100,Value=prior.Series};point=new NumericUpDown {Left=160,Top=167,Width=105,Minimum=1,Maximum=10000,Value=prior.Point};anchor=new ComboBox {Left=160,Top=201,Width=105,DropDownStyle=ComboBoxStyle.DropDownList};anchor.Items.AddRange(new object[]{"center","top","right","bottom","left"});anchor.SelectedItem=prior.Anchor;
                freeX=new NumericUpDown {Left=160,Top=235,Width=105,Minimum=0,Maximum=10000,DecimalPlaces=2,Value=(decimal)prior.X};freeY=new NumericUpDown {Left=160,Top=269,Width=105,Minimum=0,Maximum=10000,DecimalPlaces=2,Value=(decimal)prior.Y};
                Box.Controls.AddRange(new Control[]{new Label {Text="Free X (pt)",Left=10,Top=239,Width=145},new Label {Text="Free Y (pt)",Left=10,Top=273,Width=145},freeX,freeY,target,kind,new Label {Text="Site (0 = auto)",Left=10,Top=103,Width=145},new Label {Text="Series (1-based)",Left=10,Top=137,Width=145},new Label {Text="Point (1-based)",Left=10,Top=171,Width=145},new Label {Text="Point anchor",Left=10,Top=205,Width=145},site,series,point,anchor});Action enabled=()=>{target.Enabled=kind.SelectedIndex!=2;site.Enabled=kind.SelectedIndex==0;series.Enabled=point.Enabled=anchor.Enabled=kind.SelectedIndex==1;freeX.Enabled=freeY.Enabled=kind.SelectedIndex==2;};kind.SelectedIndexChanged+=(s,e)=>enabled();enabled();
            }
            internal ConnectorEndpoint Read(){if(kind.SelectedIndex==2)return new ConnectorEndpoint {Kind="free",X=(float)freeX.Value,Y=(float)freeY.Value};if(target.SelectedIndex<0)throw new InvalidDataException("Select a native target or choose free coordinates.");return new ConnectorEndpoint {Kind=kind.SelectedIndex==0?"site":"point",ShapeId=shapes[target.SelectedIndex].Id,Site=kind.SelectedIndex==0?(int)site.Value:0,Series=(int)series.Value,Point=(int)point.Value,Anchor=(string)anchor.SelectedItem};}
        }
        private static void Show(PowerPoint.Application app,bool edit)
        {
            var slide=Active(app);var deck=app.ActivePresentation;var choices=slide.Shapes.Cast<PowerPoint.Shape>().Where(s=>s.Connector!=Office.MsoTriState.msoTrue).Take(501).ToArray();if(choices.Length>500)throw new InvalidOperationException("Choose a slide with at most 500 native shapes/charts.");var spec=edit?Live(Selected(app),Read(Selected(app))):choices.Length==0?new ConnectorSpec {Begin=new ConnectorEndpoint {Kind="free",X=100,Y=100},End=new ConnectorEndpoint {Kind="free",X=300,Y=200}}:new ConnectorSpec {Begin=new ConnectorEndpoint {ShapeId=choices[0].Id,Site=1},End=new ConnectorEndpoint {ShapeId=choices[Math.Min(1,choices.Length-1)].Id,Site=choices.Length==1?2:1}};
            var selection=app.ActiveWindow.Selection;if(!edit&&selection.Type==PowerPoint.PpSelectionType.ppSelectionShapes&&selection.ShapeRange.Count==2){spec.Begin.ShapeId=selection.ShapeRange[1].Id;spec.End.ShapeId=selection.ShapeRange[2].Id;}
            using(var form=new Form {Text="Universal connector · NexWave",Width=615,Height=555,StartPosition=FormStartPosition.CenterParent,FormBorderStyle=FormBorderStyle.FixedDialog,MaximizeBox=false,MinimizeBox=false})
            {var begin=new EndpointEditor("Start",choices,spec.Begin,12);var end=new EndpointEditor("End",choices,spec.End,305);var route=new ComboBox {Left=90,Top=344,Width=125,DropDownStyle=ComboBoxStyle.DropDownList};route.Items.AddRange(new object[]{"straight","elbow","curve"});route.SelectedItem=spec.Route;var dash=new ComboBox {Left=310,Top=344,Width=125,DropDownStyle=ComboBoxStyle.DropDownList};dash.Items.AddRange(new object[]{"solid","dash","dot","dashdot"});dash.SelectedItem=spec.Dash;var weight=new NumericUpDown {Left=90,Top=380,Width=125,Minimum=.25m,Maximum=12,DecimalPlaces=2,Value=(decimal)spec.Weight};var arrow=new ComboBox {Left=310,Top=380,Width=125,DropDownStyle=ComboBoxStyle.DropDownList};arrow.Items.AddRange(new object[]{"none","triangle","open","stealth"});arrow.SelectedItem=spec.EndArrow;var startArrow=new ComboBox {Left=90,Top=418,Width=125,DropDownStyle=ComboBoxStyle.DropDownList};startArrow.Items.AddRange(new object[]{"none","triangle","open","stealth"});startArrow.SelectedItem=spec.BeginArrow;var autoRoute=new CheckBox {Left=245,Top=420,Width=330,Text="Route automatically when both sites are auto",Checked=spec.AutoRoute};var color=new Button {Text="Line color",Left=455,Top=344,Width=120};int rgb=spec.Color;color.Click+=(s,e)=>{using(var picker=new ColorDialog {FullOpen=true})if(picker.ShowDialog()==DialogResult.OK)rgb=picker.Color.R|(picker.Color.G<<8)|(picker.Color.B<<16);};var again=new CheckBox {Text="Keep open for multiple connectors",Left=20,Top=462,Width=300,Enabled=!edit};var apply=new Button {Text=edit?"Apply":"Insert",Left=395,Top=462,Width=85};var cancel=new Button {Text="Close",Left=490,Top=462,Width=85,DialogResult=DialogResult.Cancel};form.Controls.AddRange(new Control[]{new Label {Text="Start arrow",Left=20,Top=422},startArrow,autoRoute,begin.Box,end.Box,new Label {Text="Route",Left=20,Top=348},new Label {Text="Dash",Left=245,Top=348},new Label {Text="Weight (pt)",Left=20,Top=384},new Label {Text="End arrow",Left=245,Top=384},route,dash,weight,arrow,color,again,apply,cancel});form.AcceptButton=apply;form.CancelButton=cancel;
                apply.Click+=(s,e)=>{try{if(app.ActivePresentation!=deck||((PowerPoint.Slide)app.ActiveWindow.View.Slide).SlideID!=slide.SlideID)throw new InvalidOperationException("Return to the slide that opened this connector editor.");var choice=new ConnectorSpec {Begin=begin.Read(),End=end.Read(),Route=(string)route.SelectedItem,Dash=(string)dash.SelectedItem,Weight=(float)weight.Value,Color=rgb,BeginArrow=(string)startArrow.SelectedItem,EndArrow=(string)arrow.SelectedItem,AutoRoute=autoRoute.Checked};if(edit)Update(app,choice);else Create(app,choice);if(!again.Checked)form.DialogResult=DialogResult.OK;}catch(Exception error) when(error is InvalidDataException||error is InvalidOperationException||error is System.Runtime.InteropServices.COMException||error is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException){MessageBox.Show(form,error.Message,"Connector");}};form.ShowDialog();}
        }
    }
}




