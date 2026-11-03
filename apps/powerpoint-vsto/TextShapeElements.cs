using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Office=Microsoft.Office.Core;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
namespace PowerPointIntelligence.AddIn
{
    [Flags] internal enum TextEdgeLocks {None=0,Left=1,Right=2,Top=4,Bottom=8}
    internal sealed class TextShapeSpec
    {
        internal string Kind="textbox";
        internal TextEdgeLocks Locks=TextEdgeLocks.Left|TextEdgeLocks.Right|TextEdgeLocks.Top;
        internal float Adjustment=.25f;
        internal bool AutoFit=true;
        internal TextShapeSpec Copy(){return new TextShapeSpec {Kind=Kind,Locks=Locks,Adjustment=Adjustment,AutoFit=AutoFit};}
    }
    internal static class TextShapePlan
    {
        internal static string Kind(string kind)
        {
            string key=(kind??"").Replace("-","").Replace(" ","").ToLowerInvariant();
            if(key=="text"||key=="textbox")return "textbox";if(key=="pentagon"||key=="rightpentagon")return "rightPentagon";if(key=="chevron")return "chevron";if(key=="roundedrectangle"||key=="roundedrect")return "roundedRectangle";
            throw new InvalidDataException("Choose Text Box, Right Pentagon, Chevron or Rounded Rectangle.");
        }
        internal static bool Finite(float n){return !Single.IsNaN(n)&&!Single.IsInfinity(n);}
        internal static void Validate(TextShapeSpec s){if(s==null||Kind(s.Kind)!=s.Kind||((int)s.Locks&~15)!=0||!Finite(s.Adjustment)||s.Adjustment<.05f||s.Adjustment>.5f)throw new InvalidDataException("Invalid text element settings.");}
        internal static float ParseLength(string input,float slideWidth,float slideHeight)
        {
            if(!Finite(slideWidth)||!Finite(slideHeight)||slideWidth<=0||slideHeight<=0||slideWidth>10000||slideHeight>10000)throw new InvalidDataException("Invalid slide dimensions.");
            if(input==null||input.Length>64)throw new InvalidDataException("Enter a bounded length such as 24 pt, 2 cm, 1 in or 20%w.");
            var fraction=Regex.Match(input.Trim(),@"\A([0-9]+(?:\.[0-9]+)?)\s*/\s*([0-9]+(?:\.[0-9]+)?)\s*(sw|sh|slidewidth|slideheight)\z",RegexOptions.IgnoreCase|RegexOptions.CultureInvariant);
            if(fraction.Success){double numerator,denominator;if(!Double.TryParse(fraction.Groups[1].Value,NumberStyles.AllowDecimalPoint,CultureInfo.InvariantCulture,out numerator)||!Double.TryParse(fraction.Groups[2].Value,NumberStyles.AllowDecimalPoint,CultureInfo.InvariantCulture,out denominator)||denominator<=0)throw new InvalidDataException("Use a positive fraction denominator.");double result=numerator/denominator*(fraction.Groups[3].Value.Equals("sw",StringComparison.OrdinalIgnoreCase)||fraction.Groups[3].Value.Equals("slidewidth",StringComparison.OrdinalIgnoreCase)?slideWidth:slideHeight);if(Double.IsNaN(result)||Double.IsInfinity(result)||result<1||result>10000)throw new InvalidDataException("Lengths must be between 1 and 10000 points.");return (float)result;}
            var match=Regex.Match(input.Trim(),@"\A([0-9]+(?:\.[0-9]+)?)\s*(pt|mm|cm|in|%w|%h|sw|sh|slidewidth|slideheight)?\z",RegexOptions.IgnoreCase|RegexOptions.CultureInvariant);
            double value;if(!match.Success||!Double.TryParse(match.Groups[1].Value,NumberStyles.AllowDecimalPoint,CultureInfo.InvariantCulture,out value))throw new InvalidDataException("Use invariant decimal lengths with pt, mm, cm, in, %w or %h units.");
            string unit=match.Groups[2].Value.ToLowerInvariant();double factor=unit=="mm"?72/25.4:unit=="cm"?72/2.54:unit=="in"?72:unit=="%w"?slideWidth/100.0:unit=="%h"?slideHeight/100.0:unit=="sw"||unit=="slidewidth"?slideWidth:unit=="sh"||unit=="slideheight"?slideHeight:1;
            double points=value*factor;if(Double.IsNaN(points)||Double.IsInfinity(points)||points<1||points>10000)throw new InvalidDataException("Lengths must be between 1 and 10000 points.");return (float)points;
        }
        internal static RectangleF Fit(RectangleF bounds,SizeF required,TextEdgeLocks locks,float slideWidth,float slideHeight)
        {
            if(((int)locks&~15)!=0||new[]{bounds.X,bounds.Y,bounds.Width,bounds.Height,required.Width,required.Height,slideWidth,slideHeight}.Any(n=>!Finite(n))||bounds.X<0||bounds.Y<0||bounds.Width<1||bounds.Height<1||required.Width<0||required.Height<0||slideWidth<1||slideHeight<1||slideWidth>10000||slideHeight>10000)throw new InvalidDataException("Invalid fitting geometry.");
            float width=Math.Max(1,required.Width),height=Math.Max(1,required.Height);bool left=locks.HasFlag(TextEdgeLocks.Left),right=locks.HasFlag(TextEdgeLocks.Right),top=locks.HasFlag(TextEdgeLocks.Top),bottom=locks.HasFlag(TextEdgeLocks.Bottom);
            if(left&&right&&required.Width>bounds.Width+.5f||top&&bottom&&required.Height>bounds.Height+.5f)throw new InvalidOperationException("Text overflows locked edges. Unlock an edge or enlarge the element; font size was preserved.");
            if(left&&right)width=bounds.Width;if(top&&bottom)height=bounds.Height;
            float x=left?bounds.Left:right?bounds.Right-width:bounds.Left+(bounds.Width-width)/2,y=top?bounds.Top:bottom?bounds.Bottom-height:bounds.Top+(bounds.Height-height)/2;
            if(x<0||y<0||x+width>slideWidth+.01f||y+height>slideHeight+.01f)throw new InvalidOperationException("Text fitting would leave the slide. Reposition the element or choose different locks.");return new RectangleF(x,y,width,height);
        }
    }
    internal static class TextShapeCodec
    {
        internal const string Namespace="urn:powerpoint-intelligence:text-element:v1";
        internal static string Encode(TextShapeSpec spec){TextShapePlan.Validate(spec);return new XElement(XName.Get("textElement",Namespace),new XAttribute("version","1"),new XAttribute("kind",spec.Kind),new XAttribute("locks",(int)spec.Locks),new XAttribute("adjustment",spec.Adjustment.ToString("R",CultureInfo.InvariantCulture)),new XAttribute("autoFit",spec.AutoFit?"true":"false")).ToString(SaveOptions.DisableFormatting);}
        internal static TextShapeSpec Decode(string xml)
        {
            if(String.IsNullOrEmpty(xml)||xml.Length>4096)throw new InvalidDataException("Missing or excessive text element metadata.");
            try{XDocument doc;using(var r=XmlReader.Create(new StringReader(xml),new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=4096,IgnoreComments=false,IgnoreProcessingInstructions=false}))doc=XDocument.Load(r);
                var root=doc.Root;var names=new[]{"version","kind","locks","adjustment","autoFit"};if(root==null||root.Name!=XName.Get("textElement",Namespace)||root.Nodes().Any()||doc.Nodes().Any(n=>!(n is XElement))||root.Attributes().Any(a=>!a.IsNamespaceDeclaration&&(a.Name.NamespaceName!=""||!names.Contains(a.Name.LocalName)))||names.Any(n=>root.Attribute(n)==null)||(string)root.Attribute("version")!="1")throw new InvalidDataException("Unknown text element XML schema or fields.");
                int locks;float adjustment;string fit=(string)root.Attribute("autoFit");if(!Int32.TryParse((string)root.Attribute("locks"),NumberStyles.None,CultureInfo.InvariantCulture,out locks)||!Single.TryParse((string)root.Attribute("adjustment"),NumberStyles.Float,CultureInfo.InvariantCulture,out adjustment)||fit!="true"&&fit!="false")throw new InvalidDataException("Malformed text element settings.");var spec=new TextShapeSpec {Kind=(string)root.Attribute("kind"),Locks=(TextEdgeLocks)locks,Adjustment=adjustment,AutoFit=fit=="true"};TextShapePlan.Validate(spec);return spec;
            }catch(XmlException ex){throw new InvalidDataException("Malformed text element XML.",ex);}
        }
    }
    internal static class TextShapeElements
    {
        internal const string Tag="PPI_TEXT_ELEMENT";
        internal static bool IsManaged(PowerPoint.Shape shape){return shape!=null&&!String.IsNullOrEmpty(shape.Tags[Tag]);}
        private static PowerPoint.Slide Active(PowerPoint.Application app){if(app==null||app.Presentations.Count==0||app.ActiveWindow.ViewType!=PowerPoint.PpViewType.ppViewNormal)throw new InvalidOperationException("Open a presentation in Normal view.");var slide=app.ActiveWindow.View.Slide as PowerPoint.Slide;if(slide==null)throw new InvalidOperationException("Select a slide.");return slide;}
        private static PowerPoint.Shape Selected(PowerPoint.Application app){Active(app);var selected=app.ActiveWindow.Selection;if(selected.Type!=PowerPoint.PpSelectionType.ppSelectionShapes&&selected.Type!=PowerPoint.PpSelectionType.ppSelectionText||selected.ShapeRange.Count!=1)throw new InvalidOperationException("Select one native managed text element.");var shape=selected.ShapeRange[1];Geometry(shape);return shape;}
        private static void Geometry(PowerPoint.Shape shape){if(shape.Type!=Office.MsoShapeType.msoAutoShape||shape.HasTextFrame!=Office.MsoTriState.msoTrue||Math.Abs(shape.Rotation)>.01||shape.HorizontalFlip==Office.MsoTriState.msoTrue||shape.VerticalFlip==Office.MsoTriState.msoTrue||shape.TextFrame.Orientation!=Office.MsoTextOrientation.msoTextOrientationHorizontal)throw new InvalidOperationException("Use one unrotated horizontal native AutoShape. Grouped, flipped and native text-box conversion are not supported.");if(shape.LockAspectRatio==Office.MsoTriState.msoTrue||shape.TextFrame.TextRange.Length>2000)throw new InvalidOperationException("Unlock native aspect ratio and use at most 2000 text characters.");}
        private static Office.MsoAutoShapeType Native(string kind){return kind=="rightPentagon"?Office.MsoAutoShapeType.msoShapePentagon:kind=="chevron"?Office.MsoAutoShapeType.msoShapeChevron:kind=="roundedRectangle"?Office.MsoAutoShapeType.msoShapeRoundedRectangle:Office.MsoAutoShapeType.msoShapeRectangle;}
        private static Office.CustomXMLPart Part(PowerPoint.Shape shape){var parts=XmlPartUpdates.AttachedParts(shape,TextShapeCodec.Namespace);if(parts.Length!=1)throw new InvalidDataException("Text element metadata is missing or ambiguous (canonicalMatches="+parts.Length+"); objects were preserved.");return parts[0];}
        internal static TextShapeSpec Read(PowerPoint.Shape shape){if(!IsManaged(shape)||shape.Tags[Tag]!="1")throw new InvalidDataException("Select a supported managed text element.");Geometry(shape);var spec=TextShapeCodec.Decode(Part(shape).XML);if(shape.AutoShapeType!=Native(spec.Kind))throw new InvalidDataException("Native shape geometry differs from its managed kind. Reconvert the element explicitly.");return spec;}
        private static RectangleF Bounds(PowerPoint.Shape shape){return new RectangleF(shape.Left,shape.Top,shape.Width,shape.Height);}
        private static void Bounds(PowerPoint.Shape shape,RectangleF value){shape.Width=value.Width;shape.Height=value.Height;shape.Left=value.Left;shape.Top=value.Top;}
        private static void Adjust(PowerPoint.Shape shape,TextShapeSpec spec){if(spec.Kind!="textbox"){if(shape.Adjustments.Count!=1)throw new InvalidOperationException("Native adjustment geometry is unavailable.");shape.Adjustments[1]=spec.Adjustment;}}
        private static SizeF Measure(PowerPoint.Shape shape,TextShapeSpec spec)
        {
            PowerPoint.Application application=(PowerPoint.Application)shape.Application;PowerPoint.DocumentWindow window=application.ActiveWindow;PowerPoint.Selection selected=window.Selection;
            PowerPoint.ShapeRange selectedShapes=null;PowerPoint.TextRange selectedText=null;
            if(selected.Type==PowerPoint.PpSelectionType.ppSelectionShapes)selectedShapes=selected.ShapeRange;
            else if(selected.Type==PowerPoint.PpSelectionType.ppSelectionText)selectedText=selected.TextRange;
            PowerPoint.Shape probe=null;try{probe=shape.Duplicate()[1];probe.TextFrame2.AutoSize=Office.MsoAutoSize.msoAutoSizeNone;probe.TextFrame.AutoSize=PowerPoint.PpAutoSize.ppAutoSizeNone;if(!(spec.Locks.HasFlag(TextEdgeLocks.Left)&&spec.Locks.HasFlag(TextEdgeLocks.Right)))probe.TextFrame2.WordWrap=Office.MsoTriState.msoFalse;var frame=probe.TextFrame2;float w=frame.TextRange.BoundWidth+frame.MarginLeft+frame.MarginRight+2,h=frame.TextRange.BoundHeight+frame.MarginTop+frame.MarginBottom+2;if(!TextShapePlan.Finite(w)||!TextShapePlan.Finite(h))throw new InvalidDataException("Native text measurements are unavailable.");return new SizeF(w,h);}finally{if(probe!=null){probe.Delete();if(selectedText!=null)selectedText.Select();else if(selectedShapes!=null)selectedShapes.Select();}}
        }
        private static RectangleF Plan(PowerPoint.Application app,PowerPoint.Shape shape,TextShapeSpec spec){return TextShapePlan.Fit(Bounds(shape),Measure(shape,spec),spec.Locks,app.ActivePresentation.PageSetup.SlideWidth,app.ActivePresentation.PageSetup.SlideHeight);}
        private static void InitializeNewShape(PowerPoint.Shape shape,string xml)
        {
            // Only called for a just-created, add-in-owned shape. Never normalize a user's existing object.
            var inherited=XmlPartUpdates.AttachedParts(shape,TextShapeCodec.Namespace);
            if(inherited.Length>0){
                var slide=(PowerPoint.Slide)shape.Parent;var deck=(PowerPoint.Presentation)slide.Parent;var ids=new HashSet<string>(inherited.Select(p=>Guid.Parse(p.Id).ToString("D")),StringComparer.OrdinalIgnoreCase);int visited=0;
                Action<PowerPoint.Shape,int> inspect=null;inspect=(candidate,depth)=>{if(++visited>10000||depth>8)throw new InvalidDataException("New-shape ownership check exceeds bounds.");bool same=candidate.Id==shape.Id&&candidate.Parent is PowerPoint.Slide&&((PowerPoint.Slide)candidate.Parent).SlideID==slide.SlideID;if(!same){int parts=0;foreach(Office.CustomXMLPart p in candidate.CustomerData){if(++parts>64)throw new InvalidDataException("New-shape part ownership check exceeds bounds.");if(ids.Contains(Guid.Parse(p.Id).ToString("D")))throw new InvalidDataException("New shape inherited a shared metadata attachment; no existing owner was changed.");}}if(candidate.Type==Office.MsoShapeType.msoGroup)foreach(PowerPoint.Shape child in candidate.GroupItems)inspect(child,depth+1);};
                foreach(PowerPoint.Slide ownerSlide in deck.Slides)foreach(PowerPoint.Shape candidate in ownerSlide.Shapes)inspect(candidate,0);
                foreach(var part in inherited)shape.CustomerData.Delete(part.Id);
            }
            var added=shape.CustomerData.Add();if(!added.LoadXML(xml))throw new InvalidDataException("Could not initialize text element settings.");
            var final=XmlPartUpdates.AttachedParts(shape,TextShapeCodec.Namespace);if(final.Length!=1||!XmlPartUpdates.Equivalent(final[0].XML,xml))throw new InvalidDataException("New text metadata verification failed; inherited="+inherited.Length+"; finalCanonical="+final.Length+". Existing objects were not normalized.");
        }
        internal static PowerPoint.Shape Insert(PowerPoint.Application app,string kind,Action<PowerPoint.Shape> initializeNativeStyle=null)
        {
            var slide=Active(app);var spec=new TextShapeSpec {Kind=TextShapePlan.Kind(kind)};float w=app.ActivePresentation.PageSetup.SlideWidth,h=app.ActivePresentation.PageSetup.SlideHeight;if(w<240||h<140)throw new InvalidOperationException("The slide is too small for a text element.");PowerPoint.Shape shape=null;app.StartNewUndoEntry();try{shape=slide.Shapes.AddShape(Native(spec.Kind),(w-200)/2,(h-70)/2,200,70);shape.TextFrame.TextRange.Text="";shape.TextFrame.TextRange.Font.Size=20;shape.TextFrame.TextRange.Font.Name="Segoe UI";shape.TextFrame.TextRange.ParagraphFormat.Alignment=PowerPoint.PpParagraphAlignment.ppAlignCenter;shape.TextFrame.VerticalAnchor=Office.MsoVerticalAnchor.msoAnchorMiddle;shape.TextFrame.AutoSize=PowerPoint.PpAutoSize.ppAutoSizeNone;shape.TextFrame2.AutoSize=Office.MsoAutoSize.msoAutoSizeNone;shape.TextFrame.WordWrap=Office.MsoTriState.msoTrue;shape.TextFrame.MarginLeft=shape.TextFrame.MarginRight=12;shape.TextFrame.MarginTop=shape.TextFrame.MarginBottom=8;shape.Fill.Solid();shape.Fill.ForeColor.ObjectThemeColor=Office.MsoThemeColorIndex.msoThemeColorAccent1;shape.Line.ForeColor.ObjectThemeColor=Office.MsoThemeColorIndex.msoThemeColorAccent1;if(spec.Kind=="textbox"){shape.Fill.Visible=Office.MsoTriState.msoFalse;shape.Line.Visible=Office.MsoTriState.msoFalse;}Adjust(shape,spec);if(initializeNativeStyle!=null){initializeNativeStyle(shape);Adjust(shape,spec);}InitializeNewShape(shape,TextShapeCodec.Encode(spec));shape.Tags.Add(Tag,"1");shape.Name="NexWave "+spec.Kind+" "+shape.Id;Read(shape);shape.Select();shape.TextFrame.TextRange.Select();return shape;}catch(Exception original){if(shape!=null)try{shape.Delete();}catch(Exception cleanup){throw new InvalidOperationException("Element insertion failed and cleanup was incomplete. Use Undo.",new AggregateException(original,cleanup));}throw;}
        }
        internal static void PreflightGridGeometry(PowerPoint.Application app,PowerPoint.Shape source,string kind,RectangleF bounds)
        {
            var spec=Read(source).Copy();spec.Kind=TextShapePlan.Kind(kind);spec.AutoFit=false;spec.Locks=(TextEdgeLocks)15;
            var selection=app.ActiveWindow.Selection;PowerPoint.ShapeRange selectedShapes=null;PowerPoint.TextRange selectedText=null;
            if(selection.Type==PowerPoint.PpSelectionType.ppSelectionShapes)selectedShapes=selection.ShapeRange;else if(selection.Type==PowerPoint.PpSelectionType.ppSelectionText)selectedText=selection.TextRange;
            PowerPoint.Shape probe=null;
            try{probe=source.Duplicate()[1];probe.AutoShapeType=Native(spec.Kind);Adjust(probe,spec);Bounds(probe,bounds);Plan(app,probe,spec);}
            finally{if(probe!=null){probe.Delete();if(selectedText!=null)selectedText.Select();else if(selectedShapes!=null)selectedShapes.Select();}}
            // Never attach, convert through managed APIs, or persist settings on a temporary probe.
        }
        internal static PowerPoint.Shape Configure(PowerPoint.Application app,TextShapeSpec requested)
        {
            var shape=Selected(app);var prior=Read(shape);TextShapePlan.Validate(requested);if(requested.Kind!=prior.Kind)throw new InvalidOperationException("Use Convert Element to change its kind.");var part=Part(shape);string xml=part.XML;var bounds=Bounds(shape);float oldAdjust=shape.Adjustments.Count==1?shape.Adjustments[1]:0;var oldAuto=shape.TextFrame2.AutoSize;app.StartNewUndoEntry();try{Adjust(shape,requested);var fitSpec=requested.Copy();if(!requested.AutoFit)fitSpec.Locks=(TextEdgeLocks)15;var fit=Plan(app,shape,fitSpec);shape.TextFrame2.AutoSize=Office.MsoAutoSize.msoAutoSizeNone;Bounds(shape,fit);part=XmlPartUpdates.Shape(shape,part,TextShapeCodec.Encode(requested));return shape;}catch(Exception original){bool restored=true;try{Bounds(shape,bounds);if(shape.Adjustments.Count==1)shape.Adjustments[1]=oldAdjust;shape.TextFrame2.AutoSize=oldAuto;part=XmlPartUpdates.Shape(shape,Part(shape),xml);}catch{restored=false;}if(!restored)throw new InvalidOperationException("Element configuration failed and rollback was incomplete. Use Undo.",original);throw;}
        }
        internal static PowerPoint.Shape FitSelected(PowerPoint.Application app)
        {
            Active(app);var selection=app.ActiveWindow.Selection;if(selection.Type!=PowerPoint.PpSelectionType.ppSelectionShapes||selection.ShapeRange.Count<1||selection.ShapeRange.Count>100)throw new InvalidOperationException("Select 1 to 100 managed text elements.");
            var shapes=selection.ShapeRange.Cast<PowerPoint.Shape>().ToArray();var plans=shapes.Select(s=>Plan(app,s,Read(s))).ToArray();var old=shapes.Select(Bounds).ToArray();app.StartNewUndoEntry();try{for(int i=0;i<shapes.Length;i++)Bounds(shapes[i],plans[i]);}catch(Exception original){bool restored=true;for(int i=0;i<shapes.Length;i++)try{Bounds(shapes[i],old[i]);}catch{restored=false;}if(!restored)throw new InvalidOperationException("Text fitting failed and restoration was incomplete. Use Undo.",original);throw;}return shapes[0];
        }
        internal static PowerPoint.Shape DuplicateSelected(PowerPoint.Application app,string direction)
        {
            var source=Selected(app);var spec=Read(source);if(!new[]{"left","right","up","down"}.Contains(direction))throw new InvalidDataException("Choose left, right, up or down.");var rect=Bounds(source);const float gap=12;
            if(direction=="left")rect.X-=rect.Width+gap;else if(direction=="right")rect.X+=rect.Width+gap;else if(direction=="up")rect.Y-=rect.Height+gap;else rect.Y+=rect.Height+gap;
            TextShapePlan.Fit(rect,rect.Size,(TextEdgeLocks)15,app.ActivePresentation.PageSetup.SlideWidth,app.ActivePresentation.PageSetup.SlideHeight);
            if(direction=="right"&&spec.Kind=="rightPentagon")spec.Kind="chevron";
            PowerPoint.Shape copy=null;app.StartNewUndoEntry();try{copy=source.Duplicate()[1];var oldPart=Part(copy);string original=Part(source).XML;var newXml=TextShapeCodec.Encode(spec);if(newXml==oldPart.XML){copy.CustomerData.Delete(oldPart.Id);if(XmlPartUpdates.AttachedParts(source,TextShapeCodec.Namespace).Length==0){if(!source.CustomerData.Add().LoadXML(original))throw new InvalidDataException("Original duplicate owner metadata restoration rejected.");}var independent=copy.CustomerData.Add();if(!independent.LoadXML(newXml))throw new InvalidDataException("Duplicate metadata rejected.");}else XmlPartUpdates.Shape(copy,oldPart,newXml);copy.AutoShapeType=Native(spec.Kind);Adjust(copy,spec);Bounds(copy,rect);copy.Name="NexWave "+spec.Kind+" "+copy.Id;copy.Select();return copy;}catch(Exception original){if(copy!=null)try{copy.Delete();}catch(Exception cleanup){throw new InvalidOperationException("Duplication failed and cleanup was incomplete. Use Undo.",new AggregateException(original,cleanup));}throw;}
        }
        internal static PowerPoint.Shape ConvertSelected(PowerPoint.Application app,string kind)
        {
            var shape=Selected(app);var prior=Read(shape);var spec=prior.Copy();spec.Kind=TextShapePlan.Kind(kind);var type=shape.AutoShapeType;float adjust=shape.Adjustments.Count==1?shape.Adjustments[1]:0;string xml=Part(shape).XML;var bounds=Bounds(shape);app.StartNewUndoEntry();try{shape.AutoShapeType=Native(spec.Kind);Adjust(shape,spec);Bounds(shape,Plan(app,shape,spec));XmlPartUpdates.Shape(shape,Part(shape),TextShapeCodec.Encode(spec));return shape;}catch(Exception original){bool restored=true;try{shape.AutoShapeType=type;if(shape.Adjustments.Count==1)shape.Adjustments[1]=adjust;Bounds(shape,bounds);XmlPartUpdates.Shape(shape,Part(shape),xml);}catch{restored=false;}if(!restored)throw new InvalidOperationException("Conversion failed and rollback was incomplete. Use Undo.",original);throw;}
        }
        private static readonly Dictionary<string,string> FitSnapshots=new Dictionary<string,string>();
        private static string FitSnapshot(PowerPoint.Shape shape,TextShapeSpec spec)
        {var frame=shape.TextFrame2;var range=frame.TextRange;return TextShapeCodec.Encode(spec)+"|"+range.Text+"|"+String.Join("|",new[]{shape.Left,shape.Top,shape.Width,shape.Height,range.BoundWidth,range.BoundHeight,frame.MarginLeft,frame.MarginRight,frame.MarginTop,frame.MarginBottom}.Select(n=>n.ToString("R",CultureInfo.InvariantCulture)));}
        internal static int AutoReflow(PowerPoint.Application app)
        {
            var slide=Active(app);if(app.ActiveWindow.Selection.Type==PowerPoint.PpSelectionType.ppSelectionText)return 0;var shapes=slide.Shapes.Cast<PowerPoint.Shape>().Where(s=>IsManaged(s)&&s.Visible==Office.MsoTriState.msoTrue&&String.IsNullOrEmpty(s.Tags[ProcessFlowTools.GridTag])).Take(101).ToArray();if(shapes.Length>100)throw new InvalidOperationException("Automatic text fitting supports at most 100 managed elements per slide.");string prefix=app.ActiveWindow.HWND+":"+app.ActivePresentation.Name+":"+slide.SlideID+":";var changed=shapes.Select(s=>new {Shape=s,Spec=Read(s)}).Where(p=>p.Spec.AutoFit).Where(p=>{string oldValue;return !FitSnapshots.TryGetValue(prefix+p.Shape.Id,out oldValue)||oldValue!=FitSnapshot(p.Shape,p.Spec);}).ToArray();var plans=changed.Select(p=>new {p.Shape,Bounds=Plan(app,p.Shape,p.Spec)}).ToArray();var old=plans.Select(p=>Bounds(p.Shape)).ToArray();try{for(int i=0;i<plans.Length;i++){var b=plans[i].Bounds;var before=old[i];if(Math.Abs(before.X-b.X)>.1||Math.Abs(before.Y-b.Y)>.1||Math.Abs(before.Width-b.Width)>.1||Math.Abs(before.Height-b.Height)>.1)Bounds(plans[i].Shape,b);}}catch(Exception original){bool restored=true;for(int i=0;i<plans.Length;i++)try{Bounds(plans[i].Shape,old[i]);}catch{restored=false;}if(!restored)throw new InvalidOperationException("Automatic fitting failed and rollback was incomplete. Use Undo.",original);throw;}if(FitSnapshots.Count>1000)FitSnapshots.Clear();foreach(var entry in changed)FitSnapshots[prefix+entry.Shape.Id]=FitSnapshot(entry.Shape,entry.Spec);return plans.Length;
        }
        private static void DialogAction(Form form,Action action){try{action();}catch(Exception error){MessageBox.Show(form,error.Message,"Text element");}}
        internal static void Edit(PowerPoint.Application app)
        {
            var shape=Selected(app);var spec=Read(shape);using(var form=new Form {Text="Text Element — NexWave",Width=470,Height=360,StartPosition=FormStartPosition.CenterParent,FormBorderStyle=FormBorderStyle.FixedDialog,MaximizeBox=false,MinimizeBox=false})
            {
                var width=new TextBox {Left=95,Top=20,Width=110,Text=shape.Width.ToString("0.##",CultureInfo.InvariantCulture)+" pt"};var height=new TextBox {Left=300,Top=20,Width=110,Text=shape.Height.ToString("0.##",CultureInfo.InvariantCulture)+" pt"};var adjustment=new NumericUpDown {Left=150,Top=65,Minimum=5,Maximum=50,Value=(decimal)(spec.Adjustment*100),Enabled=spec.Kind!="textbox"};var auto=new CheckBox {Left=30,Top=110,Width=360,Text="Fit unlocked edges; preserve font size",Checked=spec.AutoFit};var locks=new[]{new CheckBox {Text="Left",Left=30},new CheckBox {Text="Right",Left=130},new CheckBox {Text="Top",Left=230},new CheckBox {Text="Bottom",Left=330}};var bits=new[]{TextEdgeLocks.Left,TextEdgeLocks.Right,TextEdgeLocks.Top,TextEdgeLocks.Bottom};for(int i=0;i<4;i++){locks[i].Top=150;locks[i].Width=90;locks[i].Checked=spec.Locks.HasFlag(bits[i]);}var apply=new Button {Text="Apply",Left=240,Top=270,Width=80};var close=new Button {Text="Close",Left=335,Top=270,Width=80,DialogResult=DialogResult.Cancel};var fill=new Button {Text="Fill",Left=30,Top=205,Width=85};var outline=new Button {Text="Outline",Left=125,Top=205,Width=85};var text=new Button {Text="Text color",Left=220,Top=205,Width=90};var center=new Button {Text="Center",Left=320,Top=205,Width=85};form.Controls.AddRange(new Control[]{new Label {Text="Width",Left=30,Top=24},width,new Label {Text="Height",Left=240,Top=24},height,new Label {Text="Geometry %",Left=30,Top=69},adjustment,auto,apply,close,fill,outline,text,center});form.Controls.AddRange(locks);form.CancelButton=close;
                fill.Click+=(s,e)=>DialogAction(form,()=>{shape.Select();NativeFormatting.Color(app,"fill");});outline.Click+=(s,e)=>DialogAction(form,()=>{shape.Select();NativeFormatting.Color(app,"outline");});text.Click+=(s,e)=>DialogAction(form,()=>{shape.Select();NativeFormatting.Color(app,"text");});center.Click+=(s,e)=>DialogAction(form,()=>{shape.Select();NativeFormatting.Text(app,"center");NativeFormatting.Text(app,"middle");});
                apply.Click+=(s,e)=>{var old=Bounds(shape);try{float w=TextShapePlan.ParseLength(width.Text,app.ActivePresentation.PageSetup.SlideWidth,app.ActivePresentation.PageSetup.SlideHeight),h=TextShapePlan.ParseLength(height.Text,app.ActivePresentation.PageSetup.SlideWidth,app.ActivePresentation.PageSetup.SlideHeight);var proposal=new TextShapeSpec {Kind=spec.Kind,Adjustment=(float)adjustment.Value/100,AutoFit=auto.Checked,Locks=TextEdgeLocks.None};for(int i=0;i<4;i++)if(locks[i].Checked)proposal.Locks|=bits[i];var rect=new RectangleF(old.X,old.Y,w,h);TextShapePlan.Fit(rect,new SizeF(w,h),(TextEdgeLocks)15,app.ActivePresentation.PageSetup.SlideWidth,app.ActivePresentation.PageSetup.SlideHeight);Bounds(shape,rect);shape.Select();Configure(app,proposal);form.DialogResult=DialogResult.OK;}catch(Exception error){try{Bounds(shape,old);}catch(Exception rollback){throw new InvalidOperationException("Element edit failed; geometry restoration failed. Use Undo.",new AggregateException(error,rollback));}MessageBox.Show(form,error.Message,"Text element");}};form.ShowDialog();
            }
        }
    }
}




