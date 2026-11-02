using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using Office=Microsoft.Office.Core;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
namespace PowerPointIntelligence.AddIn
{
    internal sealed class ProcessGridCell {internal int Row,Column,ShapeId;internal bool Active=true;internal ProcessGridCell Copy(){return new ProcessGridCell {Row=Row,Column=Column,ShapeId=ShapeId,Active=Active};}}
    internal sealed class ProcessGridSpec
    {
        internal string Id=Guid.NewGuid().ToString("D");internal int SlideId,Rows=2,Columns=3;
        internal float Left=40,Top=100,Width=880,Height=126,Gap=8;
        internal List<ProcessGridCell> Cells=new List<ProcessGridCell>();
        internal ProcessGridSpec Copy(){return new ProcessGridSpec {Id=Id,SlideId=SlideId,Rows=Rows,Columns=Columns,Left=Left,Top=Top,Width=Width,Height=Height,Gap=Gap,Cells=Cells.Select(c=>c.Copy()).ToList()};}
    }
    internal static class ProcessGridPlan
    {
        internal static void Validate(ProcessGridSpec s,float slideWidth,float slideHeight,bool persisted=true)
        {
            Guid id;if(s==null||!Guid.TryParse(s.Id,out id)||id==Guid.Empty||s.SlideId<1||s.Rows<1||s.Rows>20||s.Columns<1||s.Columns>12||s.Cells==null||s.Cells.Count>1000||new[]{s.Left,s.Top,s.Width,s.Height,s.Gap,slideWidth,slideHeight}.Any(n=>!TextShapePlan.Finite(n))||s.Left<0||s.Top<0||s.Gap<0||s.Gap>32||s.Width<1||s.Height<1||slideWidth<1||slideHeight<1||slideWidth>10000||slideHeight>10000||s.Left+s.Width>slideWidth+.01||s.Top+s.Height>slideHeight+.01)throw new InvalidDataException("Invalid or excessive process grid geometry.");
            var active=s.Cells.Where(c=>c!=null&&c.Active).ToArray();if(s.Cells.Any(c=>c==null||c.ShapeId<0||persisted&&c.ShapeId==0||c.Active&&(c.Row<0||c.Row>=s.Rows||c.Column<0||c.Column>=s.Columns)||!c.Active&&(c.Row!=-1||c.Column!=-1))||s.Cells.Where(c=>c.ShapeId>0).Select(c=>c.ShapeId).Distinct().Count()!=s.Cells.Count(c=>c.ShapeId>0)||active.Length!=s.Rows*s.Columns||active.Select(c=>c.Row+":"+c.Column).Distinct().Count()!=active.Length)throw new InvalidDataException("Grid cells must form one complete rectangular layout with unique identities.");
            if((s.Width-s.Gap*(s.Columns-1))/s.Columns<48||(s.Height-s.Gap*(s.Rows-1))/s.Rows<28)throw new InvalidOperationException("Process cells need at least 48 by 28 points. Enlarge the grid or use fewer cells.");
        }
        internal static string Kind(ProcessGridCell c){return c.Row==0?(c.Column==0?"rightPentagon":"chevron"):"textbox";}
        internal static RectangleF Bounds(ProcessGridSpec s,ProcessGridCell c){if(!c.Active)throw new InvalidOperationException("Retired cells have no active layout.");float w=(s.Width-s.Gap*(s.Columns-1))/s.Columns,h=(s.Height-s.Gap*(s.Rows-1))/s.Rows;return new RectangleF(s.Left+c.Column*(w+s.Gap),s.Top+c.Row*(h+s.Gap),w,h);}
        internal static ProcessGridSpec Create(int slideId,int rows,int columns,float width,float height)
        {
            var s=new ProcessGridSpec {SlideId=slideId,Rows=rows,Columns=columns,Width=width-80,Height=Math.Min(64*rows+8*(rows-1),height-140)};for(int r=0;r<rows;r++)for(int c=0;c<columns;c++)s.Cells.Add(new ProcessGridCell {Row=r,Column=c});Validate(s,width,height,false);return s;
        }
        internal static ProcessGridSpec Change(ProcessGridSpec source,string action,int selectedRow,int selectedColumn,bool right=false)
        {
            Validate(source,10000,10000);if(selectedRow<0||selectedRow>=source.Rows||selectedColumn<0||selectedColumn>=source.Columns)throw new InvalidDataException("Select a current process cell.");var s=source.Copy();
            if(action=="add-row"){if(s.Rows==20)throw new InvalidOperationException("A process grid supports at most 20 rows.");s.Height+=(s.Height-s.Gap*(s.Rows-1))/s.Rows+s.Gap;int at=selectedRow+1;foreach(var c in s.Cells.Where(c=>c.Active&&c.Row>=at))c.Row++;for(int c=0;c<s.Columns;c++)s.Cells.Add(new ProcessGridCell {Row=at,Column=c});s.Rows++;}
            else if(action=="add-column"){if(s.Columns==12)throw new InvalidOperationException("A process grid supports at most 12 columns.");int at=selectedColumn+1;foreach(var c in s.Cells.Where(c=>c.Active&&c.Column>=at))c.Column++;for(int r=0;r<s.Rows;r++)s.Cells.Add(new ProcessGridCell {Row=r,Column=at});s.Columns++;}
            else if(action=="delete-row"||action=="delete-column"){bool column=action=="delete-column";if(column?s.Columns==1:s.Rows==1)throw new InvalidOperationException("The last process row or column cannot be removed.");if(!column)s.Height-=(s.Height-s.Gap*(s.Rows-1))/s.Rows+s.Gap;foreach(var c in s.Cells.Where(c=>c.Active)){int position=column?c.Column:c.Row,index=column?selectedColumn:selectedRow;if(position==index){c.Active=false;c.Row=c.Column=-1;}else if(position>index){if(column)c.Column--;else c.Row--;}}if(column)s.Columns--;else s.Rows--;}
            else if(action=="move-column"){int destination=selectedColumn+(right?1:-1);if(destination<0||destination>=s.Columns)throw new InvalidOperationException("No adjacent column exists in that direction.");foreach(var c in s.Cells.Where(c=>c.Active)){if(c.Column==selectedColumn)c.Column=destination;else if(c.Column==destination)c.Column=selectedColumn;}}
            else throw new InvalidDataException("Unknown process grid operation.");Validate(s,10000,10000,false);return s;
        }
        internal static ProcessGridSpec Resize(ProcessGridSpec source,float left,float top,float width,float height,float slideWidth,float slideHeight)
        {
            Validate(source,slideWidth,slideHeight);var s=source.Copy();s.Left=left;s.Top=top;s.Width=width;s.Height=height;Validate(s,slideWidth,slideHeight);return s;
        }
        internal static int Stripe(string axis,int row,int column,int first,int second){if(axis!="row"&&axis!="column"||row<0||column<0||first<0||first>0xFFFFFF||second<0||second>0xFFFFFF)throw new InvalidDataException("Choose row/column striping and two RGB colors.");return ((axis=="row"?row:column)%2)==0?first:second;}
    }
    internal static class ProcessGridCodec
    {
        internal const string Namespace="urn:powerpoint-intelligence:process-grid:v1";private static readonly XNamespace Ns=Namespace;
        internal static string Encode(ProcessGridSpec s){ProcessGridPlan.Validate(s,10000,10000);return new XElement(Ns+"grid",new XAttribute("version","1"),new XAttribute("id",s.Id),new XAttribute("slide",s.SlideId),new XAttribute("rows",s.Rows),new XAttribute("columns",s.Columns),new XAttribute("left",F(s.Left)),new XAttribute("top",F(s.Top)),new XAttribute("width",F(s.Width)),new XAttribute("height",F(s.Height)),new XAttribute("gap",F(s.Gap)),s.Cells.OrderBy(c=>c.ShapeId).Select(c=>new XElement(Ns+"cell",new XAttribute("row",c.Row),new XAttribute("column",c.Column),new XAttribute("shape",c.ShapeId),new XAttribute("active",c.Active?"true":"false")))).ToString(SaveOptions.DisableFormatting);}
        private static string F(float n){return n.ToString("R",CultureInfo.InvariantCulture);}
        private static int I(XElement e,string key){int n;if(!Int32.TryParse((string)e.Attribute(key),NumberStyles.AllowLeadingSign,CultureInfo.InvariantCulture,out n))throw new InvalidDataException("Invalid process grid integer.");return n;}
        private static float N(XElement e,string key){float n;if(!Single.TryParse((string)e.Attribute(key),NumberStyles.Float,CultureInfo.InvariantCulture,out n))throw new InvalidDataException("Invalid process grid length.");return n;}
        private static void Fields(XElement e,string name,string[] fields){if(e.Name!=Ns+name||fields.Any(f=>e.Attribute(f)==null)||e.Attributes().Any(a=>!a.IsNamespaceDeclaration&&(a.Name.NamespaceName!=""||!fields.Contains(a.Name.LocalName)))||e.Nodes().Any(n=>!(n is XElement)))throw new InvalidDataException("Unknown process grid XML fields.");}
        internal static ProcessGridSpec Decode(string xml)
        {
            if(String.IsNullOrEmpty(xml)||xml.Length>150000)throw new InvalidDataException("Missing or excessive process grid metadata.");try{XDocument d;using(var r=XmlReader.Create(new StringReader(xml),new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=150000,IgnoreWhitespace=true}))d=XDocument.Load(r);if(d.Nodes().Any(n=>!(n is XElement))||d.Root==null)throw new InvalidDataException("Malformed process grid document.");var root=d.Root;Fields(root,"grid",new[]{"version","id","slide","rows","columns","left","top","width","height","gap"});if((string)root.Attribute("version")!="1")throw new InvalidDataException("Unknown process grid version.");var s=new ProcessGridSpec {Id=(string)root.Attribute("id"),SlideId=I(root,"slide"),Rows=I(root,"rows"),Columns=I(root,"columns"),Left=N(root,"left"),Top=N(root,"top"),Width=N(root,"width"),Height=N(root,"height"),Gap=N(root,"gap")};foreach(var c in root.Elements()){Fields(c,"cell",new[]{"row","column","shape","active"});string active=(string)c.Attribute("active");if(c.HasElements||active!="true"&&active!="false"||s.Cells.Count>=1000)throw new InvalidDataException("Invalid process grid cells.");s.Cells.Add(new ProcessGridCell {Row=I(c,"row"),Column=I(c,"column"),ShapeId=I(c,"shape"),Active=active=="true"});}ProcessGridPlan.Validate(s,10000,10000);return s;}catch(XmlException ex){throw new InvalidDataException("Malformed process grid XML.",ex);}
        }
    }
    internal static class ProcessFlowTools
    {
        internal const string GridTag="PPI_PROCESS_GRID",RowTag="PPI_PROCESS_ROW",ColumnTag="PPI_PROCESS_COLUMN";
        private sealed class Context {internal PowerPoint.Slide Slide;internal PowerPoint.Shape Selected;internal Office.CustomXMLPart Part;internal ProcessGridSpec State;internal Dictionary<int,PowerPoint.Shape> Shapes;}
        private sealed class Saved {internal PowerPoint.Shape Shape;internal RectangleF Bounds;internal Office.MsoTriState Visible;internal string Row,Column,Xml;internal Office.MsoAutoShapeType Kind;internal float Adjustment;}
        private static PowerPoint.Slide Active(PowerPoint.Application app){if(app==null||app.Presentations.Count==0||app.ActiveWindow.ViewType!=PowerPoint.PpViewType.ppViewNormal)throw new InvalidOperationException("Open a presentation in Normal view.");var slide=app.ActiveWindow.View.Slide as PowerPoint.Slide;if(slide==null)throw new InvalidOperationException("Select a slide.");return slide;}
        private static Office.CustomXMLPart TextPart(PowerPoint.Shape shape){var parts=XmlPartUpdates.AttachedParts(shape,TextShapeCodec.Namespace);if(parts.Length!=1)throw new InvalidDataException("Cell text metadata is missing or ambiguous.");return parts[0];}
        private static Context Read(PowerPoint.Application app)
        {
            var slide=Active(app);var selection=app.ActiveWindow.Selection;if(selection.Type!=PowerPoint.PpSelectionType.ppSelectionShapes&&selection.Type!=PowerPoint.PpSelectionType.ppSelectionText||selection.ShapeRange.Count!=1)throw new InvalidOperationException("Select one current process grid cell.");var selected=selection.ShapeRange[1];string id=selected.Tags[GridTag];Guid identity;if(!Guid.TryParse(id,out identity)||identity==Guid.Empty)throw new InvalidOperationException("Select a process grid cell.");var parts=app.ActivePresentation.CustomXMLParts.SelectByNamespace(ProcessGridCodec.Namespace);if(parts.Count>100)throw new InvalidDataException("Too many process grids to inspect.");var matches=parts.Cast<Office.CustomXMLPart>().Where(p=>String.Equals(ProcessGridCodec.Decode(p.XML).Id,id,StringComparison.OrdinalIgnoreCase)).ToArray();if(matches.Length!=1)throw new InvalidDataException("Process grid metadata is missing or ambiguous.");var state=ProcessGridCodec.Decode(matches[0].XML);ProcessGridPlan.Validate(state,app.ActivePresentation.PageSetup.SlideWidth,app.ActivePresentation.PageSetup.SlideHeight);if(state.SlideId!=slide.SlideID||slide.Shapes.Count>2000)throw new InvalidDataException("Copied or excessive process objects require a fresh grid.");var shapes=slide.Shapes.Cast<PowerPoint.Shape>().Where(s=>String.Equals(s.Tags[GridTag],id,StringComparison.OrdinalIgnoreCase)).ToDictionary(s=>s.Id);if(shapes.Count!=state.Cells.Count)throw new InvalidDataException("Process cells are missing, duplicated or foreign; nothing was changed.");foreach(var c in state.Cells){PowerPoint.Shape shape;if(!shapes.TryGetValue(c.ShapeId,out shape)||shape.Tags[RowTag]!=c.Row.ToString(CultureInfo.InvariantCulture)||shape.Tags[ColumnTag]!=c.Column.ToString(CultureInfo.InvariantCulture))throw new InvalidDataException("Process cell identity or coordinate metadata is damaged.");var spec=TextShapeElements.Read(shape);if(c.Active){if(spec.Kind!=ProcessGridPlan.Kind(c)||shape.Visible!=Office.MsoTriState.msoTrue)throw new InvalidDataException("Process cell kind or visibility differs from its grid.");var expected=ProcessGridPlan.Bounds(state,c);if(Math.Abs(shape.Left-expected.X)>.5||Math.Abs(shape.Top-expected.Y)>.5||Math.Abs(shape.Width-expected.Width)>.5||Math.Abs(shape.Height-expected.Height)>.5)throw new InvalidOperationException("Process cells were moved/resized independently. Undo the manual geometry change before editing the grid.");}else if(shape.Visible!=Office.MsoTriState.msoFalse)throw new InvalidDataException("Retained process history was modified; objects were preserved.");}if(!state.Cells.Any(c=>c.ShapeId==selected.Id&&c.Active))throw new InvalidOperationException("Select a visible active process cell.");return new Context {Slide=slide,Selected=selected,State=state,Part=matches[0],Shapes=shapes};
        }
        private static void Place(PowerPoint.Shape s,RectangleF b){s.Width=b.Width;s.Height=b.Height;s.Left=b.X;s.Top=b.Y;}
        private static void Tags(PowerPoint.Shape s,ProcessGridSpec grid,ProcessGridCell c){s.Tags.Add(GridTag,grid.Id);s.Tags.Add(RowTag,c.Row.ToString(CultureInfo.InvariantCulture));s.Tags.Add(ColumnTag,c.Column.ToString(CultureInfo.InvariantCulture));}
        private static void CopyColor(PowerPoint.ColorFormat source,PowerPoint.ColorFormat target)
        {if(source.Type==Office.MsoColorType.msoColorTypeScheme){target.ObjectThemeColor=source.ObjectThemeColor;target.TintAndShade=source.TintAndShade;}else target.RGB=source.RGB;}
        private static void CopyCellStyle(PowerPoint.Shape source,PowerPoint.Shape target)
        {
            // PickUp/Apply also transfers CustomerData on the affected Office build.
            // Copy supported native formatting explicitly, without attaching source metadata.
            if(source.Fill.Visible==Office.MsoTriState.msoTrue&&source.Fill.Type!=Office.MsoFillType.msoFillSolid)throw new InvalidOperationException("Process-cell style transfer requires a solid or transparent fill; the source cell was preserved.");
            var from=source.TextFrame.TextRange.Font;var to=target.TextFrame.TextRange.Font;
            to.Name=from.Name;to.Size=from.Size;to.Bold=from.Bold;to.Italic=from.Italic;to.Underline=from.Underline;CopyColor(from.Color,to.Color);
            target.TextFrame.TextRange.ParagraphFormat.Alignment=source.TextFrame.TextRange.ParagraphFormat.Alignment;
            target.TextFrame.VerticalAnchor=source.TextFrame.VerticalAnchor;target.TextFrame.MarginLeft=source.TextFrame.MarginLeft;target.TextFrame.MarginRight=source.TextFrame.MarginRight;target.TextFrame.MarginTop=source.TextFrame.MarginTop;target.TextFrame.MarginBottom=source.TextFrame.MarginBottom;
            target.Fill.Solid();CopyColor(source.Fill.ForeColor,target.Fill.ForeColor);target.Fill.Transparency=source.Fill.Transparency;target.Fill.Visible=source.Fill.Visible;
            CopyColor(source.Line.ForeColor,target.Line.ForeColor);target.Line.Weight=source.Line.Weight;target.Line.DashStyle=source.Line.DashStyle;target.Line.Transparency=source.Line.Transparency;target.Line.Visible=source.Line.Visible;
        }
        private static PowerPoint.Shape New(PowerPoint.Application app,ProcessGridSpec grid,ProcessGridCell cell,PowerPoint.Shape style)
        {
            // Native style application precedes attachment of the new cell's canonical XML.
            var s=TextShapeElements.Insert(app,ProcessGridPlan.Kind(cell),style==null?(Action<PowerPoint.Shape>)null:target=>CopyCellStyle(style,target));try{s.TextFrame.TextRange.Text="";Place(s,ProcessGridPlan.Bounds(grid,cell));s.Select();var spec=TextShapeElements.Read(s);spec.AutoFit=false;spec.Locks=(TextEdgeLocks)15;TextShapeElements.Configure(app,spec);Tags(s,grid,cell);cell.ShapeId=s.Id;s.Select();return s;}catch{try{s.Delete();}catch{throw new InvalidOperationException("Process cell creation failed and cleanup was incomplete. Use Undo.");}throw;}
        }
        internal static PowerPoint.Shape Create(PowerPoint.Application app,int rows=2,int columns=3)
        {
            var slide=Active(app);if(app.ActivePresentation.CustomXMLParts.SelectByNamespace(ProcessGridCodec.Namespace).Count>=100)throw new InvalidOperationException("At most 100 process grids are supported per presentation.");PowerPoint.Shape style=null;var selection=app.ActiveWindow.Selection;if((selection.Type==PowerPoint.PpSelectionType.ppSelectionShapes||selection.Type==PowerPoint.PpSelectionType.ppSelectionText)&&selection.ShapeRange.Count==1&&TextShapeElements.IsManaged(selection.ShapeRange[1])){style=selection.ShapeRange[1];TextShapeElements.Read(style);}
            var state=ProcessGridPlan.Create(slide.SlideID,rows,columns,app.ActivePresentation.PageSetup.SlideWidth,app.ActivePresentation.PageSetup.SlideHeight);var added=new List<PowerPoint.Shape>();Office.CustomXMLPart part=null;app.StartNewUndoEntry();try{foreach(var cell in state.Cells)added.Add(New(app,state,cell,style));part=app.ActivePresentation.CustomXMLParts.Add(ProcessGridCodec.Encode(state));if(ProcessGridCodec.Encode(ProcessGridCodec.Decode(part.XML))!=ProcessGridCodec.Encode(state))throw new InvalidDataException("Process grid persistence could not be verified.");added[0].Select();return added[0];}catch(Exception original){bool clean=true;if(part!=null)try{part.Delete();}catch{clean=false;}foreach(var s in added)try{s.Delete();}catch{clean=false;}if(!clean)throw new InvalidOperationException("Process grid creation failed and cleanup was incomplete. Use Undo.",original);throw;}
        }
        private static void Change(PowerPoint.Application app,string action,bool right=false,float left=0,float top=0,float width=0,float height=0)
        {
            var context=Read(app);var old=context.State;var selected=old.Cells.Single(c=>c.ShapeId==context.Selected.Id);var next=action=="geometry"?ProcessGridPlan.Resize(old,left,top,width,height,app.ActivePresentation.PageSetup.SlideWidth,app.ActivePresentation.PageSetup.SlideHeight):ProcessGridPlan.Change(old,action,selected.Row,selected.Column,right);ProcessGridPlan.Validate(next,app.ActivePresentation.PageSetup.SlideWidth,app.ActivePresentation.PageSetup.SlideHeight,false);            // Measure each existing cell at its proposed geometry on an owned temporary native duplicate.
            foreach(var c in next.Cells.Where(c=>c.Active&&c.ShapeId>0))
            {
                TextShapeElements.PreflightGridGeometry(app,context.Shapes[c.ShapeId],ProcessGridPlan.Kind(c),ProcessGridPlan.Bounds(next,c));
            }
            string xml=context.Part.XML;var saved=context.Shapes.Values.Select(s=>new Saved {Shape=s,Bounds=new RectangleF(s.Left,s.Top,s.Width,s.Height),Visible=s.Visible,Row=s.Tags[RowTag],Column=s.Tags[ColumnTag],Xml=TextPart(s).XML,Kind=s.AutoShapeType,Adjustment=s.Adjustments.Count==1?s.Adjustments[1]:0}).ToArray();var added=new List<PowerPoint.Shape>();app.StartNewUndoEntry();
            try{foreach(var c in next.Cells.Where(c=>c.Active&&c.ShapeId==0))added.Add(New(app,next,c,null));var shapes=context.Shapes.Concat(added.Select(s=>new KeyValuePair<int,PowerPoint.Shape>(s.Id,s))).ToDictionary(p=>p.Key,p=>p.Value);foreach(var c in next.Cells){var s=shapes[c.ShapeId];if(c.Active){string kind=ProcessGridPlan.Kind(c);if(TextShapeElements.Read(s).Kind!=kind){s.Select();TextShapeElements.ConvertSelected(app,kind);}Place(s,ProcessGridPlan.Bounds(next,c));s.Visible=Office.MsoTriState.msoTrue;}else s.Visible=Office.MsoTriState.msoFalse;Tags(s,next,c);}context.Part=XmlPartUpdates.Deck(app.ActivePresentation,context.Part,ProcessGridCodec.Encode(next));shapes[next.Cells.First(c=>c.Active).ShapeId].Select();}
            catch(Exception original){bool clean=true;foreach(var s in added)try{s.Delete();}catch{clean=false;}foreach(var s in saved)try{s.Shape.AutoShapeType=s.Kind;if(s.Shape.Adjustments.Count==1)s.Shape.Adjustments[1]=s.Adjustment;Place(s.Shape,s.Bounds);s.Shape.Visible=s.Visible;s.Shape.Tags.Add(RowTag,s.Row);s.Shape.Tags.Add(ColumnTag,s.Column);XmlPartUpdates.Shape(s.Shape,TextPart(s.Shape),s.Xml);}catch{clean=false;}try{var current=app.ActivePresentation.CustomXMLParts.SelectByNamespace(ProcessGridCodec.Namespace).Cast<Office.CustomXMLPart>().Single(p=>ProcessGridCodec.Decode(p.XML).Id==context.State.Id);context.Part=XmlPartUpdates.Deck(app.ActivePresentation,current,xml);}catch{clean=false;}if(!clean)throw new InvalidOperationException("Process grid update failed and restoration was incomplete. Use Undo.",original);throw;}
        }
        internal static void Resize(PowerPoint.Application app,float left,float top,float width,float height){Change(app,"geometry",false,left,top,width,height);}
        internal static void AddRow(PowerPoint.Application app){Change(app,"add-row");}
        internal static void AddColumn(PowerPoint.Application app){Change(app,"add-column");}
        internal static void DeleteSelectedRowOrColumn(PowerPoint.Application app,bool column){Change(app,column?"delete-column":"delete-row");}
        internal static void ReorderColumn(PowerPoint.Application app,bool right){Change(app,"move-column",right);}
        internal static void Refresh(PowerPoint.Application app){Read(app);}
        internal static void Stripe(PowerPoint.Application app,string axis,int firstRgb,int secondRgb)
        {
            var context=Read(app);var cells=context.State.Cells.Where(c=>c.Active).ToArray();var changes=new List<NativeFormattingChange>();foreach(var c in cells){int rgb=ProcessGridPlan.Stripe(axis,c.Row,c.Column,firstRgb,secondRgb);var shape=context.Shapes[c.ShapeId];if(shape.Fill.Type!=Office.MsoFillType.msoFillSolid)throw new InvalidOperationException("Striping preserves non-solid fills. Select a grid with solid fills.");int old=shape.Fill.ForeColor.RGB;var theme=shape.Fill.ForeColor.ObjectThemeColor;float tint=shape.Fill.ForeColor.TintAndShade;var visible=shape.Fill.Visible;changes.Add(new NativeFormattingChange(()=>{shape.Fill.ForeColor.RGB=rgb;shape.Fill.Visible=Office.MsoTriState.msoTrue;},()=>{if(theme==Office.MsoThemeColorIndex.msoNotThemeColor)shape.Fill.ForeColor.RGB=old;else shape.Fill.ForeColor.ObjectThemeColor=theme;shape.Fill.ForeColor.TintAndShade=tint;shape.Fill.Visible=visible;}));}app.StartNewUndoEntry();NativeFormattingTransaction.Run(changes);
        }
    }
}



