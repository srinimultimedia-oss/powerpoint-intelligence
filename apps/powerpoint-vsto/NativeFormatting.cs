using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Office=Microsoft.Office.Core;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
namespace PowerPointIntelligence.AddIn
{
    internal static class NativeFormattingPlan
    {
        internal const int MaxTargets=500,MaxCharacters=2000,MaxDepth=16;
        internal static string ColorKind(string kind){if(kind!=null&&kind.StartsWith("color-",StringComparison.Ordinal))kind=kind.Substring(6);if(kind!="fill"&&kind!="outline"&&kind!="text")throw new InvalidDataException("Choose fill, outline or text color.");return kind;}
        internal static string TextMode(string mode){if(mode!=null&&mode.StartsWith("text-",StringComparison.Ordinal))mode=mode.Substring(5);if(mode=="vertical90"||mode=="90")mode="downward";if(mode=="vertical270"||mode=="270")mode="upward";if(!new[]{"left","center","right","justify","top","middle","bottom","horizontal","downward","upward"}.Contains(mode))throw new InvalidDataException("Unknown text formatting mode.");return mode;}
        internal static void ColorValue(int rgb){if(rgb<0||rgb>0xFFFFFF)throw new InvalidDataException("Choose an RGB color between 0 and 16777215.");}
        internal static int Rgb(Color color){return color.R|(color.G<<8)|(color.B<<16);}
        internal static void Bounds(int count,int characters,int depth){if(count<1||count>MaxTargets||characters<0||characters>MaxCharacters||depth<0||depth>MaxDepth)throw new InvalidDataException("Formatting is limited to 500 native items, 2,000 text characters and 16 nested group levels.");}
        internal static bool HorizontalAlignment(string mode){return mode=="left"||mode=="center"||mode=="right"||mode=="justify";}
        internal static bool VerticalAlignment(string mode){return mode=="top"||mode=="middle"||mode=="bottom";}
        internal static PowerPoint.PpParagraphAlignment Paragraph(string mode){mode=TextMode(mode);if(!HorizontalAlignment(mode))throw new InvalidDataException("Choose horizontal paragraph alignment.");return mode=="left"?PowerPoint.PpParagraphAlignment.ppAlignLeft:mode=="right"?PowerPoint.PpParagraphAlignment.ppAlignRight:mode=="justify"?PowerPoint.PpParagraphAlignment.ppAlignJustify:PowerPoint.PpParagraphAlignment.ppAlignCenter;}
        internal static Office.MsoVerticalAnchor Anchor(string mode){mode=TextMode(mode);if(!VerticalAlignment(mode))throw new InvalidDataException("Choose vertical text alignment.");return mode=="top"?Office.MsoVerticalAnchor.msoAnchorTop:mode=="bottom"?Office.MsoVerticalAnchor.msoAnchorBottom:Office.MsoVerticalAnchor.msoAnchorMiddle;}
        internal static Office.MsoTextOrientation Orientation(string mode){mode=TextMode(mode);if(HorizontalAlignment(mode)||VerticalAlignment(mode))throw new InvalidDataException("Choose horizontal, downward or upward text orientation.");return mode=="downward"?Office.MsoTextOrientation.msoTextOrientationDownward:mode=="upward"?Office.MsoTextOrientation.msoTextOrientationUpward:Office.MsoTextOrientation.msoTextOrientationHorizontal;}
    }
    internal sealed class NativeFormattingChange
    {
        internal readonly Action Apply,Restore;
        internal NativeFormattingChange(Action apply,Action restore){Apply=apply;Restore=restore;}
    }
    internal static class NativeFormattingTransaction
    {
        internal static void Run(IEnumerable<NativeFormattingChange> changes)
        {
            if(changes==null)throw new ArgumentNullException(nameof(changes));var plan=changes.ToArray();if(plan.Length==0||plan.Length>10000||plan.Any(p=>p==null||p.Apply==null||p.Restore==null))throw new InvalidDataException("Invalid formatting mutation plan.");var applied=new List<NativeFormattingChange>();
            try{foreach(var change in plan){applied.Add(change);change.Apply();}}
            catch(Exception original){bool restored=true;for(int i=applied.Count-1;i>=0;i--)try{applied[i].Restore();}catch{restored=false;}if(!restored)throw new InvalidOperationException("Formatting failed and rollback was incomplete. Use PowerPoint Undo and inspect the selected objects.",original);throw;}
        }
    }
    internal static class NativeFormatting
    {
        private sealed class Owner:IWin32Window {internal IntPtr Value;public IntPtr Handle {get{return Value;}}}
        private sealed class Target {internal PowerPoint.Shape Shape;internal PowerPoint.Cell Cell;internal PowerPoint.TextRange Range;internal PowerPoint.TextFrame Frame;}
        private sealed class SelectionPlan {internal readonly List<Target> Items=new List<Target>();internal int Nodes,Characters,Depth;}
        private sealed class ColorState
        {
            private readonly PowerPoint.ColorFormat color;private readonly int rgb;private readonly Office.MsoThemeColorIndex theme;private readonly float tint;
            internal ColorState(PowerPoint.ColorFormat color){this.color=color;rgb=color.RGB;theme=color.ObjectThemeColor;tint=color.TintAndShade;if(rgb<0||theme==Office.MsoThemeColorIndex.msoThemeColorMixed||Single.IsNaN(tint)||Single.IsInfinity(tint)||tint< -1||tint>1)throw new InvalidDataException("Mixed or unsupported color state cannot be safely preserved. Select individual native items.");}
            internal void Restore(){if(theme!=Office.MsoThemeColorIndex.msoNotThemeColor)color.ObjectThemeColor=theme;else color.RGB=rgb;color.TintAndShade=tint;}
        }
        internal static bool CanApply(PowerPoint.Application app)
        {try{if(app==null||app.Presentations.Count==0||app.ActiveWindow.ViewType!=PowerPoint.PpViewType.ppViewNormal)return false;var selection=app.ActiveWindow.Selection;return selection.Type==PowerPoint.PpSelectionType.ppSelectionText||selection.Type==PowerPoint.PpSelectionType.ppSelectionShapes&&selection.ShapeRange.Count>0;}catch(System.Runtime.InteropServices.COMException){return false;}catch(InvalidOperationException){return false;}}
        internal static PowerPoint.ShapeRange SelectedShapes(PowerPoint.Selection selection){return selection.HasChildShapeRange?selection.ChildShapeRange:selection.ShapeRange;}
        private static PowerPoint.Selection Selection(PowerPoint.Application app)
        {if(!CanApply(app))throw new InvalidOperationException("In Normal view select editable native text, shapes, a group or table cells.");return app.ActiveWindow.Selection;}
        private static void Geometry(PowerPoint.Shape shape)
        {foreach(float value in new[]{shape.Left,shape.Top,shape.Width,shape.Height})if(Single.IsNaN(value)||Single.IsInfinity(value)||Math.Abs(value)>100000)throw new InvalidDataException("Selection contains invalid or excessive native shape bounds.");if(shape.Width<0||shape.Height<0)throw new InvalidDataException("Selection has invalid native dimensions.");}
        private static void Add(SelectionPlan plan,PowerPoint.Shape shape,PowerPoint.Cell cell,bool text,int depth)
        {
            if(depth>NativeFormattingPlan.MaxDepth||++plan.Nodes>NativeFormattingPlan.MaxTargets)throw new InvalidDataException("Selection exceeds formatting limits.");plan.Depth=Math.Max(plan.Depth,depth);Geometry(shape);
            if(cell!=null){var frame=shape.TextFrame;var range=frame.TextRange;plan.Characters+=range.Length;if(plan.Characters>NativeFormattingPlan.MaxCharacters)throw new InvalidDataException("Select at most 2,000 table characters.");plan.Items.Add(new Target {Shape=shape,Cell=cell,Frame=frame,Range=range});return;}
            if(shape.Type==Office.MsoShapeType.msoGroup){foreach(PowerPoint.Shape child in shape.GroupItems)Add(plan,child,null,text,depth+1);return;}
            if(shape.HasTable==Office.MsoTriState.msoTrue){var table=shape.Table;int rows=table.Rows.Count,columns=table.Columns.Count;if((long)rows*columns>400)throw new InvalidDataException("Select a table with at most 400 cells.");bool selected=false;for(int r=1;r<=rows;r++)for(int c=1;c<=columns;c++)if(table.Cell(r,c).Selected)selected=true;for(int r=1;r<=rows;r++)for(int c=1;c<=columns;c++){var current=table.Cell(r,c);if(!selected||current.Selected)Add(plan,current.Shape,current,text,depth+1);}return;}
            var target=new Target {Shape=shape,Cell=cell};if(shape.HasTextFrame==Office.MsoTriState.msoTrue){target.Frame=shape.TextFrame;target.Range=target.Frame.TextRange;plan.Characters+=target.Range.Length;if(plan.Characters>NativeFormattingPlan.MaxCharacters)throw new InvalidDataException("Select at most 2,000 text characters for a reversible formatting change.");}if(!text||target.Range!=null)plan.Items.Add(target);
        }
        private static SelectionPlan Collect(PowerPoint.Application app,bool text)
        {
            var selection=Selection(app);var result=new SelectionPlan();if(selection.Type==PowerPoint.PpSelectionType.ppSelectionText){if(!text)throw new InvalidOperationException("For fill/outline select the shape or table cells; text color works on the selected text itself.");var range=selection.TextRange;var frame=range.Parent as PowerPoint.TextFrame;if(frame==null){var shape=range.Parent as PowerPoint.Shape;if(shape!=null)frame=shape.TextFrame;}if(frame==null)throw new InvalidOperationException("The selected text frame is unavailable.");result.Items.Add(new Target {Range=range,Frame=frame});result.Nodes=1;result.Characters=range.Length;}
            else{if(SelectedShapes(selection).Count>NativeFormattingPlan.MaxTargets)throw new InvalidDataException("Select at most 500 objects.");foreach(PowerPoint.Shape shape in SelectedShapes(selection))Add(result,shape,null,text,0);}
            NativeFormattingPlan.Bounds(result.Nodes,result.Characters,result.Depth);if(result.Items.Count==0)throw new InvalidOperationException("No editable native text or formatting target is selected.");return result;
        }
        internal static bool Color(PowerPoint.Application app,string kind,Func<int?> choose=null)
        {
            kind=NativeFormattingPlan.ColorKind(kind);if(choose==null){Selection(app);choose=()=>{using(var dialog=new ColorDialog {FullOpen=true,AnyColor=true})return dialog.ShowDialog(new Owner {Value=new IntPtr(app.ActiveWindow.HWND)})==DialogResult.OK?(int?)NativeFormattingPlan.Rgb(dialog.Color):null;};}var rgb=choose();if(!rgb.HasValue)return false;ApplyColor(app,kind,rgb.Value);return true;
        }
        internal static void ApplyColor(PowerPoint.Application app,string kind,int rgb)
        {
            kind=NativeFormattingPlan.ColorKind(kind);NativeFormattingPlan.ColorValue(rgb);var selection=Collect(app,kind=="text");var changes=new List<NativeFormattingChange>();foreach(var target in selection.Items){
                if(kind=="text"){var range=target.Range;if(range.Length==0){var color=range.Font.Color;var state=new ColorState(color);changes.Add(new NativeFormattingChange(()=>color.RGB=rgb,state.Restore));}else for(int i=1;i<=range.Length;i++){var color=range.Characters(i,1).Font.Color;var state=new ColorState(color);changes.Add(new NativeFormattingChange(()=>color.RGB=rgb,state.Restore));}}
                else if(kind=="fill"){var shape=target.Shape;if(target.Cell==null&&shape.Type==Office.MsoShapeType.msoLine)continue;if(target.Cell==null&&shape.Type!=Office.MsoShapeType.msoAutoShape&&shape.Type!=Office.MsoShapeType.msoFreeform&&shape.Type!=Office.MsoShapeType.msoTextBox)throw new InvalidOperationException("Fill color supports native vector shapes, text boxes and table cells. Select editable parts rather than picture/chart/OLE items.");var fill=shape.Fill;if(fill.Type!=Office.MsoFillType.msoFillSolid)throw new InvalidOperationException("Gradient, texture and patterned fills are preserved. Select solid-fill shapes for this command.");var state=new ColorState(fill.ForeColor);var visible=fill.Visible;changes.Add(new NativeFormattingChange(()=>{fill.ForeColor.RGB=rgb;fill.Visible=Office.MsoTriState.msoTrue;},()=>{state.Restore();fill.Visible=visible;}));}
                else if(target.Cell!=null){foreach(var border in new[]{PowerPoint.PpBorderType.ppBorderTop,PowerPoint.PpBorderType.ppBorderBottom,PowerPoint.PpBorderType.ppBorderLeft,PowerPoint.PpBorderType.ppBorderRight})Outline(target.Cell.Borders[border],rgb,changes);}
                else Outline(target.Shape.Line,rgb,changes);
            }
            if(changes.Count==0)throw new InvalidOperationException("No applicable native formatting targets were selected.");app.StartNewUndoEntry();NativeFormattingTransaction.Run(changes);
        }
        private static void Outline(PowerPoint.LineFormat line,int rgb,List<NativeFormattingChange> changes)
        {var state=new ColorState(line.ForeColor);var visible=line.Visible;changes.Add(new NativeFormattingChange(()=>{line.Visible=Office.MsoTriState.msoTrue;line.ForeColor.RGB=rgb;},()=>{state.Restore();line.Visible=visible;}));}
        internal static void Text(PowerPoint.Application app,string mode)
        {
            mode=NativeFormattingPlan.TextMode(mode);var selection=Collect(app,true);var changes=new List<NativeFormattingChange>();foreach(var target in selection.Items){
                if(NativeFormattingPlan.HorizontalAlignment(mode)){var range=target.Range;int count=range.Paragraphs().Count;for(int i=1;i<=Math.Max(1,count);i++){var paragraph=count==0?range:range.Paragraphs(i,1);var old=paragraph.ParagraphFormat.Alignment;if(old==PowerPoint.PpParagraphAlignment.ppAlignmentMixed)throw new InvalidOperationException("Paragraph alignment is mixed. Select an individual text range.");var next=NativeFormattingPlan.Paragraph(mode);changes.Add(new NativeFormattingChange(()=>paragraph.ParagraphFormat.Alignment=next,()=>paragraph.ParagraphFormat.Alignment=old));}}
                else if(NativeFormattingPlan.VerticalAlignment(mode)){var frame=target.Frame;var old=frame.VerticalAnchor;if(old==Office.MsoVerticalAnchor.msoVerticalAnchorMixed)throw new InvalidOperationException("Text anchor is mixed. Select individual text frames.");var next=NativeFormattingPlan.Anchor(mode);changes.Add(new NativeFormattingChange(()=>frame.VerticalAnchor=next,()=>frame.VerticalAnchor=old));}
                else{var frame=target.Frame;var old=frame.Orientation;if(old==Office.MsoTextOrientation.msoTextOrientationMixed)throw new InvalidOperationException("Text orientation is mixed. Select individual text frames.");var next=NativeFormattingPlan.Orientation(mode);changes.Add(new NativeFormattingChange(()=>frame.Orientation=next,()=>frame.Orientation=old));}
            }
            app.StartNewUndoEntry();NativeFormattingTransaction.Run(changes);
        }
    }
}

