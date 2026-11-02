using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Office=Microsoft.Office.Core;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
namespace PowerPointIntelligence.AddIn
{
    internal sealed class ContextNativeState
    {
        internal ContextSnapshot Snapshot;
        internal FontSelectionState Font;
        internal string Horizontal,Vertical,Direction,Note;
        internal int? TextColor,FillColor,OutlineColor;
        internal Rectangle SelectionScreen,WindowScreen,WorkArea;
        internal IntPtr Owner;
        internal bool ReadOnly;
    }
    internal static class ContextSelectionBridge
    {
        [ThreadStatic]internal static string InspectionStage;
        [StructLayout(LayoutKind.Sequential)]private struct Rect {internal int Left,Top,Right,Bottom;}
        [DllImport("user32.dll")]private static extern bool GetWindowRect(IntPtr hwnd,out Rect rect);
        [DllImport("user32.dll")]private static extern IntPtr GetAncestor(IntPtr hwnd,uint flags);
        [DllImport("user32.dll")]internal static extern IntPtr GetForegroundWindow();
        private static PowerPoint.Shape[] Shapes(PowerPoint.Selection selection){return NativeFormatting.SelectedShapes(selection).Cast<PowerPoint.Shape>().ToArray();}
        private static string ObjectKey(object value){IntPtr id=Marshal.GetIUnknownForObject(value);try{return id.ToInt64().ToString("X",CultureInfo.InvariantCulture);}finally{Marshal.Release(id);}}
        internal static string Key(PowerPoint.Application app)
        {
            InspectionStage="key.context";if(app==null||app.Presentations.Count==0||app.SlideShowWindows.Count>0)return null;var window=app.ActiveWindow;if(window==null||window.ViewType!=PowerPoint.PpViewType.ppViewNormal)return null;var selection=window.Selection;if(selection.Type!=PowerPoint.PpSelectionType.ppSelectionShapes&&selection.Type!=PowerPoint.PpSelectionType.ppSelectionText)return null;var slide=window.View.Slide as PowerPoint.Slide;if(slide==null)return null;var shapes=Shapes(selection);if(shapes.Length<1||shapes.Length>500)return null;
            InspectionStage="key.identity";string key=ObjectKey(app.ActivePresentation)+":"+ContextPanelModel.SelectionKey(window.HWND,slide.SlideID,shapes.Select(s=>s.Id),selection.Type==PowerPoint.PpSelectionType.ppSelectionText?(int?)selection.TextRange.Start:null,selection.Type==PowerPoint.PpSelectionType.ppSelectionText?(int?)selection.TextRange.Length:null);
            if(selection.Type==PowerPoint.PpSelectionType.ppSelectionText&&shapes.Any(s=>s.HasTable==Office.MsoTriState.msoTrue))key+=":cell-text:"+selection.TextRange.BoundLeft.ToString("R",CultureInfo.InvariantCulture)+","+selection.TextRange.BoundTop.ToString("R",CultureInfo.InvariantCulture);
            foreach(var shape in shapes)if(shape.HasTable==Office.MsoTriState.msoTrue){var table=shape.Table;if((long)table.Rows.Count*table.Columns.Count>400){key+=":table-over-limit";continue;}var cells=new List<string>();for(int r=1;r<=table.Rows.Count;r++)for(int c=1;c<=table.Columns.Count;c++)if(table.Cell(r,c).Selected)cells.Add(r+","+c);key+=":cells:"+String.Join(";",cells);}
            return key;
        }
        internal static void Ensure(PowerPoint.Application app,string expected){if(String.IsNullOrEmpty(expected)||!String.Equals(expected,Key(app),StringComparison.Ordinal))throw new InvalidOperationException("The selection changed. Select the element again before applying properties.");if(app.ActivePresentation.ReadOnly==Office.MsoTriState.msoTrue)throw new InvalidOperationException("This presentation is read-only.");}
        private sealed class Targets {internal readonly List<PowerPoint.Shape> Leaves=new List<PowerPoint.Shape>();internal readonly List<PowerPoint.TextRange> Text=new List<PowerPoint.TextRange>();internal readonly List<PowerPoint.TextFrame> Frames=new List<PowerPoint.TextFrame>();internal readonly List<PowerPoint.LineFormat> Outlines=new List<PowerPoint.LineFormat>();internal int Count,Characters;internal bool Fill=true,Outline=true,Font=true;}
        private static void TargetsOf(PowerPoint.Shape shape,Targets targets,int depth,PowerPoint.Cell cell=null)
        {
            if(++targets.Count>500||depth>16)throw new InvalidDataException("Property inspection exceeds500 native parts or16 group levels.");
            if(cell!=null){InspectionStage="capture.table-cell-text";targets.Leaves.Add(shape);var frame=shape.TextFrame;var range=frame.TextRange;targets.Frames.Add(frame);targets.Text.Add(range);targets.Characters+=range.Length;if(targets.Characters>2000)throw new InvalidDataException("Select at most 2,000 table characters.");InspectionStage="capture.table-cell-fill";targets.Fill&=shape.Fill.Type==Office.MsoFillType.msoFillSolid;InspectionStage="capture.table-cell-borders";foreach(var border in new[]{PowerPoint.PpBorderType.ppBorderTop,PowerPoint.PpBorderType.ppBorderBottom,PowerPoint.PpBorderType.ppBorderLeft,PowerPoint.PpBorderType.ppBorderRight})targets.Outlines.Add(cell.Borders[border]);return;}
            if(shape.Type==Office.MsoShapeType.msoGroup){foreach(PowerPoint.Shape child in shape.GroupItems)TargetsOf(child,targets,depth+1);return;}
            if(shape.HasTable==Office.MsoTriState.msoTrue){var t=shape.Table;if((long)t.Rows.Count*t.Columns.Count>400)throw new InvalidDataException("Property inspection supports400 table cells.");bool any=false;for(int r=1;r<=t.Rows.Count;r++)for(int c=1;c<=t.Columns.Count;c++)any|=t.Cell(r,c).Selected;for(int r=1;r<=t.Rows.Count;r++)for(int c=1;c<=t.Columns.Count;c++)if(!any||t.Cell(r,c).Selected)TargetsOf(t.Cell(r,c).Shape,targets,depth+1,t.Cell(r,c));return;}
            targets.Leaves.Add(shape);targets.Outline&=shape.Type!=Office.MsoShapeType.msoEmbeddedOLEObject&&shape.Type!=Office.MsoShapeType.msoLinkedOLEObject;if(shape.HasTextFrame==Office.MsoTriState.msoTrue){var frame=shape.TextFrame;var range=frame.TextRange;targets.Frames.Add(frame);targets.Text.Add(range);targets.Characters+=range.Length;if(targets.Characters>2000)throw new InvalidDataException("Property inspection supports at most2,000 selected text characters.");if(shape.Type!=Office.MsoShapeType.msoAutoShape&&shape.Type!=Office.MsoShapeType.msoFreeform&&shape.Type!=Office.MsoShapeType.msoTextBox&&shape.Type!=Office.MsoShapeType.msoPlaceholder)targets.Font=false;}else if(depth==0)targets.Font=false;
            bool vector=shape.Type==Office.MsoShapeType.msoAutoShape||shape.Type==Office.MsoShapeType.msoFreeform||shape.Type==Office.MsoShapeType.msoTextBox;targets.Fill&=(vector||depth>0&&shape.HasTextFrame==Office.MsoTriState.msoTrue)&&shape.Fill.Type==Office.MsoFillType.msoFillSolid;
            if(shape.Type==Office.MsoShapeType.msoLine)targets.Fill=false;if(targets.Outline)targets.Outlines.Add(shape.Line);
        }
        private static int? Common(IEnumerable<int> values){var mixed=MixedValue<int>.From(values);return mixed.HasValue?(int?)mixed.Value:null;}
        private static string Common(IEnumerable<string> values){var mixed=MixedValue<string>.From(values);return mixed.HasValue?mixed.Value:null;}
        private static PowerPoint.TextFrame Frame(PowerPoint.TextRange range){var frame=range.Parent as PowerPoint.TextFrame;if(frame!=null)return frame;var shape=range.Parent as PowerPoint.Shape;return shape==null?null:shape.TextFrame;}
        private static string Paragraph(PowerPoint.PpParagraphAlignment a){return a==PowerPoint.PpParagraphAlignment.ppAlignLeft?"left":a==PowerPoint.PpParagraphAlignment.ppAlignCenter?"center":a==PowerPoint.PpParagraphAlignment.ppAlignRight?"right":a==PowerPoint.PpParagraphAlignment.ppAlignJustify?"justify":null;}
        private static string Vertical(Office.MsoVerticalAnchor a){return a==Office.MsoVerticalAnchor.msoAnchorTop?"top":a==Office.MsoVerticalAnchor.msoAnchorMiddle?"middle":a==Office.MsoVerticalAnchor.msoAnchorBottom?"bottom":null;}
        private static string Direction(Office.MsoTextOrientation a){return a==Office.MsoTextOrientation.msoTextOrientationHorizontal?"horizontal":a==Office.MsoTextOrientation.msoTextOrientationDownward?"downward":a==Office.MsoTextOrientation.msoTextOrientationUpward?"upward":null;}
        internal static ContextNativeState Capture(PowerPoint.Application app,bool overviewOnly=false)
        {
            string key=Key(app);if(key==null)return null;var window=app.ActiveWindow;var slide=(PowerPoint.Slide)window.View.Slide;var selection=window.Selection;var shapes=Shapes(selection);var descriptors=new List<ContextDescriptor>();var aggregate=new Targets();bool selectedText=selection.Type==PowerPoint.PpSelectionType.ppSelectionText;bool readOnly=app.ActivePresentation.ReadOnly==Office.MsoTriState.msoTrue;
            foreach(var shape in shapes){InspectionStage="capture.targets";var target=new Targets();if(overviewOnly){target.Font=target.Fill=target.Outline=false;}else if(selectedText){target.Text.Add(selection.TextRange);target.Count=1;target.Characters=selection.TextRange.Length;target.Fill=target.Outline=false;}else TargetsOf(shape,target,0);aggregate.Leaves.AddRange(target.Leaves);aggregate.Text.AddRange(target.Text);aggregate.Frames.AddRange(target.Frames);aggregate.Outlines.AddRange(target.Outlines);aggregate.Characters+=target.Characters;InspectionStage="capture.classification";var d=new ContextDescriptor {ShapeId=shape.Id,IsTextBox=shape.Type==Office.MsoShapeType.msoTextBox,IsTable=shape.HasTable==Office.MsoTriState.msoTrue,IsChart=shape.HasChart==Office.MsoTriState.msoTrue,IsConnector=shape.Connector==Office.MsoTriState.msoTrue,IsImage=shape.Type==Office.MsoShapeType.msoPicture||shape.Type==Office.MsoShapeType.msoLinkedPicture,IsShape=true,CanText=target.Text.Count>0,CanFont=target.Font&&target.Text.Count>0,CanFill=target.Fill&&!selectedText,CanOutline=target.Outline&&!selectedText};
                if(overviewOnly){d.CanText=d.CanFont=d.CanFill=d.CanOutline=false;descriptors.Add(d);continue;}
                if(shape.Tags["PPI_INDICATOR"]=="1"){d.ManagedType="indicator";d.CanText=d.CanFont=false;d.CanFill=d.CanOutline=shapes.Length==1;}
                else if(shape.Tags[ProcessFlowTools.GridTag]!="")d.ManagedType="process-grid";
                else if(UniversalConnectors.IsManaged(shape)){d.ManagedType="connector";d.CanOutline=shapes.Length==1;}
                else if(TextShapeElements.IsManaged(shape))d.ManagedType="text";
                else if(shape.Tags[TemplateInserter.TemplateTag]!="") {d.ManagedType="template";d.ActionIds=new[]{"template-palette","template-alignment"};}
                else if(shape.Tags[PresentationAgenda.OwnerTag]!=""&&ChapterAgenda.IsManaged(slide)){d.ManagedType="agenda";d.ActionIds=new[]{"agenda-settings","agenda-refresh"};}
                else if(shape.Tags["PPI_CA_SCHEMA"]=="1"){d.ManagedType="annotation";d.ActionIds=new[]{"annotation-refresh"};}
                if(d.IsChart){d.CanText=false;d.CanFont=shapes.Length==1;d.CanFill=!selectedText;d.ActionIds=new[]{"chart-data","chart-series","chart-axes","chart-style","chart-legend","chart-labels","chart-data-table","data-links"};}
                else if(d.IsTable)d.ActionIds=new[]{"table-data"};
                d.ActionIds=(d.ActionIds??new string[0]).Concat(new[]{"arrange"}).ToArray();
                if(readOnly){d.CanText=d.CanFont=d.CanFill=d.CanOutline=false;d.ActionIds=new string[0];d.ManagedType=null;}
                descriptors.Add(d);
            }
            if(aggregate.Leaves.Count>500||aggregate.Characters>2000)throw new InvalidDataException("Select a smaller group or text range for the property box.");
            InspectionStage="capture.capabilities";var state=new ContextNativeState {Snapshot=ContextPanelModel.Build(descriptors,key),ReadOnly=readOnly};if(selectedText&&!overviewOnly){aggregate.Text.Clear();aggregate.Text.Add(selection.TextRange);}
            InspectionStage="capture.font";if(!overviewOnly&&!readOnly&&state.Snapshot.CanFont)try{state.Font=state.Snapshot.Kind==ContextKind.Chart?ContextChartFormatting.Read(app):ContextFormatting.Read(app);}catch(InvalidOperationException){state.Font=null;}
            InspectionStage="capture.text-style";var ranges=aggregate.Text;var chars=ranges.SelectMany(r=>r.Length==0?new[]{r}:Enumerable.Range(1,r.Length).Select(i=>r.Characters(i,1))).ToArray();
            if(chars.Length>0){state.TextColor=Common(chars.Select(r=>r.Font.Color.RGB));state.Horizontal=Common(chars.Select(r=>Paragraph(r.ParagraphFormat.Alignment)));var frames=selectedText?ranges.Select(Frame).Where(f=>f!=null):aggregate.Frames;state.Vertical=Common(frames.Select(f=>Vertical(f.VerticalAnchor)));state.Direction=Common(frames.Select(f=>Direction(f.Orientation)));}
            InspectionStage="capture.surface-style";if(overviewOnly){state.Note="Select a smaller editable part to change properties.";}else if(state.Snapshot.Kind==ContextKind.Chart){state.TextColor=ContextChartFormatting.ReadTextColor(app);var fill=shapes[0].Chart.ChartArea.Format.Fill;state.FillColor=fill.Type==Office.MsoFillType.msoFillSolid?(int?)fill.ForeColor.RGB:null;state.Note="Chart-wide text; use chart settings for series and axes.";}
            else if(state.Snapshot.Kind==ContextKind.Indicator){var spec=IndicatorElements.Read(shapes[0]);state.FillColor=spec.FilledColor;state.OutlineColor=spec.OutlineColor;state.Note="Indicator state and colors are saved with this element.";}
            else {state.FillColor=state.Snapshot.CanFill?Common(aggregate.Leaves.Where(s=>s.Fill.Visible==Office.MsoTriState.msoTrue).Select(s=>s.Fill.ForeColor.RGB)):null;state.OutlineColor=state.Snapshot.CanOutline?Common(aggregate.Outlines.Where(line=>line.Visible==Office.MsoTriState.msoTrue).Select(line=>line.ForeColor.RGB)):null;}
            InspectionStage="capture.screen-coordinates";float left=shapes.Min(s=>s.Left),top=shapes.Min(s=>s.Top),right=shapes.Max(s=>s.Left+s.Width),bottom=shapes.Max(s=>s.Top+s.Height);int x=(int)window.PointsToScreenPixelsX(left),y=(int)window.PointsToScreenPixelsY(top),x2=(int)window.PointsToScreenPixelsX(right),y2=(int)window.PointsToScreenPixelsY(bottom);state.SelectionScreen=Rectangle.FromLTRB(Math.Min(x,x2),Math.Min(y,y2),Math.Max(x,x2),Math.Max(y,y2));InspectionStage="capture.window-bounds";state.Owner=GetAncestor(new IntPtr(window.HWND),2);Rect rect;if(!GetWindowRect(state.Owner,out rect))throw new InvalidOperationException("The PowerPoint window bounds are unavailable.");state.WindowScreen=Rectangle.FromLTRB(rect.Left,rect.Top,rect.Right,rect.Bottom);state.WorkArea=System.Windows.Forms.Screen.FromRectangle(state.SelectionScreen).WorkingArea;
            if(readOnly)state.Note="Read-only presentation";return state;
        }
        internal static void ApplyColor(PowerPoint.Application app,string expected,string kind,int rgb)
        {
            Ensure(app,expected);NativeFormattingPlan.ColorValue(rgb);var shape=Shapes(app.ActiveWindow.Selection).First();var snapshot=Capture(app).Snapshot;
            if(snapshot.Kind==ContextKind.Indicator){var spec=IndicatorElements.Read(shape);if(kind=="fill")spec.FilledColor=rgb;else if(kind=="outline")spec.OutlineColor=rgb;else throw new InvalidOperationException("Use Indicator settings for its symbol.");IndicatorElements.Configure(app,spec);return;}
            if(snapshot.Kind==ContextKind.Connector&&UniversalConnectors.IsManaged(shape)&&kind=="outline"){UniversalConnectors.Reflow(app);var spec=UniversalConnectors.Read(shape);spec.Color=rgb;UniversalConnectors.Update(app,spec);return;}
            if(snapshot.Kind==ContextKind.Chart&&kind=="text"){ContextChartFormatting.ApplyColor(app,rgb);return;}
            if(snapshot.Kind==ContextKind.Chart&&kind=="fill"){
                var fill=shape.Chart.ChartArea.Format.Fill;if(fill.Type!=Office.MsoFillType.msoFillSolid)throw new InvalidOperationException("Gradient and patterned chart backgrounds are preserved. Use native chart formatting.");var color=fill.ForeColor;int prior=color.RGB;var theme=color.ObjectThemeColor;float tint=color.TintAndShade;var visible=fill.Visible;app.StartNewUndoEntry();NativeFormattingTransaction.Run(new[]{new NativeFormattingChange(()=>{color.RGB=rgb;fill.Visible=Office.MsoTriState.msoTrue;},()=>{if(theme==Office.MsoThemeColorIndex.msoNotThemeColor)color.RGB=prior;else color.ObjectThemeColor=theme;color.TintAndShade=tint;fill.Visible=visible;})});return;
            }
            NativeFormatting.ApplyColor(app,kind,rgb);
        }
    }
}
