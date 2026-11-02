using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using Office=Microsoft.Office.Core;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
namespace PowerPointIntelligence.AddIn
{
    internal sealed class QualityInspectionOptions
    {
        internal float MinimumFontPoints=10;
        internal int MaxSlides=500,MaxShapes=20000,MaxGroupDepth=8,MaxFontCharacters=256,MaxIssues=2000;
        internal void Validate(){if(float.IsNaN(MinimumFontPoints)||MinimumFontPoints<1||MinimumFontPoints>100||MaxSlides<1||MaxSlides>2000||MaxShapes<1||MaxShapes>100000||MaxGroupDepth<0||MaxGroupDepth>16||MaxFontCharacters<1||MaxFontCharacters>4096||MaxIssues<1||MaxIssues>10000)throw new ArgumentOutOfRangeException("options");}
    }
    internal sealed class QualityShapeSnapshot
    {
        internal int Id,LocateShapeId;
        internal float Left,Top,Width,Height,Rotation;
        internal bool TopLevel=true,EditableText,HasText,Picture,HasAlternativeText,FontScanTruncated;
        internal readonly List<float> FontSizes=new List<float>();
    }
    internal sealed class QualitySlideSnapshot
    {
        internal int Id,Index;
        internal float Width,Height;
        internal bool Hidden;
        internal readonly List<QualityShapeSnapshot> Shapes=new List<QualityShapeSnapshot>();
    }
    internal sealed class QualityIssue
    {
        internal int SlideId,SlideIndex,ShapeId,LocateShapeId;
        internal string Rule,Message;
    }
    internal sealed class QualityInspectionReport
    {
        internal readonly List<QualityIssue> Issues=new List<QualityIssue>();
        internal readonly List<string> Limitations=new List<string>();
        internal int SlidesInspected,ShapesInspected;
        internal bool Truncated;
    }
    internal static class PresentationQuality
    {
        internal static QualityInspectionReport Evaluate(IEnumerable<QualitySlideSnapshot> slides,QualityInspectionOptions options=null)
        {
            options=options??new QualityInspectionOptions();options.Validate();
            if(slides==null)throw new ArgumentNullException("slides");
            var result=new QualityInspectionReport();
            result.Limitations.Add("Review heuristics only; not an accessibility or branding certification. No content is saved. Masters, notes, chart-internal text, SmartArt and table cells are not inspected.");
            result.Limitations.Add("Off-slide checks use unrotated top-level bounding rectangles only; rotated objects and group-child geometry are skipped. Effects and 3D extents are not measured.");
            foreach(var slide in slides)
            {
                if(result.SlidesInspected>=options.MaxSlides){result.Truncated=true;break;}
                if(slide==null)throw new ArgumentException("Null slide snapshot.");
                if(!Finite(slide.Width)||!Finite(slide.Height)||slide.Width<=0||slide.Height<=0)throw new ArgumentException("Invalid slide dimensions.");
                result.SlidesInspected++;
                if(slide.Hidden)Add(result,options,slide,null,"hidden-slide","Slide is hidden during the slide show; review whether intentional.");
                foreach(var shape in slide.Shapes)
                {
                    if(result.ShapesInspected>=options.MaxShapes){result.Truncated=true;break;}
                    if(shape==null)throw new ArgumentException("Null shape snapshot.");
                    result.ShapesInspected++;
                    if(shape.TopLevel&&Finite(shape.Rotation)&&Math.Abs(shape.Rotation%360)<0.01f&&Finite(shape.Left)&&Finite(shape.Top)&&Finite(shape.Width)&&Finite(shape.Height)&&shape.Width>=0&&shape.Height>=0&&(shape.Left<-.5f||shape.Top<-.5f||shape.Left+shape.Width>slide.Width+.5f||shape.Top+shape.Height>slide.Height+.5f))Add(result,options,slide,shape,"off-slide","Unrotated object bounds extend outside the slide (0.5 pt tolerance).");
                    if(shape.EditableText&&!shape.HasText)Add(result,options,slide,shape,"empty-text","Editable text box or placeholder has no text; review whether intentional.");
                    if(shape.HasText&&shape.FontSizes.Any(n=>Finite(n)&&n>0&&n<options.MinimumFontPoints))Add(result,options,slide,shape,"small-text","Text below the selected "+options.MinimumFontPoints.ToString("0.##",CultureInfo.InvariantCulture)+" pt heuristic was found.");
                    if(shape.Picture&&!shape.HasAlternativeText)Add(result,options,slide,shape,"image-alt-text","Picture has no alternative text; review whether decorative.");
                    if(shape.FontScanTruncated){result.Truncated=true;Add(result,options,slide,shape,"font-scan-limit","Mixed-font text exceeds the character inspection limit; remaining characters were not inspected.");}
                }
                if(result.ShapesInspected>=options.MaxShapes){result.Truncated=true;break;}
            }
            if(result.Truncated)result.Limitations.Add("Inspection reached a configured limit; absence of a finding does not mean all content was checked.");
            return result;
        }
        private static bool Finite(float value){return !float.IsNaN(value)&&!float.IsInfinity(value);}
        private static void Add(QualityInspectionReport report,QualityInspectionOptions options,QualitySlideSnapshot slide,QualityShapeSnapshot shape,string rule,string message)
        {
            if(report.Issues.Count>=options.MaxIssues){report.Truncated=true;return;}
            report.Issues.Add(new QualityIssue{SlideId=slide.Id,SlideIndex=slide.Index,ShapeId=shape==null?0:shape.Id,LocateShapeId=shape==null?0:shape.LocateShapeId,Rule=rule,Message=message});
        }
        internal static QualityInspectionReport Inspect(PowerPoint.Presentation presentation,QualityInspectionOptions options=null)
        {
            if(presentation==null)throw new ArgumentNullException("presentation");
            options=options??new QualityInspectionOptions();options.Validate();
            var slides=new List<QualitySlideSnapshot>();var warnings=new List<string>();int shapes=0;bool limited=false;
            int slideCount=presentation.Slides.Count;
            for(int i=1;i<=Math.Min(slideCount,options.MaxSlides);i++)
            {
                var slide=presentation.Slides[i];
                var dto=new QualitySlideSnapshot{Id=slide.SlideID,Index=i,Width=presentation.PageSetup.SlideWidth,Height=presentation.PageSetup.SlideHeight,Hidden=slide.SlideShowTransition.Hidden==Office.MsoTriState.msoTrue};
                slides.Add(dto);
                for(int j=1;j<=slide.Shapes.Count;j++){if(shapes>=options.MaxShapes){limited=true;break;}Capture(slide.Shapes[j],dto,0,0,options,ref shapes,ref limited,warnings);}
                if(limited&&shapes>=options.MaxShapes)break;
            }
            var report=Evaluate(slides,options);report.Truncated|=limited||slideCount>options.MaxSlides;
            foreach(var warning in warnings.Distinct())report.Limitations.Add(warning);
            if(report.Truncated)report.Limitations.Add("Native scan was bounded; some slides, nested shapes or font characters may be uninspected.");
            return report;
        }
        private static void Capture(PowerPoint.Shape shape,QualitySlideSnapshot slide,int depth,int outerId,QualityInspectionOptions options,ref int count,ref bool limited,List<string> warnings)
        {
            if(count>=options.MaxShapes){limited=true;return;}count++;
            try
            {
                int id=shape.Id;var type=shape.Type;
                var dto=new QualityShapeSnapshot{Id=id,LocateShapeId=outerId==0?id:outerId,TopLevel=depth==0,Left=shape.Left,Top=shape.Top,Width=shape.Width,Height=shape.Height,Rotation=shape.Rotation,Picture=type==Office.MsoShapeType.msoPicture||type==Office.MsoShapeType.msoLinkedPicture};
                slide.Shapes.Add(dto);
                if(dto.Picture)dto.HasAlternativeText=!String.IsNullOrWhiteSpace(shape.AlternativeText);
                if(shape.HasTextFrame==Office.MsoTriState.msoTrue)
                {
                    dto.EditableText=type==Office.MsoShapeType.msoTextBox||type==Office.MsoShapeType.msoPlaceholder;
                    dto.HasText=shape.TextFrame.HasText==Office.MsoTriState.msoTrue;
                    if(dto.HasText)
                    {
                        var range=shape.TextFrame.TextRange;float size=range.Font.Size;
                        if(size>0)dto.FontSizes.Add(size);
                        else {int length=range.Length;for(int k=1;k<=Math.Min(length,options.MaxFontCharacters);k++)dto.FontSizes.Add(range.Characters(k,1).Font.Size);dto.FontScanTruncated=length>options.MaxFontCharacters;}
                    }
                }
                if(type==Office.MsoShapeType.msoGroup)
                {
                    if(depth>=options.MaxGroupDepth){limited=true;warnings.Add("Group recursion limit reached; nested objects were not inspected.");return;}
                    for(int j=1;j<=shape.GroupItems.Count;j++){if(count>=options.MaxShapes){limited=true;break;}Capture(shape.GroupItems[j],slide,depth+1,dto.LocateShapeId,options,ref count,ref limited,warnings);}
                }
            }
            catch(System.Runtime.InteropServices.COMException ex){limited=true;warnings.Add("A native object could not be fully inspected (HRESULT 0x"+ex.HResult.ToString("X8",CultureInfo.InvariantCulture)+"). No content was modified.");}
        }
        internal static void ShowInspection(PowerPoint.Application application)
        {
            if(application==null||application.Presentations.Count==0){MessageBox.Show("Open a presentation first.","Presentation quality");return;}
            var presentation=application.ActivePresentation;
            var form=new Form{Text="Presentation quality — read-only review",Width=1050,Height=650,StartPosition=FormStartPosition.CenterScreen};
            var panel=new FlowLayoutPanel{Dock=DockStyle.Top,Height=38};
            panel.Controls.Add(new Label{Text="Small text threshold (pt):",AutoSize=true,Padding=new Padding(0,7,0,0)});
            var threshold=new NumericUpDown{Minimum=1,Maximum=100,Value=10,Width=65};panel.Controls.Add(threshold);
            var refresh=new Button{Text="Inspect",AutoSize=true};var locate=new Button{Text="Locate selected",AutoSize=true};panel.Controls.Add(refresh);panel.Controls.Add(locate);
            var detail=new TextBox{Dock=DockStyle.Bottom,Height=110,Multiline=true,ReadOnly=true,ScrollBars=ScrollBars.Vertical};
            var grid=new DataGridView{Dock=DockStyle.Fill,ReadOnly=true,AllowUserToAddRows=false,AllowUserToDeleteRows=false,AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.Fill,SelectionMode=DataGridViewSelectionMode.FullRowSelect,MultiSelect=false};
            grid.Columns.Add("slide","Slide");grid.Columns[0].FillWeight=15;grid.Columns.Add("shape","Object ID");grid.Columns[1].FillWeight=20;grid.Columns.Add("rule","Review");grid.Columns.Add("message","Finding");grid.Columns[3].FillWeight=180;
            form.Controls.Add(grid);form.Controls.Add(detail);form.Controls.Add(panel);
            Action inspect=()=>{try{var report=Inspect(presentation,new QualityInspectionOptions{MinimumFontPoints=(float)threshold.Value});grid.Rows.Clear();foreach(var issue in report.Issues){int row=grid.Rows.Add(issue.SlideIndex,issue.ShapeId==0?"—":issue.ShapeId.ToString(CultureInfo.InvariantCulture),issue.Rule,issue.Message);grid.Rows[row].Tag=issue;}detail.Text=report.SlidesInspected+" slides; "+report.ShapesInspected+" objects; "+report.Issues.Count+" findings. "+(report.Truncated?"INCOMPLETE SCAN. ":"")+String.Join(Environment.NewLine,report.Limitations);}catch(Exception ex){detail.Text="Inspection unavailable ("+ex.GetType().Name+"). The presentation may have closed; reopen this review.";}};
            refresh.Click+=(sender,args)=>inspect();
            locate.Click+=(sender,args)=>{if(grid.SelectedRows.Count==0)return;var issue=grid.SelectedRows[0].Tag as QualityIssue;if(issue==null)return;try{presentation.Windows[1].Activate();var slide=presentation.Slides.FindBySlideID(issue.SlideId);slide.Select();if(issue.LocateShapeId!=0){bool found=false;for(int i=1;i<=slide.Shapes.Count;i++){var shape=slide.Shapes[i];if(shape.Id==issue.LocateShapeId){shape.Select(Office.MsoTriState.msoTrue);found=true;break;}}if(!found)detail.Text="Object changed or was deleted; inspect again.";} }catch(Exception ex){detail.Text="Locate unavailable ("+ex.GetType().Name+"). Inspect again; nested findings select their containing group.";}};
            form.Shown+=(sender,args)=>inspect();form.Show();
        }
    }
}
