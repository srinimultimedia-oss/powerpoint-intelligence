using PowerPointIntelligence.Security;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Windows.Forms;
using PowerPointIntelligence.AddIn;
using PowerPointIntelligence.Host;
using Office=Microsoft.Office.Core;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;

// Exercises production command implementations with explicit fixture inputs.
// It does not replace verification of native file pickers or the installed UI.
internal sealed class DemoScenario
{
    private readonly PowerPoint.Application app;
    private readonly string id,root;
    internal PowerPoint.Presentation Deck,CreatedTour;
    private PowerPoint.Slide slide;
    private PowerPoint.Shape chartShape,anchor;
    private readonly List<PowerPoint.Shape> shapes=new List<PowerPoint.Shape>();
    private readonly List<string> messages=new List<string>();
    private readonly DemoInputs inputs;
    private string output,sourceWorkbook,managedId;
    private bool priorLegend,priorDataTable,toolbarClicked;
    private int beforeCount;
    private static readonly HashSet<string> Native=new HashSet<string>(new[]{"column","bar","stacked-column","percent-column","stacked-bar","percent-bar","line","profile","area","pie","doughnut","scatter","bubble","combo"});
    private void Check(bool value,string failure){if(!value)throw new Exception(failure);}
    private bool Near(float a,float b){return Math.Abs(a-b)<0.05f;}
    internal DemoScenario(PowerPoint.Application app,string id,string root)
    {
        this.app=app;this.id=id;this.root=PathPolicy.TestPath(root);
        inputs=new DemoInputs {
            Prompt=(title,initial)=>PromptValue(title),
            OpenFile=ChooseFile,SaveFile=(filter,name)=>output=Path.Combine(root,id+".export"+Path.GetExtension(name)),
            Folder=()=>Assets(),ChooseAsset=files=>Array.FindIndex(files,f=>Path.GetFileName(f)=="sample.png"),
            Confirm=(message,title)=>true,Notify=(message,title)=>messages.Add(title+"\n"+message),
            EditTable=data=>{data[0,0]="Edited fixture";return data;},
            ShowToolbar=form=>{try{form.Show();Application.DoEvents();var button=form.Controls[0].Controls.OfType<Button>().Single(b=>b.Text=="left");button.PerformClick();toolbarClicked=true;}finally{form.Close();form.Dispose();}},
            PresentationCreated=created=>CreatedTour=created
        };
    }
    private string PromptValue(string title)
    {
        switch(title){case "Reference value":return "20";case "Series, from category, to category (1-based)":return "1,1,2";case "Symbol or text":return "✓";case "One line per selected object":return "Alpha\nBeta\nGamma";
        case "Font family":return "Arial";case "Font size (points)":return "24";case "Office language ID":return "1033";
        case "Native chart style (1–48)":return "4";case "Chart title":return "Fixture title";case "Value-axis title":return "Fixture axis";
        case "Start, end, years (comma separated)":return "100,121,2";case "Decimal places (0–10)":return "0";
        case "Start, end, years, destination cell":return "100,121,2,E2";default:throw new Exception("Unexpected prompt: "+title);}
    }
    private string Assets()
    {
        var directory=Path.Combine(root,"assets");Directory.CreateDirectory(directory);var path=Path.Combine(directory,"sample.png");
        if(!File.Exists(path))using(var image=new Bitmap(30,20)){using(var graphics=Graphics.FromImage(image))graphics.Clear(Color.Navy);image.Save(path,System.Drawing.Imaging.ImageFormat.Png);}
        return directory;
    }
    private string ChooseFile(string filter)
    {
        if(filter.StartsWith("Images"))return Path.Combine(Assets(),"sample.png");
        if(filter.StartsWith("NexWave demo style"))return Path.GetFullPath("docs/demo/sample-style.json");
        if(filter.StartsWith("S0 chart JSON"))return Path.GetFullPath("docs/demo/sample-chart.s0.json");
        if(filter.StartsWith("Excel workbook"))return Workbook();
        if(filter.StartsWith("PowerPoint template"))return SourceDeck(true);
        if(filter.StartsWith("PowerPoint|"))return SourceDeck(false);
        throw new Exception("Unexpected file input: "+filter);
    }
    private string SourceDeck(bool template)
    {
        var path=Path.Combine(root,template?"source.potx":"source.pptx");
        if(!File.Exists(path))
        {
            var source=app.Presentations.Add(Office.MsoTriState.msoTrue);
            try{source.Slides.Add(1,PowerPoint.PpSlideLayout.ppLayoutBlank);source.SlideMaster.Background.Fill.Solid();source.SlideMaster.Background.Fill.ForeColor.RGB=0x123456;
                source.Slides[1].Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,20,20,200,50).TextFrame.TextRange.Text="Imported fixture";
                source.SaveAs(path,template?PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLTemplate:PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);}
            finally{source.Saved=Office.MsoTriState.msoTrue;source.Close();Deck.Windows[1].Activate();slide.Select();}
        }
        return path;
    }
    private string Workbook()
    {
        if(sourceWorkbook==null){sourceWorkbook=Path.Combine(root,id+"-source.xlsx");WriteWorkbook(10);}
        return sourceWorkbook;
    }
    private void WriteWorkbook(double value)
    {
        // New Excel application instance; never attaches to or quits a user's instance.
        dynamic excel=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));dynamic book=null;
        try{excel.Visible=true;book=excel.Workbooks.Add();dynamic sheet=book.Worksheets[1];sheet.Cells[1,1].Value2="Category";sheet.Cells[1,2].Value2="Actual";
            sheet.Cells[2,1].Value2="A";sheet.Cells[2,2].Value2=value;sheet.Cells[3,1].Value2="B";sheet.Cells[3,2].Value2=20.0;excel.DisplayAlerts=false;book.SaveAs(sourceWorkbook,51);}
        finally{try{if(book!=null)book.Close(false);}finally{excel.Quit();Deck.Windows[1].Activate();slide.Select();}}
    }
    internal void Prepare()
    {
        Deck=app.Presentations.Add(Office.MsoTriState.msoTrue);Deck.Windows[1].ViewType=PowerPoint.PpViewType.ppViewNormal;
        slide=Deck.Slides.Add(1,PowerPoint.PpSlideLayout.ppLayoutBlank);slide.Select();
        string linkedSource=Path.Combine(root,id+"-linked.csv");
        File.WriteAllText(linkedSource,id=="waterfall"?"Category,Change\nA,137.5\nB,-8\nC,20":"Category,Actual,Plan\n1,137.5,15\n2,22,18\n3,30,35");
        DataSources.Save(Deck,DataSources.Create("csv",Path.GetFullPath(linkedSource)));

        for(int i=0;i<3;i++){var shape=slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle,40+i*180,80+i*40,100+i*10,60+i*10);shape.LockAspectRatio=Office.MsoTriState.msoFalse;shape.TextFrame.TextRange.Text="1.25";shapes.Add(shape);}
        if(new[]{"legend","labels","data-table","plot-area","edit-data","sheet-style","axes","chart-title","chart-style","round-data","excel-cagr","scan","slide-workbook","value-line","difference"}.Contains(id))
        {
            DemoTools.Catalog.Single(t=>t.Id=="column").Execute(new DemoTools(app));chartShape=slide.Shapes[slide.Shapes.Count];
            priorLegend=chartShape.Chart.HasLegend;priorDataTable=chartShape.Chart.HasDataTable;
            if(id=="round-data"){chartShape.Chart.ChartData.Activate();dynamic book=chartShape.Chart.ChartData.Workbook;try{book.Worksheets[1].Cells[2,2].Value2=10.456;}finally{PowerPointHost.RetryRejectedCall(()=>book.Close(true));}}
            if(id=="scan")slide.Shapes.AddPicture(Path.Combine(Assets(),"sample.png"),Office.MsoTriState.msoFalse,Office.MsoTriState.msoTrue,10,10,-1,-1);
            chartShape.Select(Office.MsoTriState.msoTrue);
        }
        else if(id=="export-json")
        {var sample=MetadataCodec.Sample();managedId=sample.ChartId;new PowerPointHost(app).InsertEditableChart(sample);chartShape=slide.Shapes[slide.Shapes.Count];}
        else if(id=="table-data")
        {DemoTools.Catalog.Single(t=>t.Id=="table").Execute(new DemoTools(app));slide.Shapes[slide.Shapes.Count].Select(Office.MsoTriState.msoTrue);}
        else if(id=="refresh-excel"||id=="manage-links")
        {DemoLinks.Insert(app,Workbook());chartShape=slide.Shapes[slide.Shapes.Count];if(id=="refresh-excel")WriteWorkbook(42);chartShape.Select(Office.MsoTriState.msoTrue);}
        else if(id=="ungroup")slide.Shapes.Range(shapes.Select(s=>(object)s.Name).ToArray()).Group().Select(Office.MsoTriState.msoTrue);
        else if(id=="save-slides"){Deck.Windows[1].ViewType=PowerPoint.PpViewType.ppViewSlideSorter;slide.Select();}
        else
        {
            if(id=="front"||id=="back"){anchor=slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle,5,5,20,20);if(id=="back")anchor.ZOrder(Office.MsoZOrderCmd.msoSendToBack);}
            slide.Shapes.Range(shapes.Take(id=="swap"?2:3).Select(s=>(object)s.Name).ToArray()).Select(Office.MsoTriState.msoTrue);
        }
        if(id=="clean-comments")slide.Comments.Add(10,10,"Fixture","F","Remove fixture comment");
        beforeCount=slide.Shapes.Count;
    }
    internal void AssertReport(PowerPoint.Presentation presentation)
    {
        Check(presentation.Slides.Count==8,"Linked report slide count changed");var data=LinkedChartData.Load(DataSources.Read(presentation).Single());
        for(int i=2;i<=8;i++){var charts=presentation.Slides[i].Shapes.Cast<PowerPoint.Shape>().Where(s=>s.HasChart==Office.MsoTriState.msoTrue).ToArray();Check(charts.Length==1,"Report slide must retain one editable chart");var chart=charts[0].Chart;var series=(PowerPoint.SeriesCollection)chart.SeriesCollection();int count=chart.ChartType==Office.XlChartType.xlPie?1:data.Headers.Length-1;Check(series.Count==count,"Report series count changed");for(int si=1;si<=count;si++){var values=PowerPointHost.NativeSeriesValues(series,si);Check(values.Length==data.Rows.Length,"Report category count changed");for(int j=0;j<values.Length;j++)Check(values[j]!=null&&Math.Abs(Convert.ToDouble(values[j],CultureInfo.InvariantCulture)-data.Number(j,si).Value)<1e-9,"Report linked value changed");}}
    }
    internal void ExecuteAndAssert(DemoTool tool)
    {
        new DemoTools(app,inputs).Execute(tool);
        if(id=="demo-tour"){Check(app.ActivePresentation.Name==Deck.Name&&Deck.Slides.Count==8,"Linked report must append seven slides to current presentation");AssertReport(Deck);return;}
        if(Native.Contains(id)){AssertChart();return;}
        switch(id)
        {
            case "waterfall":case "mekko":case "mekko-units":case "gantt":Check(slide.Shapes.Count==beforeCount+1&&slide.Shapes[slide.Shapes.Count].Type==Office.MsoShapeType.msoGroup,"Editable grouped diagram missing");break;
            case "text":Check(LastText()==""&&TextShapeElements.Read(slide.Shapes[slide.Shapes.Count]).Kind=="textbox","Editable blank text element insertion failed");break;
            case "rounded":Check(slide.Shapes[slide.Shapes.Count].AutoShapeType==Office.MsoAutoShapeType.msoShapeRoundedRectangle&&LastText()==""&&TextShapeElements.Read(slide.Shapes[slide.Shapes.Count]).Kind=="roundedRectangle","Rounded rectangle failed");break;
            case "process":Check(slide.Shapes.Count>beforeCount&&!String.IsNullOrEmpty(slide.Shapes[slide.Shapes.Count].Tags[ProcessFlowTools.GridTag]),"Managed native process cells missing");break;
            case "checkbox":case "harvey":Check(IndicatorElements.Read(slide.Shapes[slide.Shapes.Count]).Kind==id&&slide.Shapes[slide.Shapes.Count].Type==Office.MsoShapeType.msoGroup,"Editable vector indicator failed");break;case "symbol":Check(LastText()=="✓","Symbol failed");break;
            case "table":Check(slide.Shapes[slide.Shapes.Count].Table.Rows.Count==4&&slide.Shapes[slide.Shapes.Count].Table.Columns.Count==3&&slide.Shapes[slide.Shapes.Count].Table.Cell(2,2).Shape.TextFrame.TextRange.Text=="137.5","Linked table dimensions/data missing");break;
            case "agenda":Check(Deck.Slides.Count==2&&ChapterAgenda.IsManaged(Deck.Slides[2])&&ChapterAgenda.Read(app).Chapters.Count==1,"Managed chapter slide missing");break;
            case "agenda-layout":Check(Deck.Slides.Count==2&&Deck.Slides[2].Tags[ChapterAgenda.Kind]=="CONTENTS"&&ChapterAgenda.Read(app).Chapters.Count==0,"Empty managed table of contents missing");break;
            case "left":Check(shapes.All(s=>Near(s.Left,40)),"Left alignment");break;case "center":Check(shapes.All(s=>Near(s.Left+s.Width/2,90)),"Center alignment");break;case "right":Check(shapes.All(s=>Near(s.Left+s.Width,140)),"Right alignment");break;
            case "top":Check(shapes.All(s=>Near(s.Top,80)),"Top alignment");break;case "middle":Check(shapes.All(s=>Near(s.Top+s.Height/2,110)),"Middle alignment");break;case "bottom":Check(shapes.All(s=>Near(s.Top+s.Height,140)),"Bottom alignment");break;
            case "width":Check(shapes.All(s=>Near(s.Width,100)),"Width match");break;case "height":Check(shapes.All(s=>Near(s.Height,60)),"Height match");break;case "size":Check(shapes.All(s=>Near(s.Width,100)&&Near(s.Height,60)),"Size match");break;
            case "horizontal":Check(Near(shapes[1].Left,215),"Horizontal spacing");break;case "vertical":Check(Near(shapes[1].Top,115),"Vertical spacing");break;
            case "swap":Check(Near(shapes[0].Left,220)&&Near(shapes[0].Top,120)&&Near(shapes[1].Left,40)&&Near(shapes[1].Top,80),"Position swap");break;
            case "scale-up":case "scale-down":float scale=id=="scale-up"?1.1f:1/1.1f;Check(shapes.Select((s,i)=>Near(s.Width,(100+i*10)*scale)&&Near(s.Height,(60+i*10)*scale)).All(v=>v),"Scale");break;
            case "rotate":Check(shapes.All(s=>Near(s.Rotation,90)),"Rotation");break;case "flip-horizontal":Check(shapes.All(s=>s.HorizontalFlip==Office.MsoTriState.msoTrue),"Flip H");break;case "flip-vertical":Check(shapes.All(s=>s.VerticalFlip==Office.MsoTriState.msoTrue),"Flip V");break;
            case "duplicate":Check(slide.Shapes.Count==6,"Duplicate shapes");break;case "group":Check(slide.Shapes.Count==1&&slide.Shapes[1].GroupItems.Count==3,"Group");break;case "ungroup":Check(slide.Shapes.Count==3,"Ungroup");break;
            case "front":Check(shapes.All(s=>s.ZOrderPosition>anchor.ZOrderPosition),"Bring front");break;case "back":Check(shapes.All(s=>s.ZOrderPosition<anchor.ZOrderPosition),"Send back");break;
            case "auto-layout":Check(shapes.All(s=>s.Left>=40&&s.Top>=60&&s.Left+s.Width<=Deck.PageSetup.SlideWidth-40&&s.Top+s.Height<=Deck.PageSetup.SlideHeight-40),"Grid outside slide");for(int i=0;i<shapes.Count;i++)for(int j=i+1;j<shapes.Count;j++)Check(shapes[i].Left+shapes[i].Width<=shapes[j].Left||shapes[j].Left+shapes[j].Width<=shapes[i].Left||shapes[i].Top+shapes[i].Height<=shapes[j].Top||shapes[j].Top+shapes[j].Height<=shapes[i].Top,"Grid overlap");break;
            case "paste-text":Check(shapes.Select(s=>s.TextFrame.TextRange.Text).SequenceEqual(new[]{"Alpha","Beta","Gamma"}),"Text lines");break;
            case "font":Check(shapes.All(s=>s.TextFrame.TextRange.Font.Name=="Arial"),"Font family");break;case "font-size":Check(shapes.All(s=>Near(s.TextFrame.TextRange.Font.Size,24)),"Font size");break;
            case "decimal":Check(shapes.All(s=>s.TextFrame.TextRange.Text=="1,25"),"Decimal replacement");break;case "language":Check(shapes.All(s=>(int)s.TextFrame.TextRange.LanguageID==1033),"Proofing language");break;
            case "legend":Check(chartShape.Chart.HasLegend!=priorLegend,"Legend toggle");break;case "labels":Check(Convert.ToBoolean(((PowerPoint.SeriesCollection)chartShape.Chart.SeriesCollection()).Item(1).HasDataLabels),"Data labels");break;
            case "data-table":Check(chartShape.Chart.HasDataTable!=priorDataTable,"Data table toggle");break;case "chart-style":Check(Convert.ToInt32(chartShape.Chart.ChartStyle)==4,"Chart style");break;
            case "chart-title":Check(chartShape.Chart.ChartTitle.Text=="Fixture title","Chart title");break;case "axes":Check(((PowerPoint.Axis)chartShape.Chart.Axes(PowerPoint.XlAxisType.xlValue,PowerPoint.XlAxisGroup.xlPrimary)).AxisTitle.Text=="Fixture axis","Value axis title");break;
            case "plot-area":Check(chartShape.Chart.ChartArea.Format.Fill.ForeColor.RGB==0xFFFFFF&&chartShape.Chart.PlotArea.Format.Fill.ForeColor.RGB==0xF5F2EF,"Chart areas");break;
            case "edit-data":ReadWorkbook(book=>Check(book!=null,"Embedded workbook missing"));break;
            case "sheet-style":ReadWorkbook(book=>Check((bool)book.Worksheets[1].UsedRange.Rows[1].Font.Bold,"Workbook header style"));break;
            case "round-data":ReadWorkbook(book=>Check(Convert.ToDouble(book.Worksheets[1].Cells[2,2].Value2)==10,"Rounding numeric constant"));break;
            case "excel-cagr":ReadWorkbook(book=>Check((bool)book.Worksheets[1].Range["E2"].HasFormula&&Math.Abs(Convert.ToDouble(book.Worksheets[1].Range["E2"].Value2)-0.1)<1e-8,"Excel CAGR formula"));break;
            case "value-line":Check(slide.Shapes.Count==beforeCount+1&&slide.Shapes[slide.Shapes.Count].Type==Office.MsoShapeType.msoGroup&&ChartAnnotationSpec.Read(k=>slide.Shapes[slide.Shapes.Count].Tags[k]).Kind=="reference","Axis-linked value annotation");break;
            case "difference":Check(slide.Shapes.Count==beforeCount+1&&ChartAnnotationSpec.Read(k=>slide.Shapes[slide.Shapes.Count].Tags[k]).From==1&&ChartAnnotationSpec.Read(k=>slide.Shapes[slide.Shapes.Count].Tags[k]).To==2,"Data-driven difference annotation");break;case "cagr":Check(LastText().Contains("10.0"),"CAGR result");break;
            case "scan":Check(messages.Single().Contains("Native charts: 1")&&messages.Single().Contains("Pictures: 1"),"Scan counts");break;
            case "path":Check(LastText()==Deck.FullName,"Presentation path");break;case "theme":Check(shapes[0].Fill.ForeColor.RGB==0x704020&&shapes.All(s=>s.TextFrame.TextRange.Font.Name=="Segoe UI"),"Palette application");break;
            case "image":case "library-browser":Check(slide.Shapes.Count==beforeCount+1&&slide.Shapes[slide.Shapes.Count].Type==Office.MsoShapeType.msoPicture,"Image insertion");break;
            case "library":Check(Deck.Slides.Count==2&&Deck.Slides[2].Shapes[1].TextFrame.TextRange.Text=="Imported fixture","Slide import");break;
            case "template":Check(Deck.SlideMaster.Background.Fill.ForeColor.RGB==0x123456,"Template background application");break;
            case "export-slide":using(var image=Image.FromFile(output))Check(image.Width==1600&&image.Height==900,"PNG export dimensions");break;
            case "save-slides":var saved=app.Presentations.Open(output);try{Check(saved.Slides.Count==1&&saved.Slides[1].Shapes.Count==beforeCount,"Selected-slide export");}finally{saved.Close();Deck.Windows[1].Activate();}break;
            case "duplicate-slide":Check(Deck.Slides.Count==2,"Duplicate slide");break;case "clean-comments":Check(slide.Comments.Count==0,"Comment removal");break;
            case "import-json":Check(slide.Shapes.Count==beforeCount+1&&slide.Shapes[slide.Shapes.Count].HasChart==Office.MsoTriState.msoTrue&&new PowerPointHost(app).ReadChartMetadata(slide.Shapes[slide.Shapes.Count].Tags[PowerPointHost.IdTag]).Categories.Count==3,"Managed JSON import");break;
            case "export-json":using(var stream=File.OpenRead(output)){var model=(ChartDocument)new DataContractJsonSerializer(typeof(ChartDocument)).ReadObject(stream);MetadataCodec.Validate(model);Check(model.ChartId==managedId,"Exported chart identity");}break;
            case "mini-toolbar":Check(toolbarClicked&&shapes.All(s=>Near(s.Left,40)),"Toolbar button action");break;
            case "table-data":Check(slide.Shapes[slide.Shapes.Count].Table.Cell(1,1).Shape.TextFrame.TextRange.Text=="Edited fixture","Table datasheet apply");break;
            case "style-file":Check(shapes.All(s=>s.Fill.ForeColor.RGB==7356448&&Near(s.TextFrame.TextRange.Font.Size,20)),"Style file apply");break;
            case "link-excel":Check(slide.Shapes[slide.Shapes.Count].HasChart==Office.MsoTriState.msoTrue&&ReadLinkedFirst(slide.Shapes[slide.Shapes.Count])==10,"Excel link insertion");break;
            case "refresh-excel":Check(ReadLinkedFirst(chartShape)==42,"Excel refresh");break;case "manage-links":Check(messages.Single().Contains("Demo workbook links: 1"),"Link inventory count");break;
            case "slide-workbook":Check(messages.Single().Contains(chartShape.Name),"Slide workbook list");break;
            default:throw new Exception("Missing feature-specific assertion: "+id);
        }
    }
    private string LastText(){Check(slide.Shapes.Count>beforeCount,"No object inserted");return slide.Shapes[slide.Shapes.Count].TextFrame.TextRange.Text;}
    private void ReadWorkbook(Action<dynamic> assertion){chartShape.Chart.ChartData.Activate();dynamic book=chartShape.Chart.ChartData.Workbook;try{assertion(book);}finally{PowerPointHost.RetryRejectedCall(()=>book.Close(true));}}
    private double ReadLinkedFirst(PowerPoint.Shape shape){shape.Chart.ChartData.Activate();dynamic book=shape.Chart.ChartData.Workbook;try{return Convert.ToDouble(book.Worksheets[1].Cells[2,2].Value2);}finally{PowerPointHost.RetryRejectedCall(()=>book.Close(true));}}
    private void AssertChart()
    {
        Check(slide.Shapes.Count==beforeCount+1,"Chart insertion count");var shape=slide.Shapes[slide.Shapes.Count];Check(shape.HasChart==Office.MsoTriState.msoTrue,"Native chart missing");
        var chart=shape.Chart;var series=(PowerPoint.SeriesCollection)chart.SeriesCollection();Check(series.Count>0,"Series missing");
        Check(PowerPointHost.NativeSeriesValues(series,1).Length==(id=="profile"?2:3),"Series point count");
        if(id!="profile")Check(Convert.ToDouble(PowerPointHost.NativeSeriesValues(series,1)[0],CultureInfo.InvariantCulture)==137.5,"Chart must use linked source value, not sample data");
        if(id=="combo")Check(series.Item(1).ChartType==Office.XlChartType.xlColumnClustered&&series.Item(2).ChartType==Office.XlChartType.xlLineMarkers,"Combination series types");
        else{
            var expected=new Dictionary<string,Office.XlChartType>{{"column",Office.XlChartType.xlColumnClustered},{"bar",Office.XlChartType.xlBarClustered},{"stacked-column",Office.XlChartType.xlColumnStacked},{"percent-column",Office.XlChartType.xlColumnStacked100},{"stacked-bar",Office.XlChartType.xlBarStacked},{"percent-bar",Office.XlChartType.xlBarStacked100},{"line",Office.XlChartType.xlLineMarkers},{"profile",Office.XlChartType.xlLine},{"area",Office.XlChartType.xlAreaStacked},{"pie",Office.XlChartType.xlPie},{"doughnut",Office.XlChartType.xlDoughnut},{"scatter",Office.XlChartType.xlXYScatter},{"bubble",Office.XlChartType.xlBubble}};
            Check(chart.ChartType==expected[id],"Chart type mismatch");
        }
    }
}
