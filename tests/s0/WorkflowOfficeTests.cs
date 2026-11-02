using PowerPointIntelligence.Security;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointIntelligence.AddIn;

internal static class WorkflowOfficeTests
{
    private static int assertions;
    private static void Require(bool condition,string name)
    {
        if(!condition)throw new Exception("FAIL workflow: "+name);
        assertions++;Console.WriteLine("PASS workflow: "+name);
    }
    private static PowerPoint.Shape Field(PowerPoint.Slide slide,string role)
    {
        return slide.Shapes.Cast<PowerPoint.Shape>().Single(s=>s.Tags[PresentationAgenda.OwnerTag]==slide.Tags[PresentationAgenda.OwnerTag]&&s.Tags[PresentationAgenda.RoleTag]==role);
    }
    private static string GroupText(PowerPoint.Shape group)
    {
        return String.Join(" ",Enumerable.Range(1,group.GroupItems.Count).Select(i=>group.GroupItems[i]).Where(s=>s.HasTextFrame==Office.MsoTriState.msoTrue).Select(s=>s.TextFrame.TextRange.Text));
    }
    private static PowerPoint.Shape Annotation(PowerPoint.Slide slide,string kind)
    {
        return slide.Shapes.Cast<PowerPoint.Shape>().Single(s=>s.Tags["PPI_CA_SCHEMA"]=="1"&&String.Equals(s.Tags["PPI_CA_KIND"],kind,StringComparison.OrdinalIgnoreCase));
    }
    private static void SetValues(PowerPoint.Shape chart,double last)
    {
        var all=(PowerPoint.SeriesCollection)chart.Chart.SeriesCollection();
        while(all.Count>1)all.Item(all.Count).Delete();
        if(all.Count==0)all.NewSeries();
        PowerPoint.Series series=all.Item(1);series.Name="Workflow verification";series.Values=new double[]{10,30,last};series.XValues=new string[]{"One","Two","Three"};
        chart.Chart.Refresh();
        var values=PowerPointIntelligence.Host.PowerPointHost.SeriesVector((object)series.Values);
        Require(Math.Abs(Convert.ToDouble(values[2],CultureInfo.InvariantCulture)-last)<.001,"native series accepts edited data");
    }
    internal static void Run(PowerPoint.Application app,string outputDirectory)
    {outputDirectory=PathPolicy.TestPath(outputDirectory);
        if(app==null)throw new ArgumentNullException("app");
        Directory.CreateDirectory(outputDirectory);assertions=0;
        string file=Path.Combine(Path.GetFullPath(outputDirectory),"workflow-native-"+Guid.NewGuid().ToString("N")+".pptx");
        PowerPoint.Presentation owned=null;
        Console.WriteLine("WORKFLOW native Office="+app.Version+"; process="+(Environment.Is64BitProcess?"x64":"x86"));
        Exception primaryFailure=null;
        try
        {
            owned=app.Presentations.Add(Office.MsoTriState.msoTrue);
            owned.PageSetup.SlideWidth=960;owned.PageSetup.SlideHeight=540;
            owned.Windows[1].ViewType=PowerPoint.PpViewType.ppViewNormal;
            var first=owned.Slides.Add(1,PowerPoint.PpSlideLayout.ppLayoutTitleOnly);first.Shapes.Title.TextFrame.TextRange.Text="Overview";
            var second=owned.Slides.Add(2,PowerPoint.PpSlideLayout.ppLayoutTitleOnly);second.Shapes.Title.TextFrame.TextRange.Text="Analysis";
            owned.SectionProperties.AddBeforeSlide(1,"Introduction");owned.SectionProperties.AddBeforeSlide(2,"Evidence");
            int firstId=first.SlideID,secondId=second.SlideID;first.Select();
            var agenda=PresentationAgenda.Insert(app,true);int agendaId=agenda.SlideID;
            Require(owned.Slides.Count==3&&app.ActivePresentation==owned,"agenda inserted into owned current deck");
            Require(agenda.Tags[PresentationAgenda.SchemaTag]=="1","agenda schema persisted on native slide");
            var firstField=Field(agenda,"item:"+firstId);var secondField=Field(agenda,"item:"+secondId);
            Require(firstField.TextFrame.TextRange.Text.Contains("Introduction")&&secondField.TextFrame.TextRange.Text.Contains("Evidence"),"agenda uses actual section names");
            Require(firstField.TextFrame.TextRange.Font.Bold==Office.MsoTriState.msoTrue,"current section highlighted");
            CheckLink(firstField,firstId,first.SlideIndex);CheckLink(secondField,secondId,second.SlideIndex);
            var sentinel=agenda.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,25,490,250,25);sentinel.Name="workflow_unrelated";sentinel.TextFrame.TextRange.Text="Keep this user note";
            owned.SectionProperties.Rename(2,"Results");agenda.Select();PresentationAgenda.RefreshSelected(app,true);
            Require(Field(agenda,"item:"+secondId).TextFrame.TextRange.Text.Contains("Results"),"explicit refresh reflects renamed section");
            Require(agenda.Shapes["workflow_unrelated"].TextFrame.TextRange.Text=="Keep this user note","agenda refresh preserves unrelated objects");
            Field(agenda,"title").TextFrame.TextRange.Text="Editable agenda heading";
            Require(Field(agenda,"title").TextFrame.TextRange.Text=="Editable agenda heading","agenda fields are native editable text");
            second.Select();var chart=second.Shapes.AddChart2(-1,Office.XlChartType.xlColumnClustered,80,130,550,280,false);
            SetValues(chart,60);var axis=(PowerPoint.Axis)chart.Chart.Axes(PowerPoint.XlAxisType.xlValue,PowerPoint.XlAxisGroup.xlPrimary);axis.MinimumScale=0;axis.MaximumScale=100;
            chart.Select();var reference=ChartAnnotations.Reference(app,40);
            Require(reference.Type==Office.MsoShapeType.msoGroup&&reference.GroupItems.Count==2&&GroupText(reference)=="40","reference consists of native line and editable label");
            chart.Select();var difference=ChartAnnotations.Difference(app,1,1,3);
            Require(GroupText(difference).Contains("+50"),"difference derives label from native series");
            var label=Enumerable.Range(1,difference.GroupItems.Count).Select(i=>difference.GroupItems[i]).Single(s=>s.HasTextFrame==Office.MsoTriState.msoTrue);label.TextFrame.TextRange.Text="Custom editable label";
            Require(GroupText(difference)=="Custom editable label","annotation label is editable");
            SetValues(chart,80);difference.Select();difference=ChartAnnotations.RefreshSelected(app);
            Require(GroupText(difference).Contains("+70"),"annotation refresh reads changed native chart data");
            int before=second.Shapes.Count;var copied=chart.Duplicate()[1];difference.Select();bool refused=false;
            try{ChartAnnotations.RefreshSelected(app);}catch(InvalidOperationException ex){refused=ex.Message.IndexOf("ambiguous",StringComparison.OrdinalIgnoreCase)>=0;}
            Require(refused&&second.Shapes.Count==before+1,"duplicate identity refusal preserves chart and annotations");copied.Delete();
            difference.Select();difference=ChartAnnotations.RefreshSelected(app);
            Require(GroupText(difference).Contains("+70"),"refresh recovers after ambiguity removed");
            chart.Select();NativeChartControls.SetAxis(app,new NativeAxisSettings {Automatic=false,Minimum=0,Maximum=120,Major=20,NumberFormat="0.0",Reversed=false});
            axis=(PowerPoint.Axis)chart.Chart.Axes(PowerPoint.XlAxisType.xlValue,PowerPoint.XlAxisGroup.xlPrimary);
            Require(Math.Abs(axis.MaximumScale-120)<.001&&Math.Abs(axis.MajorUnit-20)<.001,"explicit native value axis settings applied");
            agenda.Select();sentinel.Select();NativeFormatting.Text(app,"center");NativeFormatting.Text(app,"middle");NativeFormatting.Text(app,"downward");NativeFormatting.ApplyColor(app,"text",0x563412);
            Require(sentinel.TextFrame.TextRange.ParagraphFormat.Alignment==PowerPoint.PpParagraphAlignment.ppAlignCenter&&sentinel.TextFrame.VerticalAnchor==Office.MsoVerticalAnchor.msoAnchorMiddle&&sentinel.TextFrame.Orientation==Office.MsoTextOrientation.msoTextOrientationDownward&&sentinel.TextFrame.TextRange.Font.Color.RGB==0x563412,"native text formatting and chosen color applied");
            second.Select();chart.Select();NativeSeriesControls.Apply(app,1,Office.XlChartType.xlLineMarkers,false);Require(((PowerPoint.SeriesCollection)chart.Chart.SeriesCollection()).Item(1).ChartType==Office.XlChartType.xlLineMarkers,"selected series representation changed natively");NativeSeriesControls.Apply(app,1,Office.XlChartType.xlColumnClustered,false);
            int qualityCount=owned.Slides.Cast<PowerPoint.Slide>().Sum(slide=>slide.Shapes.Count);PresentationQuality.Inspect(owned);Require(qualityCount==owned.Slides.Cast<PowerPoint.Slide>().Sum(slide=>slide.Shapes.Count),"quality review preserves native objects");
            string chartId=chart.Tags[ChartAnnotations.ChartTag];
            owned.SaveAs(file,PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,Office.MsoTriState.msoFalse);owned.Close();owned=null;
            owned=app.Presentations.Open(file,Office.MsoTriState.msoFalse,Office.MsoTriState.msoFalse,Office.MsoTriState.msoTrue);
            agenda=owned.Slides.FindBySlideID(agendaId);second=owned.Slides.FindBySlideID(secondId);first=owned.Slides.FindBySlideID(firstId);
            Require(Field(agenda,"title").TextFrame.TextRange.Text=="Editable agenda heading","edited agenda heading survives save/reopen");
            Require(agenda.Shapes["workflow_unrelated"].TextFrame.TextRange.Text=="Keep this user note","unrelated note survives save/reopen");
            sentinel=agenda.Shapes["workflow_unrelated"];Require(sentinel.TextFrame.TextRange.ParagraphFormat.Alignment==PowerPoint.PpParagraphAlignment.ppAlignCenter&&sentinel.TextFrame.Orientation==Office.MsoTextOrientation.msoTextOrientationDownward&&sentinel.TextFrame.TextRange.Font.Color.RGB==0x563412,"native text formatting survives save/reopen");
            CheckLink(Field(agenda,"item:"+firstId),firstId,first.SlideIndex);CheckLink(Field(agenda,"item:"+secondId),secondId,second.SlideIndex);
            chart=second.Shapes.Cast<PowerPoint.Shape>().Single(s=>s.HasChart==Office.MsoTriState.msoTrue&&s.Tags[ChartAnnotations.ChartTag]==chartId);
            Require(chart.HasChart==Office.MsoTriState.msoTrue,"chart remains native after save/reopen");
            axis=(PowerPoint.Axis)chart.Chart.Axes(PowerPoint.XlAxisType.xlValue,PowerPoint.XlAxisGroup.xlPrimary);Require(Math.Abs(axis.MaximumScale-120)<.001&&Math.Abs(axis.MajorUnit-20)<.001,"axis settings survive save/reopen");
            reference=Annotation(second,"reference");difference=Annotation(second,"difference");
            Require(GroupText(reference)=="40"&&GroupText(difference).Contains("+70"),"annotation labels and versioned links survive save/reopen");
            second.Select();SetValues(chart,90);difference.Select();difference=ChartAnnotations.RefreshSelected(app);
            Require(GroupText(difference).Contains("+80"),"native chart remains editable and annotation refresh works after reopen");
            agenda.Select();PresentationAgenda.RefreshSelected(app,false);
            Require(Field(agenda,"item:"+firstId).TextFrame.TextRange.Font.Bold==Office.MsoTriState.msoFalse,"highlight can be disabled after reopen");
            Require(Field(agenda,"title").TextFrame.TextRange.Text=="Editable agenda heading","legacy overview custom heading survives refresh");
            string pdf=Path.Combine(Path.GetFullPath(outputDirectory),"workflow-native.pdf");NativePdfExport.Export(owned,pdf);Require(new FileInfo(pdf).Length>5&&System.Text.Encoding.ASCII.GetString(File.ReadAllBytes(pdf),0,4)=="%PDF","native presentation PDF export produced a PDF");
            owned.Save();Console.WriteLine("WORKFLOW "+assertions+" native assertions passed. Evidence deck: "+file);
        }
        catch(Exception error){primaryFailure=error;throw;}
        finally
        {
            try{if(owned!=null){owned.Saved=Office.MsoTriState.msoTrue;owned.Close();}}
            catch(Exception cleanup){if(primaryFailure==null)throw;Console.Error.WriteLine("Workflow cleanup also failed: "+cleanup.GetType().Name+"; HRESULT=0x"+cleanup.HResult.ToString("X8",CultureInfo.InvariantCulture));}
        }
    }
    private static void CheckLink(PowerPoint.Shape shape,int slideId,int slideIndex)
    {
        var action=shape.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick];
        string address=action.Hyperlink.SubAddress??"";var fields=address.Split(new[]{','},3);int actualId,actualIndex;
        // PowerPoint may populate the third, display-name field on save. Identity and index must remain exact.
        Require(action.Action==PowerPoint.PpActionType.ppActionHyperlink&&String.IsNullOrEmpty(action.Hyperlink.Address)&&fields.Length==3&&Int32.TryParse(fields[0],NumberStyles.None,CultureInfo.InvariantCulture,out actualId)&&actualId==slideId&&Int32.TryParse(fields[1],NumberStyles.None,CultureInfo.InvariantCulture,out actualIndex)&&actualIndex==slideIndex,"native hyperlink targets actual slide identity and index");
    }
}
