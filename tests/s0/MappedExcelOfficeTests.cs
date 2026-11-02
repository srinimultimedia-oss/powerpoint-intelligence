using System;
using System.IO;
using System.Linq;
using System.Globalization;
using PowerPointIntelligence.AddIn;
using Office=Microsoft.Office.Core;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
internal static class MappedExcelOfficeTests
{
    private static void Require(bool pass,string message){if(!pass)throw new Exception("FAIL mapped Excel: "+message);Console.WriteLine("PASS mapped Excel: "+message);}
    private static void Workbook(string path,bool create,Action<object> action)
    {
        dynamic excel=null,book=null;try{excel=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true));excel.Visible=false;excel.DisplayAlerts=false;excel.AutomationSecurity=3;excel.AskToUpdateLinks=false;book=create?excel.Workbooks.Add():excel.Workbooks.Open(path,0,false);action((object)book);if(create)book.SaveAs(path,51);else book.Save();}finally{try{if(book!=null)book.Close(false);}finally{if(excel!=null)System.Runtime.InteropServices.Marshal.ReleaseComObject((object)excel);}}
    }
    internal static void Run(PowerPoint.Application app,string output)
    {
        string path=Path.Combine(Path.GetFullPath(output),"mapped-range.xlsx"),deckPath=Path.Combine(Path.GetFullPath(output),"mapped-range.pptx");
        Workbook(path,true,book=>{dynamic native=book;dynamic sheet=native.Worksheets[1];sheet.Name="Explicit data";string[,] values={{"Actual","Category","Notes","Plan"},{"10","Q1","unused","12"},{"20","Q2","unused","22"},{"30","Q3","unused","32"}};for(int r=0;r<4;r++)for(int c=0;c<4;c++)sheet.Cells[r+4,c+3].Value2=values[r,c];});
        PowerPoint.Presentation owned=null;
        try{
            var link=DataSources.Create("excel",path);link.Mapping=new ExcelRangeMapping {Worksheet="Explicit data",Address="C4:F7",Category=1,Series=new[]{0,3},Headers=new[]{"Actual","Category","Notes","Plan"}};var table=LinkedChartData.Load(link);Require(table.Headers.SequenceEqual(new[]{"Category","Actual","Plan"})&&table.Number(0,1)==10,"explicit non-A1 range and column mapping");
            owned=app.Presentations.Add(Office.MsoTriState.msoTrue);var slide=owned.Slides.Add(1,PowerPoint.PpSlideLayout.ppLayoutBlank);slide.Select();link=DataSources.Save(owned,link);var shape=slide.Shapes.AddChart2(-1,Office.XlChartType.xlColumnClustered,60,80,600,340,false);LinkedChartData.Write(shape,table,link);int chartId=shape.Id;
            Workbook(path,false,book=>{dynamic native=book;native.Worksheets["Explicit data"].Cells[5,3].Value2=42;});MappedLinkOperations.Refresh(shape,link);var series=(PowerPoint.SeriesCollection)shape.Chart.SeriesCollection();var actual=PowerPointIntelligence.Host.PowerPointHost.NativeSeriesValues(series,1);Require(Math.Abs(Convert.ToDouble(actual.GetValue(actual.GetLowerBound(0)),CultureInfo.InvariantCulture)-42)<.001,"mapped source refresh commits changed native chart values");
            MappedLinkOperations.Refresh(shape,link);int references=0;foreach(Office.CustomXMLPart part in shape.CustomerData)if(part.NamespaceURI==DataSources.Namespace)references++;Require(references==1,"repeated refresh reuses one source association");
            Workbook(path,false,book=>{dynamic native=book;native.Worksheets["Explicit data"].Cells[4,3].Value2="Changed actual";});bool refused=false;try{MappedLinkOperations.Refresh(shape,link);}catch(InvalidDataException){refused=true;}actual=PowerPointIntelligence.Host.PowerPointHost.NativeSeriesValues(series,1);Require(refused&&Convert.ToDouble(actual.GetValue(actual.GetLowerBound(0)),CultureInfo.InvariantCulture)==42,"changed headers stop refresh before chart mutation");
            owned.SaveAs(deckPath,PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);owned.Close();owned=null;owned=app.Presentations.Open(deckPath);var restored=DataSources.Read(owned).Single();Require(restored.Mapping.Key==link.Mapping.Key,"range and columns persist across save/reopen");shape=owned.Slides[1].Shapes.Cast<PowerPoint.Shape>().Single(s=>s.Id==chartId);actual=PowerPointIntelligence.Host.PowerPointHost.NativeSeriesValues((PowerPoint.SeriesCollection)shape.Chart.SeriesCollection(),1);Require(Convert.ToDouble(actual.GetValue(actual.GetLowerBound(0)),CultureInfo.InvariantCulture)==42,"edited native linked chart values persist");
        }finally{if(owned!=null){owned.Saved=Office.MsoTriState.msoTrue;owned.Close();}}
    }
}
