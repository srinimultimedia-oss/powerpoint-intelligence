using PowerPointIntelligence.Security;
using System;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using System.Xml.Serialization;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointIntelligence.AddIn
{
    [XmlRoot("WorkbookLink", Namespace="urn:powerpoint-intelligence:demo:workbook-link:v1")]
    public sealed class DemoWorkbookLink { public int Version {get;set;} = 1; public string Path {get;set;} }
    internal static class DemoLinks
    {
        internal const string Namespace="urn:powerpoint-intelligence:demo:workbook-link:v1";
        internal static void Insert(PowerPoint.Application app,string path)
        {
            {
                var slide=(PowerPoint.Slide)app.ActiveWindow.View.Slide;
                var shape=slide.Shapes.AddChart2(-1,Office.XlChartType.xlColumnClustered,60,80,600,340,false);
                try
                {
                    Transfer(shape,path);
                    var part=shape.CustomerData.Add();var serializer=new XmlSerializer(typeof(DemoWorkbookLink));
                    using(var writer=new StringWriter(CultureInfo.InvariantCulture)){serializer.Serialize(writer,new DemoWorkbookLink{Path=path});if(!part.LoadXML(writer.ToString()))throw new InvalidDataException("Link metadata was rejected.");}
                    shape.Select(Office.MsoTriState.msoTrue);
                }
                catch{shape.Delete();throw;}
            }
        }
        private static PowerPoint.Shape Selected(PowerPoint.Application app)
        {
            var selection=app.ActiveWindow.Selection;
            if(selection.Type!=PowerPoint.PpSelectionType.ppSelectionShapes||selection.ShapeRange.Count!=1||selection.ShapeRange[1].HasChart!=Office.MsoTriState.msoTrue)throw new InvalidOperationException("Select one linked native chart.");
            return selection.ShapeRange[1];
        }
        internal static void Refresh(PowerPoint.Application app)
        {
            var shape=Selected(app);DemoWorkbookLink link=null;
            foreach(Office.CustomXMLPart part in shape.CustomerData)if(part.NamespaceURI==Namespace){if(link!=null)throw new InvalidDataException("Multiple link records.");using(var reader=System.Xml.XmlReader.Create(new StringReader(part.XML),new System.Xml.XmlReaderSettings{DtdProcessing=System.Xml.DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=16384}))link=(DemoWorkbookLink)new XmlSerializer(typeof(DemoWorkbookLink)).Deserialize(reader);}
            if(link==null||link.Version!=1||String.IsNullOrWhiteSpace(link.Path))throw new InvalidOperationException("No supported demo workbook link on this chart.");
            Transfer(shape,link.Path);
        }
        internal static void Describe(PowerPoint.Application app,Action<string,string> notify)
        {
            int count=0;foreach(PowerPoint.Slide slide in app.ActivePresentation.Slides)foreach(PowerPoint.Shape shape in slide.Shapes)foreach(Office.CustomXMLPart part in shape.CustomerData)if(part.NamespaceURI==Namespace)count++;
            notify("Demo workbook links: "+count+"\nUse Refresh selected Excel link to read the original file again. Missing files fail without rebinding. Links do not refresh automatically.","Local workbook links");
        }
        private static void Transfer(PowerPoint.Shape shape,string path)
        {
            path=PathPolicy.Selected(path);var info=new FileInfo(path);if(!info.Exists||!String.Equals(info.Extension,".xlsx",StringComparison.OrdinalIgnoreCase)||info.Length>20*1024*1024)throw new InvalidDataException("Choose an existing .xlsx file under 20 MB.");
            var chart=shape.Chart;chart.ChartData.Activate();dynamic book=chart.ChartData.Workbook;dynamic source=null;
            try
            {
                dynamic excel=book.Application;source=excel.Workbooks.Open(info.FullName,0,true);
                dynamic range=source.Worksheets[1].UsedRange;int rows=(int)range.Rows.Count,cols=(int)range.Columns.Count;
                if(rows<2||cols<2||rows>50||cols>15)throw new InvalidDataException("Use 2–50 rows and 2–15 columns: headers, category labels and numeric data.");
                var values=(Array)(object)range.Value2;if(values.Rank!=2)throw new InvalidDataException("Expected a rectangular data range.");
                var data=new object[rows,cols];int r0=values.GetLowerBound(0),c0=values.GetLowerBound(1);
                for(int r=0;r<rows;r++)for(int c=0;c<cols;c++)
                {
                    object value=values.GetValue(r+r0,c+c0);
                    if(r==0||c==0){string label=Convert.ToString(value,CultureInfo.InvariantCulture);if(label.Length>200)throw new InvalidDataException("Label too long.");data[r,c]=label;}
                    else if(value==null)data[r,c]=null;
                    else{if(!(value is double))throw new InvalidDataException("Data values must be numeric constants or evaluated numeric formulas.");double number=(double)value;if(Double.IsNaN(number)||Double.IsInfinity(number))throw new InvalidDataException("Invalid numeric data.");data[r,c]=number;}
                }
                source.Close(false);source=null;
                dynamic sheet=book.Worksheets[1];sheet.Cells.Clear();
                for(int r=0;r<rows;r++)for(int c=0;c<cols;c++){if(r==0||c==0)sheet.Cells[r+1,c+1].NumberFormat="@";sheet.Cells[r+1,c+1].Value2=data[r,c];}
                dynamic destination=sheet.Range[sheet.Cells[1,1],sheet.Cells[rows,cols]];chart.SetSourceData((string)destination.Address[true,true,1,true],PowerPoint.XlRowCol.xlColumns);
            }
            finally{try{if(source!=null)source.Close(false);}finally{Host.PowerPointHost.RetryRejectedCall(()=>book.Close(true));}}
            chart.Refresh();
        }
    }
}
