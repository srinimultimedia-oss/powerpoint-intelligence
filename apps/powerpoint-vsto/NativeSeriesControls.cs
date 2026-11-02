using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Office=Microsoft.Office.Core;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
namespace PowerPointIntelligence.AddIn
{
    internal static class NativeSeriesControls
    {
        internal static readonly Office.XlChartType[] Allowed={Office.XlChartType.xlColumnClustered,Office.XlChartType.xlLine,Office.XlChartType.xlLineMarkers,Office.XlChartType.xlArea};
        internal static void Validate(int index,int count,Office.XlChartType type){if(index<1||index>count||count>100||!Allowed.Contains(type))throw new InvalidDataException("Choose an existing series (maximum 100) and a supported 2D column/line/area type.");}
        internal static void Apply(PowerPoint.Application app,int index,Office.XlChartType type,bool secondary)
        {
            var shape=NativeChartControls.Selected(app);foreach(Office.CustomXMLPart part in shape.CustomerData)if(part.NamespaceURI==Host.MetadataCodec.Namespace)throw new InvalidOperationException("The S0 chart contract is clustered-column only. Use a linked native chart for series conversion.");
            var chart=shape.Chart;var collection=(PowerPoint.SeriesCollection)chart.SeriesCollection();Validate(index,collection.Count,type);for(int seriesIndex=1;seriesIndex<=collection.Count;seriesIndex++){PowerPoint.Series item=collection.Item(seriesIndex);if(!Allowed.Contains(item.ChartType))throw new InvalidOperationException("This editor supports unstacked column, line and area series only; existing chart semantics were preserved.");}
            var series=collection.Item(index);var oldType=series.ChartType;var oldAxis=series.AxisGroup;app.StartNewUndoEntry();try{series.ChartType=type;series.AxisGroup=secondary?PowerPoint.XlAxisGroup.xlSecondary:PowerPoint.XlAxisGroup.xlPrimary;chart.Refresh();if(series.ChartType!=type||series.AxisGroup!=(secondary?PowerPoint.XlAxisGroup.xlSecondary:PowerPoint.XlAxisGroup.xlPrimary))throw new InvalidOperationException("PowerPoint did not retain the requested series settings.");}
            catch(Exception error){try{series.ChartType=oldType;series.AxisGroup=oldAxis;}catch(Exception rollback){throw new InvalidOperationException("Series update failed and rollback was incomplete. Use PowerPoint Undo.",new AggregateException(error,rollback));}throw new InvalidOperationException("Series update failed; the prior series type and axis group were restored.",error);}
        }
        internal static void Show(PowerPoint.Application app)
        {
            var chart=NativeChartControls.Selected(app).Chart;var collection=(PowerPoint.SeriesCollection)chart.SeriesCollection();if(collection.Count<1||collection.Count>100)throw new InvalidOperationException("Choose a chart with 1–100 series.");
            using(var form=new Form {Text="Series type and axis · Intelligence",Width=480,Height=250,StartPosition=FormStartPosition.CenterParent,FormBorderStyle=FormBorderStyle.FixedDialog,MinimizeBox=false,MaximizeBox=false})
            {
                var series=new ComboBox {Left=135,Top=20,Width=310,DropDownStyle=ComboBoxStyle.DropDownList};var type=new ComboBox {Left=135,Top=65,Width=310,DropDownStyle=ComboBoxStyle.DropDownList};var secondary=new CheckBox {Left=135,Top=108,Width=310,Text="Use secondary value axis"};for(int seriesIndex=1;seriesIndex<=collection.Count;seriesIndex++){PowerPoint.Series item=collection.Item(seriesIndex);series.Items.Add(seriesIndex+" · "+item.Name);}type.Items.AddRange(new object[]{"Clustered column","Line","Line with markers","Area"});
                series.SelectedIndexChanged+=(s,e)=>{var selected=collection.Item(series.SelectedIndex+1);type.SelectedIndex=Array.IndexOf(Allowed,selected.ChartType);secondary.Checked=selected.AxisGroup==PowerPoint.XlAxisGroup.xlSecondary;};series.SelectedIndex=0;
                var apply=new Button {Text="Apply",Left=265,Top=153,Width=85};var cancel=new Button {Text="Cancel",Left=360,Top=153,Width=85,DialogResult=DialogResult.Cancel};form.Controls.AddRange(new Control[]{new Label {Text="Series",Left=20,Top=24},new Label {Text="Representation",Left=20,Top=69,Width=110},series,type,secondary,apply,cancel});form.CancelButton=cancel;form.AcceptButton=apply;apply.Click+=(s,e)=>{if(type.SelectedIndex<0){MessageBox.Show(form,"Choose an available representation.","Series settings");return;}form.DialogResult=DialogResult.OK;};if(form.ShowDialog()==DialogResult.OK)Apply(app,series.SelectedIndex+1,Allowed[type.SelectedIndex],secondary.Checked);
            }
        }
    }
    internal static class NativePdfExport
    {
        internal static void Export(PowerPoint.Presentation presentation,string path)
        {
            if(presentation==null||presentation.Slides.Count==0)throw new InvalidOperationException("Open a presentation containing slides first.");
            if(String.IsNullOrWhiteSpace(path)||!Path.IsPathRooted(path)||!String.Equals(Path.GetExtension(path),".pdf",StringComparison.OrdinalIgnoreCase))throw new InvalidDataException("Choose an absolute PDF output path.");
            presentation.ExportAsFixedFormat(path,PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentPrint,Office.MsoTriState.msoFalse,PowerPoint.PpPrintHandoutOrder.ppPrintHandoutVerticalFirst,PowerPoint.PpPrintOutputType.ppPrintOutputSlides,Office.MsoTriState.msoFalse,null,PowerPoint.PpPrintRangeType.ppPrintAll);
        }
        internal static void Show(PowerPoint.Application app)
        {using(var dialog=new SaveFileDialog {Title="Export current presentation as PDF",Filter="PDF presentation|*.pdf",DefaultExt="pdf",AddExtension=true,OverwritePrompt=true,FileName="presentation.pdf"})if(dialog.ShowDialog()==DialogResult.OK)Export(app.ActivePresentation,dialog.FileName);}
    }
}
