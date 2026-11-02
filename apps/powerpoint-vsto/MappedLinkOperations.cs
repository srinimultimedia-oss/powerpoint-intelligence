using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Office=Microsoft.Office.Core;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
namespace PowerPointIntelligence.AddIn
{
    internal static class MappedLinkOperations
    {
        private static List<Office.CustomXMLPart> Parts(PowerPoint.Shape shape)
        {var result=new List<Office.CustomXMLPart>();foreach(Office.CustomXMLPart part in shape.CustomerData)if(part.NamespaceURI==DataSources.Namespace)result.Add(part);return result;}
        internal static void ValidateRefresh(PowerPoint.Shape shape,DataSourceLink link)
        {
            if(link.Mapping==null)throw new InvalidOperationException("Refresh requires an explicit worksheet/range mapping. Use Relink selected object and confirm the mapping first.");
            if(shape.HasChart!=Office.MsoTriState.msoTrue)throw new InvalidOperationException("Mapped refresh currently supports native charts. Table/diagram refresh is not implemented.");
            var refs=Parts(shape);if(refs.Count!=1||!DataLinkInventory.Same(DataSources.Decode(refs[0].XML),link))throw new InvalidDataException("Source reference is damaged, ambiguous or changed. Refresh stopped.");
            if(shape.Tags["PPI_SOURCE_ROWS"]!="0"&&shape.Tags["PPI_SOURCE_ROWS"]!="1"||shape.Tags["PPI_SOURCE_NUMERIC_X"]!="0"&&shape.Tags["PPI_SOURCE_NUMERIC_X"]!="1")throw new InvalidOperationException("This older chart lacks orientation metadata. Recreate it from the explicitly mapped source before refreshing.");
            var state=DataLinkInventory.ProbeLocal(link.Path);if(state.State!=LocalLinkFileState.Available)throw new InvalidOperationException("Source is missing, remote or unavailable; choose and confirm a local replacement first.");
        }
        internal static void Refresh(PowerPoint.Shape shape,DataSourceLink link)
        {
            ValidateRefresh(shape,link);var data=LinkedChartData.Load(link);bool rows=shape.Tags["PPI_SOURCE_ROWS"]=="1",xy=shape.Tags["PPI_SOURCE_NUMERIC_X"]=="1";var chart=shape.Chart;
            var collection=(PowerPoint.SeriesCollection)chart.SeriesCollection();int expected=rows?data.Rows.Length:data.Headers.Length-1;var type=chart.ChartType;
            if(type==Office.XlChartType.xlBubble||type==Office.XlChartType.xlPie||type==Office.XlChartType.xlDoughnut)throw new InvalidOperationException("Mapped refresh does not yet support single-series or bubble projections. Reinsert this chart to preserve explicit series semantics.");
            if(collection.Count!=expected)throw new InvalidOperationException("The series shape differs from the current chart. Refresh stopped; explicitly rebuild the chart.");
            if(xy)throw new InvalidOperationException("Mapped XY refresh would rebuild series and may lose formatting. Reinsert this chart explicitly.");
            for(int seriesIndex=1;seriesIndex<=collection.Count;seriesIndex++){PowerPoint.Series series=collection.Item(seriesIndex);if(series.ChartType!=type||series.AxisGroup!=PowerPoint.XlAxisGroup.xlPrimary)throw new InvalidOperationException("Mixed/secondary-axis charts need a reviewed series mapping before refresh.");}
            // This is an explicit user action. Native workbook changes are grouped by the
            // caller's Undo boundary; a COM failure is reported, never called a successful refresh.
            LinkedChartData.Write(shape,data,link,rows,false);
        }
        internal static void RelinkSelected(PowerPoint.Application app)
        {
            var shape=NativeChartControls.Selected(app);var refs=Parts(shape);if(refs.Count!=1)throw new InvalidOperationException("Choose a chart with exactly one valid source reference.");var original=DataSources.Decode(refs[0].XML);var replacement=MappedExcelData.Choose(app,false);if(replacement==null)return;
            string description="Replace only this chart's source reference?\n\n"+Path.GetFileName(original.Path)+" → "+Path.GetFileName(replacement.Path)+"\n"+replacement.Mapping.Worksheet+"!"+replacement.Mapping.Address+"\n\nChart values stay unchanged until you explicitly refresh. Old source records are retained.";
            if(MessageBox.Show(description,"Confirm mapped source replacement",MessageBoxButtons.OKCancel,MessageBoxIcon.Question)!=DialogResult.OK)return;
            // Read and validate the complete replacement before touching presentation metadata.
            LinkedChartData.Load(replacement);var saved=DataSources.Save(app.ActivePresentation,replacement);string xml=DataSources.Encode(saved);app.StartNewUndoEntry();Office.CustomXMLPart added=null;bool removalAttempted=false;
            try{added=shape.CustomerData.Add();if(!added.LoadXML(xml)||!DataLinkInventory.Same(DataSources.Decode(added.XML),saved))throw new InvalidDataException("Replacement source metadata could not be verified.");removalAttempted=true;refs[0].Delete();}
            catch(Exception error){if(removalAttempted)throw new InvalidOperationException("Replacement metadata was retained, but removal of the old reference could not be confirmed. Rescan; ambiguous references disable refresh. Use Undo if necessary.",error);if(added!=null)try{added.Delete();}catch(Exception cleanup){throw new InvalidOperationException("Relinking failed and metadata cleanup was incomplete. Preserve the chart and use Undo.",new AggregateException(error,cleanup));}throw new InvalidOperationException("Relinking failed. Rescan the source inventory before another action.",error);}
        }
    }
}
