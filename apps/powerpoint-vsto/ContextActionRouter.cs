using System;
using System.Linq;
using System.Windows.Forms;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
namespace PowerPointIntelligence.AddIn
{
    internal static class ContextActionRouter
    {
        internal static void Run(PowerPoint.Application app,string expected,string id)
        {
            ContextSelectionBridge.Ensure(app,expected);
            if(ElementCommands.Ids.Contains(id)){ElementCommands.Run(app,id);return;}
            switch(id){
                case "chart-series":NativeSeriesControls.Show(app);return;
                case "chart-axes":NativeChartControls.ShowAxis(app);return;
                case "agenda-settings":ChapterAgenda.ShowManager(app);return;
                case "agenda-refresh":ChapterAgenda.RefreshAll(app);return;
                case "annotation-refresh":ChartAnnotations.RefreshSelected(app);return;
                case "template-palette":var slide=app.ActiveWindow.View.Slide as PowerPoint.Slide;string palette=PalettePicker.Choose(slide==null?null:slide.Tags["PPI_TEMPLATE_PALETTE"]);if(palette!=null){ContextSelectionBridge.Ensure(app,expected);TemplateInserter.ChangePalette(app,palette);}return;
                case "template-alignment":var alignment=AlignmentPicker.Choose();if(alignment!=null){ContextSelectionBridge.Ensure(app,expected);TemplateInserter.ChangeAlignment(app,alignment[0],alignment[1]);}return;
                case "data-links":using(var manager=new DataLinkManager(app,(shape,link)=>{app.StartNewUndoEntry();MappedLinkOperations.Refresh(shape,link);}))manager.ShowDialog();return;
            }
            string tool=id=="chart-data"?"edit-data":id=="chart-style"?"chart-style":id=="chart-legend"?"legend":id=="chart-labels"?"labels":id=="chart-data-table"?"data-table":id=="table-data"?"table-data":id=="arrange"?"mini-toolbar":null;
            var command=DemoTools.Catalog.SingleOrDefault(t=>t.Id==tool);if(command==null)throw new InvalidOperationException("Unsupported contextual element action.");new DemoTools(app).Execute(command);
        }
    }
}
