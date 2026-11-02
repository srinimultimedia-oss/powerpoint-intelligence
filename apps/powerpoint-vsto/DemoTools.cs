using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointIntelligence.AddIn
{
    internal static class GridPlanner
    {
        internal static System.Drawing.PointF[] Plan(float slideWidth,float slideHeight,float[] widths,float[] heights)
        {
            if(widths==null||heights==null||widths.Length==0||widths.Length!=heights.Length||!Finite(slideWidth)||!Finite(slideHeight)||slideWidth<=80||slideHeight<=100||widths.Any(w=>!Finite(w)||w<=0)||heights.Any(h=>!Finite(h)||h<=0))throw new InvalidOperationException("Use valid slide and object dimensions.");
            foreach(int columns in Enumerable.Range(1,Math.Min(12,widths.Length)).OrderBy(c=>Math.Abs(c-3)))
            {
                int rows=(int)Math.Ceiling(widths.Length/(double)columns);var cw=new float[columns];var rh=new float[rows];
                for(int i=0;i<widths.Length;i++){cw[i%columns]=Math.Max(cw[i%columns],widths[i]);rh[i/columns]=Math.Max(rh[i/columns],heights[i]);}
                if(cw.Sum()+16*(columns-1)>slideWidth-80||rh.Sum()+16*(rows-1)>slideHeight-100)continue;
                var result=new System.Drawing.PointF[widths.Length];
                for(int i=0;i<widths.Length;i++)result[i]=new System.Drawing.PointF(40+cw.Take(i%columns).Sum()+16*(i%columns),60+rh.Take(i/columns).Sum()+16*(i/columns));
                return result;
            }
            throw new InvalidOperationException("Selected objects do not fit inside this slide without resizing. Select fewer objects or resize them first.");
        }
        private static bool Finite(float value){return !Single.IsNaN(value)&&!Single.IsInfinity(value);}
    }
    internal sealed class DemoTool
    {
        public string Id, Category, Name, Description;
        public Action<DemoTools> Execute;
        public override string ToString() { return Name; }
    }

    // Independent local commands. Charts use embedded workbooks; diagrams use
    // editable native shapes. Demo data is explicit, never claimed as researched data.
    internal sealed class DemoTools
    {
        private readonly PowerPoint.Application app;
        private readonly DemoInputs inputs;
        internal DemoTools(PowerPoint.Application application, DemoInputs inputs=null) { app = application; this.inputs=inputs??new DemoInputs(); }
        private PowerPoint.Slide Slide { get { return (PowerPoint.Slide)app.ActiveWindow.View.Slide; } }
        private List<PowerPoint.Shape> Selected()
        {
            var selection = app.ActiveWindow.Selection;
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes) throw new InvalidOperationException("Select one or more objects first.");
            var result = new List<PowerPoint.Shape>();
            for (int i = 1; i <= selection.ShapeRange.Count; i++) result.Add(selection.ShapeRange[i]);
            return result;
        }
        private PowerPoint.Chart Chart()
        {
            var shapes = Selected();
            if (shapes.Count != 1 || shapes[0].HasChart != Office.MsoTriState.msoTrue) throw new InvalidOperationException("Select one native chart first.");
            return shapes[0].Chart;
        }
        internal static readonly DemoTool[] Catalog = BuildCatalog();
        private static DemoTool[] BuildCatalog()
        {
            var tools = new List<DemoTool>();
            Action<string,string,string,string,Action<DemoTools>> add = (id,category,name,description,execute) => tools.Add(new DemoTool { Id=id, Category=category, Name=name, Description=description, Execute=execute });
            add("demo-tour","Start here","Generate linked report","Add chart report slides using a linked source in the current presentation.",t=>t.CreateTour());
            add("column","Charts","Clustered column","Native chart from the linked data source; edit values through Edit Data.", t=>t.InsertChart(Office.XlChartType.xlColumnClustered));
            add("bar","Charts","Clustered bar","Native horizontal comparison chart from the linked data source.", t=>t.InsertChart(Office.XlChartType.xlBarClustered));
            add("stacked-column","Charts","Stacked column","Native stacked column chart from the linked data source.", t=>t.InsertChart(Office.XlChartType.xlColumnStacked));
            add("percent-column","Charts","100% stacked column","Native normalized stacked column chart from the linked data source.", t=>t.InsertChart(Office.XlChartType.xlColumnStacked100));
            add("stacked-bar","Charts","Stacked bar","Native stacked bar chart from the linked data source.", t=>t.InsertChart(Office.XlChartType.xlBarStacked));
            add("percent-bar","Charts","100% stacked bar","Native normalized stacked bar chart from the linked data source.", t=>t.InsertChart(Office.XlChartType.xlBarStacked100));
            add("line","Charts","Line","Native line chart from the linked data source.", t=>t.InsertChart(Office.XlChartType.xlLineMarkers));
            add("profile","Charts","Profile","Native line comparison with series plotted in rows; not a rotated value-axis renderer.", t=>t.InsertChart(Office.XlChartType.xlLine, true));
            add("area","Charts","Area","Native stacked area chart from the linked data source.", t=>t.InsertChart(Office.XlChartType.xlAreaStacked));
            add("pie","Charts","Pie","Native pie chart from the first linked numeric series.", t=>t.InsertChart(Office.XlChartType.xlPie));
            add("doughnut","Charts","Doughnut","Native doughnut chart from the first linked numeric series.", t=>t.InsertChart(Office.XlChartType.xlDoughnut));
            add("scatter","Charts","Scatter","Native numeric XY chart from the linked data source.", t=>t.InsertChart(Office.XlChartType.xlXYScatter));
            add("bubble","Charts","Bubble","Native chart from linked X/Y and Size columns.", t=>t.InsertChart(Office.XlChartType.xlBubble));
            add("combo","Charts","Column and line","Native column chart with its second series rendered as a line.", t=>t.InsertCombo());
            add("waterfall","Charts","Waterfall diagram","Linked signed movements from zero rendered as editable shapes.", t=>t.InsertWaterfall());
            add("mekko","Charts","Mekko diagram (%)","Linked nonnegative values: category totals set widths; series shares set percent heights.", t=>t.InsertMekko(true));
            add("mekko-units","Charts","Mekko diagram (units)","Linked nonnegative category totals set widths and heights.", t=>t.InsertMekko(false));
            add("gantt","Charts","Gantt diagram","Linked Task, numeric Start and positive Duration columns.", t=>t.InsertGantt());
            add("text","Elements","Text box","Insert a native editable text box.", t=>TextShapeElements.Insert(t.app,"textbox"));
            add("rounded","Elements","Rounded rectangle","Insert an editable rounded rectangle.", t=>TextShapeElements.Insert(t.app,"roundedRectangle"));
            add("process","Elements","Process flow","Insert an editable two-row process grid with managed cell controls.", t=>ProcessFlowTools.Create(t.app));
            add("checkbox","Elements","Checkbox","Insert an editable vector checkbox with empty, check and cross states.", t=>IndicatorElements.Insert(t.app,"checkbox"));
            add("harvey","Elements","Harvey ball","Insert an editable segmented vector indicator; use Element Controls to change its fill.", t=>IndicatorElements.Insert(t.app,"harvey"));
            add("table","Elements","Table","Insert a native editable table from linked headers/text/values, limited to 400 cells.", t=>t.InsertTable());
            add("agenda","Elements","Agenda chapter","Insert a native editable chapter slide and synchronize managed chapter agendas.", t=>t.Agenda());
            add("image","Assets","Image from file","Insert a local image chosen by you.", t=>t.InsertImage());
            add("symbol","Elements","Insert symbol","Choose text or a Unicode symbol to insert.", t=> { var text=t.Prompt("Symbol or text", "→"); if(text!=null)t.Text(text,80,100,200,70,32); });
            foreach (var mode in new[]{"left","center","right","top","middle","bottom","width","height","size","horizontal","vertical","swap","scale-up","scale-down","rotate","flip-horizontal","flip-vertical","duplicate","group","ungroup","front","back"})
            {
                var captured = mode;
                add(mode,"Arrange",mode.Replace('-',' '),"Applies to selected native objects. Use PowerPoint Undo to revert.",t=>t.Arrange(captured));
            }
            add("auto-layout","Arrange","Grid layout","Arrange selected objects using their dimensions and slide bounds.",t=>t.Grid());
            add("paste-text","Text","Paste lines into objects","Enter one line per selected text object; validates the count before editing.",t=>t.PasteText());
            add("font","Text","Replace font","Apply a chosen font family to selected text objects.",t=>t.Font(false));
            add("font-size","Text","Resize font","Apply a chosen point size to selected text objects.",t=>t.Font(true));
            add("decimal","Text","Switch decimal separators","Swap dots and commas only between digits in selected text; chart data is unchanged.",t=>t.SwitchDecimals());
            add("language","Text","Proofing language","Apply an Office language ID to selected text (for example 1033 English or 1031 German).",t=>t.Language());
            add("edit-data","Data","Edit chart data","Open the selected native chart's embedded Excel workbook.",t=>t.Chart().ChartData.Activate());
            add("legend","Chart options","Toggle legend","Show or hide the selected native chart's legend.",t=>{var c=t.Chart();c.HasLegend=!c.HasLegend;});
            add("labels","Chart options","Add value labels","Apply native data labels to the selected chart.",t=>t.Chart().ApplyDataLabels());
            add("data-table","Chart options","Toggle data table","Show or hide a native chart data table where supported.",t=>{var c=t.Chart();c.HasDataTable=!c.HasDataTable;});
            add("chart-style","Chart options","Chart style","Apply a numbered native Office chart style (1–48).",t=>t.ChartStyle());
            add("value-line","Chart options","Value line annotation","Insert a native editable reference line mapped to the selected chart axis; explicit annotation refresh supported.",t=>t.ValueLine());
            add("difference","Chart options","Difference arrow annotation","Calculate a difference between native series points and insert an editable annotation; explicit refresh supported.",t=>t.Difference());
            add("cagr","Chart options","CAGR annotation","Calculate CAGR from entered start/end/years and insert editable text.",t=>t.Cagr());
            add("scan","Data","Chart scanner","Count native charts and pictures in the current presentation; no chart recognition from pixels.",t=>t.Scan());
            add("path","Assets","File path on slide","Insert the current presentation's path as editable text.",t=>t.Text(t.app.ActivePresentation.FullName,24,t.app.ActivePresentation.PageSetup.SlideHeight-35,650,24,10));
            add("library","Assets","Import slides","Choose a local PowerPoint file and insert its slides after the active slide.",t=>t.ImportSlides());
            add("template","Assets","Apply template","Choose and apply a local .potx template to this presentation.",t=>t.ApplyTemplate());
            add("export-slide","Assets","Export slide image","Save the current slide to a PNG file chosen by you.",t=>t.ExportSlide());
            add("theme","Assets","Apply local brand colors","Apply an original navy/teal palette to selected shapes and Segoe UI to their text. Does not change the master theme.",t=>t.Brand());
            add("axes","Chart options","Toggle value-axis title","Set a native value-axis title where the chart supports axes.",t=>t.AxisTitle());
            add("plot-area","Chart options","Style chart and plot area","Apply a white chart area and pale plot area to the selected native chart.",t=>{var c=t.Chart();c.ChartArea.Format.Fill.ForeColor.RGB=0xFFFFFF;c.PlotArea.Format.Fill.ForeColor.RGB=0xF5F2EF;});
            add("chart-title","Chart options","Chart title","Set the selected native chart title.",t=>{var value=t.Prompt("Chart title","Chart title");if(value!=null){var c=t.Chart();c.HasTitle=true;c.ChartTitle.Text=value;}});
            add("duplicate-slide","Slides","Duplicate slide","Duplicate the current slide using PowerPoint's native operation.",t=>t.Slide.Duplicate());
            add("save-slides","Slides","Save selected slides","Save selected native slides to a new PowerPoint file; no email is sent.",t=>t.SaveSlides());
            add("clean-comments","Slides","Remove presentation comments","Explicitly confirm removal of legacy comments from this presentation. Modern threaded comments are not handled.",t=>t.CleanComments());
            add("round-data","Data","Round chart workbook data","Round numeric constants in the selected chart's embedded workbook, preserving formulas. Maximum 1,000 used cells.",t=>t.RoundData());
            add("sheet-style","Data","Style chart workbook header","Apply a bold shaded header to the selected chart's embedded workbook.",t=>t.StyleSheet());
            add("import-json","Automation","Insert chart from JSON","Choose a bounded, validated S0 ChartDocument JSON file. Inserts a managed clustered-column chart; not an arbitrary COM command interface.",t=>t.ImportJson());
            add("export-json","Automation","Export selected chart JSON","Export validated metadata of a managed S0 chart to a JSON file. Other demo charts do not carry this model.",t=>t.ExportJson());
            add("mini-toolbar","Arrange","Quick arrange toolbar","Open a floating toolbar for selected objects. Select slide objects, then use alignment or size actions.",t=>t.MiniToolbar());
            add("table-data","Data","Edit table datasheet","Edit a selected native table in a grid, then Apply. Limited to 400 cells.",t=>t.TableData());
            add("library-browser","Assets","Browse local asset library","Browse a folder for local images, presentations and templates. No network library service or deletion is provided.",t=>t.Library());
            add("agenda-layout","Assets","Agenda table of contents","Insert a synchronized table of contents with chapter hyperlinks and no current-chapter highlight.",t=>t.AgendaCurrent());
            add("style-file","Assets","Apply style file","Choose an original JSON style with FontName, FontSize, FillRgb and TextRgb; applies to selected text shapes.",t=>t.StyleFile());
            add("excel-cagr","Data","Insert Excel CAGR formula","Write a CAGR formula to a chosen cell in the selected chart workbook; enter positive inputs and a single-cell address.",t=>t.ExcelCagr());
            add("link-excel","Data","Chart from Excel file","Insert a native column chart from the first worksheet of a chosen local .xlsx. Manual refresh; maximum 50 rows by 15 columns. Source path is stored in the presentation.",t=>{var path=t.OpenFile("Excel workbook|*.xlsx");if(path!=null)DemoLinks.Insert(t.app,path);});
            add("refresh-excel","Data","Refresh selected Excel link","Read the saved file path again. Missing sources fail without automatically rebinding. Does not refresh in the background.",t=>DemoLinks.Refresh(t.app));
            add("manage-links","Data","Manage local workbook links","Count workbook links and show the manual-refresh behavior.",t=>DemoLinks.Describe(t.app, t.Notify));
            add("slide-workbook","Data","Current slide workbook list","List charts on this slide; select each chart to edit its embedded workbook. No unified shared slide workbook is created.",t=>t.Notify(String.Join(Environment.NewLine,Enumerable.Range(1,t.Slide.Shapes.Count).Select(i=>t.Slide.Shapes[i]).Where(s=>s.HasChart==Office.MsoTriState.msoTrue).Select(s=>s.Name)),"Current slide chart workbooks"));
            return tools.ToArray();
        }
        internal void Execute(DemoTool tool) { app.StartNewUndoEntry(); tool.Execute(this); }
        private void CreateTour()
        {
            reportLink=LinkedChartData.Choose(app.ActivePresentation);reportData=LinkedChartData.Load(reportLink);
            try {var deck=app.ActivePresentation;
                foreach(var id in new[]{"column","bar","stacked-column","percent-column","line","area","pie"}) {
                    deck.Slides.Add(deck.Slides.Count+1,PowerPoint.PpSlideLayout.ppLayoutBlank).Select();
                    Catalog.Single(t=>t.Id==id).Execute(this);
                }
            } finally {reportLink=null;reportData=null;}
        }
        private DataSourceLink reportLink;
        private LinkedTable reportData;
        private LinkedTable Source(out DataSourceLink link)
        {link=reportLink??LinkedChartData.Choose(app.ActivePresentation);return reportData??LinkedChartData.Load(link);}
        private PowerPoint.Shape Text(string value,float x,float y,float w,float h,int size)
        {
            var shape=Slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,x,y,w,h);
            shape.TextFrame.TextRange.Text=value; shape.TextFrame.TextRange.Font.Size=size;
            shape.TextFrame.TextRange.Font.Name="Segoe UI"; return shape;
        }
        private PowerPoint.Shape Box(Office.MsoAutoShapeType kind,float x,float y,float w,float h,string text)
        {
            var shape=Slide.Shapes.AddShape(kind,x,y,w,h);shape.TextFrame.TextRange.Text=text;
            shape.Fill.ForeColor.RGB=0x807000; shape.TextFrame.TextRange.Font.Color.RGB=0xFFFFFF;
            shape.TextFrame.TextRange.Font.Size=18;shape.TextFrame.TextRange.Font.Name="Segoe UI";return shape;
        }
        internal void InsertChart(Office.XlChartType type,bool rows=false,bool combo=false)
        {
            DataSourceLink link;var data=Source(out link);
            bool xy=type==Office.XlChartType.xlXYScatter||type==Office.XlChartType.xlBubble;
            if(combo&&data.Headers.Length<3)throw new InvalidOperationException("A combination chart needs at least two numeric series.");
            double[] x=null;
            if(xy){x=data.Rows.Select(r=>{double value;if(!Double.TryParse(r[0],NumberStyles.Float,CultureInfo.InvariantCulture,out value)||Double.IsNaN(value)||Double.IsInfinity(value))throw new InvalidOperationException("Scatter and bubble charts need numeric X values in the first column.");return value;}).ToArray();}
            if(type==Office.XlChartType.xlBubble&&(data.Headers.Length!=3||Enumerable.Range(0,data.Rows.Length).Any(r=>!data.Number(r,2).HasValue||data.Number(r,2)<0)))throw new InvalidOperationException("Bubble charts need exactly X, Y and nonnegative Size columns.");
            var shape=Slide.Shapes.AddChart2(-1,type,60,80,600,340,false);
            try {
                LinkedChartData.Write(shape,data,link,rows,xy);
                var chart=shape.Chart;var series=(PowerPoint.SeriesCollection)chart.SeriesCollection();
                if(type==Office.XlChartType.xlPie||type==Office.XlChartType.xlDoughnut||type==Office.XlChartType.xlBubble)while(series.Count>1)series.Item(series.Count).Delete();


                if(combo)series.Item(2).ChartType=Office.XlChartType.xlLineMarkers;
                chart.HasTitle=true;chart.ChartTitle.Text=System.IO.Path.GetFileNameWithoutExtension(link.Path);
                shape.AlternativeText="Chart generated from linked data source. Values remain editable through Edit Data.";
                shape.Select(Office.MsoTriState.msoTrue);
            } catch(Exception original) {try{shape.Delete();}catch(Exception cleanup){throw new InvalidOperationException("Chart insertion failed and cleanup also failed.",new AggregateException(original,cleanup));}throw;}
        }
        private void InsertCombo() {InsertChart(Office.XlChartType.xlColumnClustered,false,true);}
        private void GroupNew(Action create)
        {
            var slide=Slide;int before=slide.Shapes.Count;
            try {
                create();
                var names=new object[slide.Shapes.Count-before];for(int i=0;i<names.Length;i++)names[i]=slide.Shapes[before+i+1].Name;
                if(names.Length>1)slide.Shapes.Range(names).Group();
            } catch {while(slide.Shapes.Count>before)slide.Shapes[slide.Shapes.Count].Delete();throw;}
        }
        private void ProcessFlow() { GroupNew(()=>{for(int i=0;i<4;i++)Box(Office.MsoAutoShapeType.msoShapeChevron,45+i*160,150,150,90,"Step "+(i+1));}); }
        private void Diagram(string family,bool percent=false)
        {
            DataSourceLink link;var data=Source(out link);
            if(family=="gantt"&&data.Headers.Length!=3)throw new InvalidOperationException("Gantt requires Task, numeric Start, positive Duration columns.");
            if(family=="waterfall"&&data.Headers.Length!=2)throw new InvalidOperationException("Waterfall requires Category and signed Change columns; each row is a movement from zero.");
            var totals=new double[data.Rows.Length];double min=0,max=0,end=0;
            for(int r=0;r<data.Rows.Length;r++) {
                for(int c=1;c<data.Headers.Length;c++)if(!data.Number(r,c).HasValue)throw new InvalidOperationException("This chart requires numeric values without blanks.");
                if(family=="gantt"){if(data.Number(r,2)<=0)throw new InvalidOperationException("Duration must be positive.");min=Math.Min(min,data.Number(r,1).Value);max=Math.Max(max,data.Number(r,1).Value+data.Number(r,2).Value);}
                else if(family=="waterfall"){end+=data.Number(r,1).Value;totals[r]=end;min=Math.Min(min,end);max=Math.Max(max,end);}
                else {for(int c=1;c<data.Headers.Length;c++){if(data.Number(r,c)<0)throw new InvalidOperationException("Mekko values must be nonnegative.");totals[r]+=data.Number(r,c).Value;}max=Math.Max(max,totals[r]);}
            }
            double sum=totals.Sum();if(Double.IsInfinity(max-min)||Double.IsNaN(max-min)||Double.IsInfinity(sum))throw new InvalidOperationException("Source numeric range is too large.");
            if(family=="mekko"&&sum<=0)throw new InvalidOperationException("Mekko requires a positive total.");if(max==min){min=-1;max=1;}
            GroupNew(()=>{float x=60;for(int r=0;r<data.Rows.Length;r++) {
                float width=family=="mekko"?(float)(totals[r]/sum*580):500f/data.Rows.Length;
                if(family=="gantt"){float y=90+r*280f/data.Rows.Length;Text(data.Rows[r][0],40,y,120,25,12);Box(Office.MsoAutoShapeType.msoShapeRectangle,170+(float)((data.Number(r,1).Value-min)/(max-min)*470),y,(float)(data.Number(r,2).Value/(max-min)*470),Math.Min(24,240f/data.Rows.Length),data.Rows[r][2]);}
                else if(family=="waterfall"){double start=r==0?0:totals[r-1],value=totals[r];float h=(float)(Math.Abs(value-start)/(max-min)*280);if(h>0)Box(Office.MsoAutoShapeType.msoShapeRectangle,x,(float)(90+(max-Math.Max(value,start))/(max-min)*280),width,h,data.Rows[r][1]);Text(data.Rows[r][0],x,375,width,28,12);x+=580f/data.Rows.Length;}
                else if(totals[r]>0){float height=percent?280:(float)(totals[r]/max*280),y=370-height;for(int c=1;c<data.Headers.Length;c++){float h=(float)(data.Number(r,c).Value/totals[r]*height);if(h>0){var shape=Box(Office.MsoAutoShapeType.msoShapeRectangle,x,y,width,h,data.Headers[c]+": "+data.Rows[r][c]);shape.Fill.ForeColor.RGB=c%2==0?0x704020:0x807000;y+=h;}}Text(data.Rows[r][0],x,375,width,28,12);x+=width;}
            }});
            LinkedChartData.Persist(Slide.Shapes[Slide.Shapes.Count],data,link);
        }
        private void InsertGantt() {Diagram("gantt");}
        private void InsertMekko(bool percent) {Diagram("mekko",percent);}
        private void InsertWaterfall() {Diagram("waterfall");}
        private void Arrange(string mode)
        {
            var shapes=Selected();var first=shapes[0];
            if(mode=="group"){app.ActiveWindow.Selection.ShapeRange.Group();return;}
            if(mode=="ungroup"){app.ActiveWindow.Selection.ShapeRange.Ungroup();return;}
            if(mode=="swap"){if(shapes.Count!=2)throw new InvalidOperationException("Select exactly two objects.");float l=first.Left,t=first.Top;first.Left=shapes[1].Left;first.Top=shapes[1].Top;shapes[1].Left=l;shapes[1].Top=t;return;}
            if(mode=="horizontal" || mode=="vertical")
            {
                bool horizontal=mode=="horizontal";var ordered=shapes.OrderBy(s=>horizontal?s.Left:s.Top).ToList();if(ordered.Count<3)throw new InvalidOperationException("Select at least three objects.");
                float start=horizontal?ordered[0].Left:ordered[0].Top;var last=ordered[ordered.Count-1];float end=horizontal?last.Left+last.Width:last.Top+last.Height;
                float gap=(end-start-ordered.Sum(s=>horizontal?s.Width:s.Height))/(ordered.Count-1);float position=start;
                foreach(var s in ordered){if(horizontal)s.Left=position;else s.Top=position;position+=(horizontal?s.Width:s.Height)+gap;}return;
            }
            foreach(var s in shapes)
            {
                switch(mode){case "left":s.Left=first.Left;break;case "center":s.Left=first.Left+(first.Width-s.Width)/2;break;case "right":s.Left=first.Left+first.Width-s.Width;break;
                case "top":s.Top=first.Top;break;case "middle":s.Top=first.Top+(first.Height-s.Height)/2;break;case "bottom":s.Top=first.Top+first.Height-s.Height;break;
                case "width":s.Width=first.Width;break;case "height":s.Height=first.Height;break;case "size":s.Width=first.Width;s.Height=first.Height;break;
                case "scale-up":s.Width*=1.1f;s.Height*=1.1f;break;case "scale-down":s.Width/=1.1f;s.Height/=1.1f;break;case "rotate":s.Rotation=(s.Rotation+90)%360;break;
                case "flip-horizontal":s.Flip(Office.MsoFlipCmd.msoFlipHorizontal);break;case "flip-vertical":s.Flip(Office.MsoFlipCmd.msoFlipVertical);break;
                case "duplicate":var copy=s.Duplicate();copy.Left+=16;copy.Top+=16;break;case "front":s.ZOrder(Office.MsoZOrderCmd.msoBringToFront);break;case "back":s.ZOrder(Office.MsoZOrderCmd.msoSendToBack);break;}
            }
        }
        private void Grid()
        {
            var shapes=Selected();var points=GridPlanner.Plan(app.ActivePresentation.PageSetup.SlideWidth,app.ActivePresentation.PageSetup.SlideHeight,shapes.Select(s=>s.Width).ToArray(),shapes.Select(s=>s.Height).ToArray());
            var old=shapes.Select(s=>new System.Drawing.PointF(s.Left,s.Top)).ToArray();
            try {for(int i=0;i<shapes.Count;i++){shapes[i].Left=points[i].X;shapes[i].Top=points[i].Y;}}
            catch(Exception original){bool restored=true;for(int i=0;i<shapes.Count;i++)try{shapes[i].Left=old[i].X;shapes[i].Top=old[i].Y;}catch{restored=false;}if(!restored)throw new InvalidOperationException("Grid placement failed and rollback was incomplete. Use PowerPoint Undo.",original);throw;}
        }
        private void InsertTable()
        {
            var link=reportLink??LinkedChartData.Choose(app.ActivePresentation);var data=LinkedChartData.Load(link,false);var cells=LinkedTableProjection.Cells(data);var slide=Slide;float width=app.ActivePresentation.PageSetup.SlideWidth-120,height=Math.Min(app.ActivePresentation.PageSetup.SlideHeight-140,Math.Max(180,cells.GetLength(0)*24));if(width<=100||height<=60)throw new InvalidOperationException("Slide is too small for this table.");PowerPoint.Shape shape=null;app.StartNewUndoEntry();
            try{shape=slide.Shapes.AddTable(cells.GetLength(0),cells.GetLength(1),60,90,width,height);var table=shape.Table;for(int r=0;r<cells.GetLength(0);r++)for(int c=0;c<cells.GetLength(1);c++){var frame=table.Cell(r+1,c+1).Shape.TextFrame;frame.TextRange.Text=cells[r,c];frame.TextRange.Font.Size=Math.Max(8,Math.Min(12,height/cells.GetLength(0)*.6f));frame.TextRange.Font.Bold=r==0?Office.MsoTriState.msoTrue:Office.MsoTriState.msoFalse;frame.VerticalAnchor=Office.MsoVerticalAnchor.msoAnchorMiddle;frame.TextRange.ParagraphFormat.Alignment=PowerPoint.PpParagraphAlignment.ppAlignCenter;}LinkedChartData.Persist(shape,data,link);shape.Tags.Add("PPI_LINKED_ELEMENT","TABLE_V1");shape.Select();}
            catch{if(shape!=null)shape.Delete();throw;}
        }
        private void Agenda(){ChapterAgenda.InsertChapter(app);}
        private void PasteText(){var shapes=Selected();if(shapes.Any(s=>s.HasTextFrame!=Office.MsoTriState.msoTrue))throw new InvalidOperationException("All selected objects must support text.");var input=Prompt("One line per selected object", "");if(input==null)return;var lines=input.Replace("\r","").Split('\n');if(lines.Length!=shapes.Count)throw new InvalidOperationException("Line count must match selected object count.");for(int i=0;i<shapes.Count;i++)shapes[i].TextFrame.TextRange.Text=lines[i];}
        private void Font(bool size){var value=Prompt(size?"Font size (points)":"Font family",size?"20":"Segoe UI");if(value==null)return;float points=0;if(size&&(!Single.TryParse(value,NumberStyles.Float,CultureInfo.InvariantCulture,out points)||Single.IsNaN(points)||Single.IsInfinity(points)||points<6||points>200))throw new InvalidOperationException("Use a font size between 6 and 200.");foreach(var s in Selected())if(s.HasTextFrame==Office.MsoTriState.msoTrue){if(size)s.TextFrame.TextRange.Font.Size=points;else s.TextFrame.TextRange.Font.Name=value;}}
        private void SwitchDecimals(){ProductivityTools.SwitchDecimals(app,false);}
        private void Language(){var input=Prompt("Office language ID","1033");if(input==null)return;int id;if(!Int32.TryParse(input,out id))throw new InvalidOperationException("Enter a numeric Office language ID.");foreach(var s in Selected())if(s.HasTextFrame==Office.MsoTriState.msoTrue)s.TextFrame.TextRange.LanguageID=(Office.MsoLanguageID)id;}
        private void ChartStyle(){var input=Prompt("Native chart style (1–48)","4");if(input==null)return;int style;if(!Int32.TryParse(input,out style)||style<1||style>48)throw new InvalidOperationException("Choose a style from 1 through 48.");Chart().ChartStyle=style;}
        private void ValueLine(){string text=Prompt("Reference value","0");if(text==null)return;double value;if(!Double.TryParse(text,NumberStyles.Float,CultureInfo.InvariantCulture,out value)||Double.IsNaN(value)||Double.IsInfinity(value))throw new InvalidOperationException("Use a finite reference value.");ChartAnnotations.Reference(app,value);}
        private void Difference(){string text=Prompt("Series, from category, to category (1-based)","1,1,2");if(text==null)return;var values=text.Split(',');int series,from,to;if(values.Length!=3||!Int32.TryParse(values[0],out series)||!Int32.TryParse(values[1],out from)||!Int32.TryParse(values[2],out to))throw new InvalidOperationException("Use three whole-number indices.");ChartAnnotations.Difference(app,series,from,to);}
        private void Cagr(){var input=Prompt("Start, end, years (comma separated)","100,121,2");if(input==null)return;var parts=input.Split(',');double start,end,years;if(parts.Length!=3||!Double.TryParse(parts[0],NumberStyles.Float,CultureInfo.InvariantCulture,out start)||!Double.TryParse(parts[1],NumberStyles.Float,CultureInfo.InvariantCulture,out end)||!Double.TryParse(parts[2],NumberStyles.Float,CultureInfo.InvariantCulture,out years)||start<=0||end<=0||years<=0||Double.IsInfinity(start+end+years))throw new InvalidOperationException("Use three finite positive numbers.");double rate=Math.Pow(end/start,1/years)-1;if(Double.IsInfinity(rate)||Double.IsNaN(rate))throw new InvalidOperationException("Calculation is out of range.");Text("CAGR "+rate.ToString("P1",CultureInfo.InvariantCulture),80,100,350,50,24);}
        private void Scan(){int charts=0,pictures=0;foreach(PowerPoint.Slide s in app.ActivePresentation.Slides)foreach(PowerPoint.Shape shape in s.Shapes){if(shape.HasChart==Office.MsoTriState.msoTrue)charts++;if(shape.Type==Office.MsoShapeType.msoPicture)pictures++;}Notify("Native charts: "+charts+"\nPictures: "+pictures,"Presentation scan");}
        private void InsertImage(){var path=OpenFile("Images|*.png;*.jpg;*.jpeg;*.bmp");if(path!=null)Slide.Shapes.AddPicture(path,Office.MsoTriState.msoFalse,Office.MsoTriState.msoTrue,60,80,-1,-1);}
        private void ImportSlides(){var path=OpenFile("PowerPoint|*.pptx");if(path!=null)app.ActivePresentation.Slides.InsertFromFile(path,Slide.SlideIndex);}
        private void ApplyTemplate(){var path=OpenFile("PowerPoint template|*.potx");if(path!=null)app.ActivePresentation.ApplyTemplate(path);}
        private void ExportSlide(){var path=SaveFile("PNG image|*.png","slide.png");if(path!=null)Slide.Export(path,"PNG",1600,900);}
        private void Brand(){int i=0;foreach(var s in Selected()){s.Fill.ForeColor.RGB=i++%2==0?0x704020:0x807000;if(s.HasTextFrame==Office.MsoTriState.msoTrue){s.TextFrame.TextRange.Font.Name="Segoe UI";s.TextFrame.TextRange.Font.Color.RGB=0xFFFFFF;}}}
        private void AxisTitle(){var title=Prompt("Value-axis title","Value");if(title==null)return;var axis=(PowerPoint.Axis)Chart().Axes(PowerPoint.XlAxisType.xlValue,PowerPoint.XlAxisGroup.xlPrimary);axis.HasTitle=true;axis.AxisTitle.Text=title;}
        private void CleanComments()
        {
            int count=0;foreach(PowerPoint.Slide slide in app.ActivePresentation.Slides)count+=slide.Comments.Count;
            if(!Confirm("Remove "+count+" legacy comments from this presentation? Save a copy first if you need to retain them.","Remove comments"))return;
            foreach(PowerPoint.Slide slide in app.ActivePresentation.Slides)for(int i=slide.Comments.Count;i>=1;i--)slide.Comments[i].Delete();
        }
        private void SaveSlides()
        {
            var selection=app.ActiveWindow.Selection;
            if(selection.Type!=PowerPoint.PpSelectionType.ppSelectionSlides)throw new InvalidOperationException("Select slide thumbnails first.");
            var path=SaveFile("PowerPoint presentation|*.pptx","selected-slides.pptx");
            {
                if(path==null)return;
                selection.SlideRange.Copy();PowerPoint.Presentation exported=null;
                try{exported=app.Presentations.Add(Office.MsoTriState.msoTrue);exported.Slides.Paste();exported.SaveAs(path,PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);}
                finally{if(exported!=null){exported.Saved=Office.MsoTriState.msoTrue;exported.Close();}}
            }
        }
        private void RoundData()
        {
            var input=Prompt("Decimal places (0–10)","1");if(input==null)return;int digits;if(!Int32.TryParse(input,out digits)||digits<0||digits>10)throw new InvalidOperationException("Use 0 through 10 decimal places.");
            var chart=Chart();chart.ChartData.Activate();dynamic book=chart.ChartData.Workbook;
            try{dynamic range=book.Worksheets[1].UsedRange;int rows=(int)range.Rows.Count,cols=(int)range.Columns.Count;
                if(rows>1000||cols>1000||(long)rows*cols>1000)throw new InvalidOperationException("The used range exceeds the 1,000-cell demo limit.");
                for(int r=1;r<=rows;r++)for(int c=1;c<=cols;c++){dynamic cell=range.Cells[r,c];object value=(object)cell.Value2;if(!(bool)cell.HasFormula&&value is double)cell.Value2=Math.Round((double)value,digits,MidpointRounding.AwayFromZero);}}
            finally{Host.PowerPointHost.RetryRejectedCall(()=>book.Close(true));}chart.Refresh();
        }
        private void StyleSheet(){var chart=Chart();chart.ChartData.Activate();dynamic book=chart.ChartData.Workbook;try{dynamic header=book.Worksheets[1].UsedRange.Rows[1];header.Font.Bold=true;header.Interior.Color=0xEEDFCC;}finally{Host.PowerPointHost.RetryRejectedCall(()=>book.Close(true));}}
        private void ImportJson()
        {
            var path=OpenFile("S0 chart JSON|*.json");
            {
                if(path==null)return;
                if(new System.IO.FileInfo(path).Length>Host.MetadataCodec.MaxCharacters)throw new InvalidOperationException("JSON file exceeds the demo size limit.");
                Host.ChartDocument model;
                using(var input=System.IO.File.OpenRead(path))model=(Host.ChartDocument)new System.Runtime.Serialization.Json.DataContractJsonSerializer(typeof(Host.ChartDocument)).ReadObject(input);
                Host.MetadataCodec.Validate(model);model.ChartId=Guid.NewGuid().ToString("D");model.Revision=0;
                new Host.PowerPointHost(app).InsertEditableChart(model);
            }
        }
        private void ExportJson()
        {
            var shapes=Selected();if(shapes.Count!=1)throw new InvalidOperationException("Select one managed chart.");
            var host=new Host.PowerPointHost(app);var model=host.ReadChartMetadata(shapes[0].Tags[Host.PowerPointHost.IdTag]);
            var path=SaveFile("S0 chart JSON|*.json","chart.s0.json");if(path!=null)
                using(var output=System.IO.File.Create(path))new System.Runtime.Serialization.Json.DataContractJsonSerializer(typeof(Host.ChartDocument)).WriteObject(output,model);
        }
        private void MiniToolbar()
        {
            var toolbar=new Form{Text="Arrange selected objects",Width=510,Height=135,FormBorderStyle=FormBorderStyle.FixedToolWindow,TopMost=true};
            var buttons=new FlowLayoutPanel{Dock=DockStyle.Fill};
            foreach(var id in new[]{"left","center","top","size","horizontal","vertical"}){var name=id;var button=new Button{Text=name,AutoSize=true};button.Click+=(s,e)=>{try{app.StartNewUndoEntry();Arrange(name);}catch(Exception){MessageBox.Show("Select suitable objects, then try this action again.");}};buttons.Controls.Add(button);}
            toolbar.Controls.Add(buttons);if(inputs.ShowToolbar!=null)inputs.ShowToolbar(toolbar);else toolbar.Show();
        }
        private void TableData()
        {
            var selected=Selected();if(selected.Count!=1||selected[0].HasTable!=Office.MsoTriState.msoTrue)throw new InvalidOperationException("Select one native table.");
            var table=selected[0].Table;int rows=table.Rows.Count,cols=table.Columns.Count;if(rows*cols>400)throw new InvalidOperationException("Demo editor supports at most 400 cells.");
            if(inputs.EditTable!=null)
            {
                var original=new string[rows,cols];for(int r=0;r<rows;r++)for(int c=0;c<cols;c++)original[r,c]=table.Cell(r+1,c+1).Shape.TextFrame.TextRange.Text;
                var edited=inputs.EditTable(original);if(edited==null)return;
                if(edited.GetLength(0)!=rows||edited.GetLength(1)!=cols)throw new InvalidOperationException("Table dimensions cannot change in the datasheet.");
                for(int r=0;r<rows;r++)for(int c=0;c<cols;c++)table.Cell(r+1,c+1).Shape.TextFrame.TextRange.Text=edited[r,c]??"";
                return;
            }
            using(var form=new Form{Text="Table datasheet",Width=700,Height=450,StartPosition=FormStartPosition.CenterScreen})
            {
                var grid=new DataGridView{Dock=DockStyle.Fill,AllowUserToAddRows=false,AllowUserToDeleteRows=false,ColumnCount=cols,RowCount=rows,AccessibleName="Table cells"};
                for(int r=0;r<rows;r++)for(int c=0;c<cols;c++)grid[c,r].Value=table.Cell(r+1,c+1).Shape.TextFrame.TextRange.Text;
                var apply=new Button{Text="Apply changes",Dock=DockStyle.Bottom,Height=36,DialogResult=DialogResult.OK};form.Controls.Add(grid);form.Controls.Add(apply);
                if(form.ShowDialog()==DialogResult.OK){grid.EndEdit();for(int r=0;r<rows;r++)for(int c=0;c<cols;c++)table.Cell(r+1,c+1).Shape.TextFrame.TextRange.Text=Convert.ToString(grid[c,r].Value,CultureInfo.InvariantCulture);}
            }
        }
        private void Library()
        {
            var folder=ChooseFolder();
            {
                if(folder==null)return;
                var files=System.IO.Directory.GetFiles(folder).Where(p=>new[]{".png",".jpg",".jpeg",".pptx",".potx"}.Contains(System.IO.Path.GetExtension(p).ToLowerInvariant())).Take(1000).ToArray();
                if(inputs.ChooseAsset!=null){int chosen=inputs.ChooseAsset(files);if(chosen<0)return;if(chosen>=files.Length)throw new InvalidOperationException("Invalid asset selection.");InsertAsset(files[chosen]);return;}
                using(var form=new Form{Text="Local library",Width=650,Height=430})
                {
                    var list=new ListBox{Dock=DockStyle.Fill};foreach(var file in files)list.Items.Add(System.IO.Path.GetFileName(file));var insert=new Button{Text="Insert / apply selected asset",Dock=DockStyle.Bottom,Height=38,DialogResult=DialogResult.OK};form.Controls.Add(list);form.Controls.Add(insert);
                    if(form.ShowDialog()==DialogResult.OK&&list.SelectedIndex>=0)InsertAsset(files[list.SelectedIndex]);
                }
            }
        }
        private void AgendaCurrent(){ChapterAgenda.InsertContents(app);}
        private void StyleFile()
        {
            var path=OpenFile("Intelligence demo style|*.json");
            {
                if(path==null)return;if(new System.IO.FileInfo(path).Length>8192)throw new InvalidOperationException("Style file is too large.");
                DemoStyle style;using(var input=System.IO.File.OpenRead(path))style=(DemoStyle)new System.Runtime.Serialization.Json.DataContractJsonSerializer(typeof(DemoStyle)).ReadObject(input);
                if(style==null||String.IsNullOrWhiteSpace(style.FontName)||style.FontName.Length>100||style.FontSize<6||style.FontSize>100||style.FillRgb<0||style.FillRgb>0xFFFFFF||style.TextRgb<0||style.TextRgb>0xFFFFFF)throw new InvalidOperationException("Invalid demo style.");
                foreach(var shape in Selected()){shape.Fill.ForeColor.RGB=style.FillRgb;if(shape.HasTextFrame==Office.MsoTriState.msoTrue){shape.TextFrame.TextRange.Font.Name=style.FontName;shape.TextFrame.TextRange.Font.Size=style.FontSize;shape.TextFrame.TextRange.Font.Color.RGB=style.TextRgb;}}
            }
        }
        private void ExcelCagr()
        {
            var input=Prompt("Start, end, years, destination cell","100,121,2,E2");if(input==null)return;var p=input.Split(',');double start,end,years;
            if(p.Length!=4||!Double.TryParse(p[0],NumberStyles.Float,CultureInfo.InvariantCulture,out start)||!Double.TryParse(p[1],NumberStyles.Float,CultureInfo.InvariantCulture,out end)||!Double.TryParse(p[2],NumberStyles.Float,CultureInfo.InvariantCulture,out years)||start<=0||end<=0||years<=0||Double.IsNaN(start+end+years)||Double.IsInfinity(start+end+years)||!System.Text.RegularExpressions.Regex.IsMatch(p[3],"^[A-Z]{1,2}[1-9][0-9]{0,2}$"))throw new InvalidOperationException("Enter positive finite inputs and a cell such as E2.");
            var chart=Chart();chart.ChartData.Activate();dynamic book=chart.ChartData.Workbook;try{dynamic cell=book.Worksheets[1].Range[p[3]];if(cell.Value2!=null&&!Confirm("Replace the existing destination cell?","CAGR formula"))return;cell.Formula="=("+end.ToString("R",CultureInfo.InvariantCulture)+"/"+start.ToString("R",CultureInfo.InvariantCulture)+")^(1/"+years.ToString("R",CultureInfo.InvariantCulture)+")-1";cell.NumberFormat="0.0%";}finally{Host.PowerPointHost.RetryRejectedCall(()=>book.Close(true));}
        }
        internal string Prompt(string title,string initial)
        {
            if(inputs.Prompt!=null)return inputs.Prompt(title,initial);
            using(var form=new Form{Text=title,Width=460,Height=240,StartPosition=FormStartPosition.CenterScreen,MinimizeBox=false,MaximizeBox=false})
            {
                var text=new TextBox{Multiline=true,Dock=DockStyle.Fill,Text=initial,AccessibleName=title};var buttons=new FlowLayoutPanel{Dock=DockStyle.Bottom,Height=45,FlowDirection=FlowDirection.RightToLeft};
                var ok=new Button{Text="Apply",DialogResult=DialogResult.OK};var cancel=new Button{Text="Cancel",DialogResult=DialogResult.Cancel};buttons.Controls.Add(ok);buttons.Controls.Add(cancel);form.Controls.Add(text);form.Controls.Add(buttons);form.AcceptButton=ok;form.CancelButton=cancel;
                return form.ShowDialog()==DialogResult.OK?text.Text:null;
            }
        }
        private string OpenFile(string filter)
        {
            if(inputs.OpenFile!=null)return inputs.OpenFile(filter);
            using(var dialog=new OpenFileDialog{Filter=filter})return dialog.ShowDialog()==DialogResult.OK?dialog.FileName:null;
        }
        private string SaveFile(string filter,string initial)
        {
            if(inputs.SaveFile!=null)return inputs.SaveFile(filter,initial);
            using(var dialog=new SaveFileDialog{Filter=filter,FileName=initial})return dialog.ShowDialog()==DialogResult.OK?dialog.FileName:null;
        }
        private string ChooseFolder()
        {
            if(inputs.Folder!=null)return inputs.Folder();
            using(var dialog=new FolderBrowserDialog{Description="Choose a local asset folder"})return dialog.ShowDialog()==DialogResult.OK?dialog.SelectedPath:null;
        }
        private bool Confirm(string message,string title)
        {return inputs.Confirm!=null?inputs.Confirm(message,title):MessageBox.Show(message,title,MessageBoxButtons.OKCancel)==DialogResult.OK;}
        private void Notify(string message,string title)
        {if(inputs.Notify!=null)inputs.Notify(message,title);else MessageBox.Show(message,title);}
        private void InsertAsset(string path)
        {
            var ext=System.IO.Path.GetExtension(path).ToLowerInvariant();
            if(ext==".pptx")app.ActivePresentation.Slides.InsertFromFile(path,Slide.SlideIndex);
            else if(ext==".potx")app.ActivePresentation.ApplyTemplate(path);
            else Slide.Shapes.AddPicture(path,Office.MsoTriState.msoFalse,Office.MsoTriState.msoTrue,60,80,-1,-1);
        }
    }
    public sealed class DemoStyle
    {
        public string FontName { get; set; }
        public int FontSize { get; set; }
        public int FillRgb { get; set; }
        public int TextRgb { get; set; }
    }
}
