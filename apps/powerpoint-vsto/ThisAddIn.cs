using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using PowerPointIntelligence.Compatibility;
using PowerPointIntelligence.Host;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointIntelligence.AddIn
{
    public partial class ThisAddIn
    {
        private PowerPointHost host;
        private Label inspector;
        private Microsoft.Office.Tools.CustomTaskPane pane;
        private bool busy;
        private ContextPropertiesController contextProperties;
        private Timer layoutTimer;
        private Timer agendaTimer;
        private Form layoutControls;
        private int paneWindow;
        private bool paneNeedsRefresh;
        private bool compatibilityAllowed;
        private string compatibilityStatus = "Office compatibility has not been checked.";
        private IntelligenceRibbon ribbon;
        private DemoAutomation automation;
        protected override object RequestComAddInAutomationService()
        { return automation ?? (automation = new DemoAutomation(Application, () => ribbon != null && ribbon.IsLoaded,()=>contextProperties==null?"<properties visible=\"false\" kind=\"None\" />":contextProperties.AutomationState())); }
        internal bool CommandsEnabled { get { return host != null && compatibilityAllowed; } }
        internal bool DataSourceCommandsEnabled
        {
            get { try { return CommandsEnabled && Application.Presentations.Count>0; } catch(COMException) { return false; } }
        }

        private void OnStartup(object sender, EventArgs e)
        {
            host = new PowerPointHost(Application);
            layoutTimer=new Timer {Interval=750};layoutTimer.Tick+=(timerSender,timerEvent)=>RunHostEvent(()=>{try{LibraryLayoutLinks.Reflow(Application);}catch(Exception ex) when(ex is InvalidOperationException||ex is InvalidDataException||ex is COMException){layoutTimer.Stop();if(inspector!=null)inspector.Text="Layout links paused: "+ex.Message+" Open Layout Links to resume.";}});layoutTimer.Start();
            agendaTimer=new Timer {Interval=650};agendaTimer.Tick+=(timerSender,timerEvent)=>{agendaTimer.Stop();RunHostEvent(()=>{try{ChapterAgenda.Automatic(Application);if(Application.Presentations.Count>0&&Application.ActiveWindow.ViewType==PowerPoint.PpViewType.ppViewNormal){TextShapeElements.AutoReflow(Application);UniversalConnectors.Reflow(Application);}}catch(Exception error) when(error is InvalidOperationException||error is InvalidDataException||error is COMException){if(inspector!=null)inspector.Text="Managed element update paused: "+error.Message+" Open Manage Agenda or Update agenda after correcting the issue.";}});};
            ConfigureCompatibility();
            contextProperties=new ContextPropertiesController(Application,()=>host!=null&&compatibilityAllowed&&!busy,action=>Run(action),(key,id)=>Run(()=>ContextActionRouter.Run(Application,key,id)));
            Application.WindowSelectionChange += SelectionChanged;
            Application.WindowBeforeDoubleClick += ElementDoubleClick;
            Application.PresentationOpen += Opened;
            Application.PresentationBeforeSave += BeforeSave;
            Application.WindowActivate += Activated;
            Application.PresentationClose += Closed;
            if (ribbon != null) ribbon.Refresh();
        }
        private void ConfigureCompatibility()
        {
            try
            {
                var settings = OfficeSettings.Load(typeof(ThisAddIn).Assembly.Location + ".config");
                string build = null;
                try
                {
                    var info = FileVersionInfo.GetVersionInfo(Path.Combine(Application.Path, "POWERPNT.EXE"));
                    if (info.FileMajorPart > 0)
                        build = new Version(info.FileMajorPart, info.FileMinorPart, info.FileBuildPart, info.FilePrivatePart).ToString();
                }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
                var result = OfficeCompatibility.Evaluate(settings, Application.Version, build, Environment.Is64BitProcess);
                compatibilityAllowed = result.Allowed;
                compatibilityStatus = result.Message;
            }
            catch (Exception ex) when (ex is FormatException || ex is System.Configuration.ConfigurationErrorsException ||
                ex is IOException || ex is UnauthorizedAccessException || ex is COMException)
            {
                compatibilityAllowed = false;
                compatibilityStatus = "Office compatibility configuration is invalid or unavailable. Check the add-in .dll.config file. Chart commands are disabled.";
            }
        }
        private void OnShutdown(object sender, EventArgs e)
        {
            if(contextProperties!=null)contextProperties.Dispose();if(agendaTimer!=null)agendaTimer.Dispose();if(layoutTimer!=null)layoutTimer.Dispose();if(layoutControls!=null)layoutControls.Dispose();
            Application.WindowSelectionChange -= SelectionChanged;
            Application.WindowBeforeDoubleClick -= ElementDoubleClick;
            Application.PresentationOpen -= Opened;
            Application.PresentationBeforeSave -= BeforeSave;
            Application.WindowActivate -= Activated;
            Application.PresentationClose -= Closed;
            RemovePane();
        }
        private void EnsurePane()
        {
            int currentWindow=Application.ActiveWindow.HWND;
            if (pane != null && !paneNeedsRefresh && paneWindow==currentWindow) return;
            if (pane != null) RemovePane();
            var control = new UserControl { Dock = DockStyle.Fill, AutoScaleMode = AutoScaleMode.Font,
                Font = new System.Drawing.Font("Segoe UI", 10), BackColor = System.Drawing.Color.White };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Padding = new Padding(16) };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 76));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 120));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.Controls.Add(new Label { Text = "Chart workspace", Dock = DockStyle.Fill,
                Font = new System.Drawing.Font("Segoe UI", 15, System.Drawing.FontStyle.Bold) }, 0, 0);
            layout.Controls.Add(new Label { Text = "Insert an editable chart, change its data in PowerPoint, then check and sync your changes here.", Dock = DockStyle.Fill }, 0, 1);
            var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true };
            var insert = new Button { Text = "Insert linked column chart", AutoSize = true, AccessibleName = "Insert linked clustered column chart" };
            var inspect = new Button { Text = "Check and sync charts", AutoSize = true, AccessibleName = "Check and sync charts" };
            var compatibility = new Button { Text = "Office compatibility", AutoSize = true, AccessibleName = "Office compatibility" };
            insert.Click += (sender, args) => Insert();
            inspect.Click += (sender, args) => Inspect();
            compatibility.Click += (sender, args) => ShowCompatibility();
            insert.Enabled = inspect.Enabled = compatibilityAllowed;
            actions.Controls.Add(insert); actions.Controls.Add(inspect); actions.Controls.Add(compatibility);
            layout.Controls.Add(actions, 0, 2);
            inspector = new Label { Dock = DockStyle.Fill, AutoSize = false, Padding = new Padding(0, 12, 0, 0), Text = "Select a chart to see its status.", AccessibleName = "Chart status" };
            layout.Controls.Add(inspector, 0, 3);
            control.Controls.Add(layout);
            pane = CustomTaskPanes.Add(control, "PowerPoint Intelligence", Application.ActiveWindow);
            paneWindow=currentWindow;paneNeedsRefresh=false;
            pane.Width = 330;
            pane.Visible = true;
            host.Status = text => inspector.Text = compatibilityStatus + Environment.NewLine + Environment.NewLine + text;
        }
        private void RemovePane()
        {
            if (pane == null) return;
            var previous=pane;
            pane = null; inspector = null;
            paneWindow=0;paneNeedsRefresh=false;
            if (host != null) host.Status = delegate { };
            try { CustomTaskPanes.Remove(previous); }
            catch (COMException) { }
            catch (ObjectDisposedException) { }
        }
        private void Run(Action action)
        {
            if (busy || host == null) return;
            busy = true;
            try
            {
                using(var lease=PowerPointIntelligence.Host.OfficeMutationLease.TryEnter(System.Diagnostics.Process.GetCurrentProcess().Id)){
                    if(lease==null)return;
                    EnsurePane();
                    if (!compatibilityAllowed) { inspector.Text = compatibilityStatus; return; }
                    action();
                }
            }
            catch (OperationCanceledException) { }
            catch (InvalidOperationException ex) { if(inspector!=null)inspector.Text=ex.Message; }
            catch (InvalidDataException ex) { if(inspector!=null)inspector.Text=ex.Message; }
            catch (Exception) { if (inspector != null) inspector.Text = "Operation unavailable. Check slide selection, document permissions and metadata. Native objects were not replaced during inspection."; }
            finally { busy = false; if(contextProperties!=null)contextProperties.Request(); }
        }
        internal void InsertRenderedShapes()
        {
            Run(() => {
                var link=LinkedChartData.Choose(Application.ActivePresentation);var table=LinkedChartData.Load(link);
                var model=new PowerPointIntelligence.Domain.ChartDocument {
                    SchemaVersion=2,ChartId=Guid.NewGuid().ToString("D"),Revision=0,ChartType="clustered-column",
                    Dataset=new PowerPointIntelligence.Domain.EmbeddedDataset {Id="dataset:linked",Categories=new System.Collections.Generic.List<PowerPointIntelligence.Domain.ChartCategory>()},
                    Series=new System.Collections.Generic.List<PowerPointIntelligence.Domain.ChartSeries>(),
                    Axes=new System.Collections.Generic.List<PowerPointIntelligence.Domain.ChartAxis> {
                        new PowerPointIntelligence.Domain.ChartAxis {Id="axis:category",Role="category",Scale="category",NumberFormat="General"},
                        new PowerPointIntelligence.Domain.ChartAxis {Id="axis:value",Role="value",Scale="linear",NumberFormat="General"}},
                    Layout=new PowerPointIntelligence.Domain.ChartLayout {Unit="pt",X=60,Y=90,Width=600,Height=340},
                    Style=new PowerPointIntelligence.Domain.ChartStyle {ShowLegend=true,SeriesColors=new System.Collections.Generic.List<string>()}
                };
                for(int r=0;r<table.Rows.Length;r++)model.Dataset.Categories.Add(new PowerPointIntelligence.Domain.ChartCategory {Id="category:"+r,Label=table.Rows[r][0]});
                string[] colors={"#4472C4","#ED7D31","#70AD47","#A5A5A5","#FFC000","#5B9BD5"};
                for(int c=1;c<table.Headers.Length;c++) {
                    var series=new PowerPointIntelligence.Domain.ChartSeries {Id="series:"+c,Name=table.Headers[c],Values=new System.Collections.Generic.List<double?>()};
                    for(int r=0;r<table.Rows.Length;r++)series.Values.Add(table.Number(r,c));model.Series.Add(series);model.Style.SeriesColors.Add(colors[(c-1)%colors.Length]);
                }
                RendererRibbonCommand.Insert(Application,model);
            });
        }
        internal void Insert() { Run(() => new DemoTools(Application).Execute(Array.Find(DemoTools.Catalog,t=>t.Id=="column"))); }
        internal void LinkDataSource(string format)
        {
            if(!DataSourceCommandsEnabled)return;
            Run(() => {
                try
                {
                    var target=Application.ActivePresentation;
                    var linked=DataSources.ChooseAndLink(format,DataSources.PickFile,link=>DataSources.Save(target,link));
                    if(linked!=null)inspector.Text="Linked "+DataSources.Option(linked.Format).Label+": "+Path.GetFileName(linked.Path)+Environment.NewLine+"File reference saved with this presentation.";
                }
                catch(Exception ex) when(ex is InvalidDataException || ex is IOException || ex is UnauthorizedAccessException || ex is ArgumentException)
                { inspector.Text="Could not link the data source. Choose an accessible file matching the selected format."; }
            });
        }
        internal void ChangeTemplatePalette() { Run(() => {
            var slide=Application.ActiveWindow.View.Slide as PowerPoint.Slide;
            if(slide==null||String.IsNullOrEmpty(slide.Tags[TemplateInserter.TemplateTag]))throw new InvalidOperationException("Select a library template slide first.");
            string palette=PalettePicker.Choose(slide.Tags["PPI_TEMPLATE_PALETTE"]);if(palette!=null){TemplateInserter.ChangePalette(Application,palette);inspector.Text="Template palette changed to "+TemplatePalettes.Get(palette).Name+".";}
        }); }
        internal void ChangeTemplateAlignment() { Run(() => {var choice=AlignmentPicker.Choose();if(choice!=null)TemplateInserter.ChangeAlignment(Application,choice[0],choice[1]);}); }
        internal void OpenTemplateLibrary() { Run(() => {using(var library=new LocalLibraryForm(Application,null,true))library.ShowDialog();}); }
        internal void OpenDemo()
        {
            Run(() => {
                RemovePane();
                var workspace = new DemoWorkspace(tool => Run(() => {
                    new DemoTools(Application).Execute(tool);
                    inspector.Text = "Completed: " + tool.Name + ". Use PowerPoint Undo to revert slide edits.";
                }));
                inspector = workspace.Status;
                pane = CustomTaskPanes.Add(workspace, "Intelligence • Demonstration", Application.ActiveWindow);
                paneWindow=Application.ActiveWindow.HWND;paneNeedsRefresh=false;
                pane.Width = 420; pane.Visible = true;
                host.Status = text => inspector.Text = text;
            });
        }
        internal void ShowContextProperties(){if(contextProperties!=null)contextProperties.ShowCurrent();}
        internal void RunProductivity(string id)
        {
            if(id=="data-links"){RunWorkflow("data-manager");return;}
            if(id=="review"){RunWorkflow("quality-review");return;}
            if(id=="save-selected"){RunDemo("save-slides");return;}
            if(id=="selected-style"){RunDemo("style-file");return;}
            Run(()=>{if(id=="help")MessageBox.Show("Element Controls provides native editable text shapes, process grids, universal connectors and vector indicators. Select an element and choose Edit selected element. Double-click our indicator group to advance; Indicator keyboard controls provides focused shortcuts. Process rows/columns are hidden with retained history. Tools offers selected-text decimal replacement, symbols, picture copies and chart-specific workbooks. Shared slide workbooks, external indicator links and proprietary object imports are not implemented. Native Office acceptance is pending.","PowerPoint Intelligence — Elements");else ProductivityTools.Run(Application,id);});
        }
        internal bool ElementEnabled(string id){return DataSourceCommandsEnabled&&ElementCommands.Enabled(Application,id);}
        internal void RunElement(string id){Run(()=>{ElementCommands.Run(Application,id);inspector.Text="Completed element command: "+id;});}
        private void ElementDoubleClick(PowerPoint.Selection selection,ref bool cancel)
        {
            if(!DataSourceCommandsEnabled||busy||selection.Type!=PowerPoint.PpSelectionType.ppSelectionShapes||selection.ShapeRange.Count!=1||!IndicatorElements.IsIndicator(selection))return;
            // Only our managed indicator receives this event. Native text editing and other add-ins remain untouched.
            cancel=true;Run(()=>IndicatorElements.CycleSelected(Application));
        }
        internal bool WorkflowEnabled(string id)
        {
            if(!DataSourceCommandsEnabled)return false;
            try{
                if(id.StartsWith("text-",StringComparison.Ordinal)||id.StartsWith("color-",StringComparison.Ordinal))return NativeFormatting.CanApply(Application);
                var selection=Application.ActiveWindow.Selection;
                if(id=="axis-settings"||id=="relink-selected"||id=="series-settings")return selection.Type==PowerPoint.PpSelectionType.ppSelectionShapes&&selection.ShapeRange.Count==1&&selection.ShapeRange[1].HasChart==Office.MsoTriState.msoTrue;
                if(id=="annotation-refresh")return selection.Type==PowerPoint.PpSelectionType.ppSelectionShapes&&selection.ShapeRange.Count==1&&selection.ShapeRange[1].Tags["PPI_CA_SCHEMA"]=="1";
                if(id=="match-scales"){if(selection.Type!=PowerPoint.PpSelectionType.ppSelectionShapes||selection.ShapeRange.Count<2||selection.ShapeRange.Count>30)return false;foreach(PowerPoint.Shape shape in selection.ShapeRange)if(shape.HasChart!=Office.MsoTriState.msoTrue)return false;}
                if(id=="agenda-refresh"||id=="agenda-settings")return DataSourceCommandsEnabled;
                return true;
            }catch(COMException){return false;}catch(InvalidOperationException){return false;}
        }
        internal void RunWorkflow(string id)
        {
            Run(()=>{
                switch(id){
                    case "agenda-insert": ChapterAgenda.InsertChapter(Application);break;
                    case "agenda-contents": ChapterAgenda.InsertContents(Application);break;
                    case "agenda-settings": ChapterAgenda.ShowManager(Application);break;
                    case "agenda-split": ChapterAgenda.SplitHere(Application);break;
                    case "agenda-refresh": var agendaSlide=Application.ActiveWindow.View.Slide as PowerPoint.Slide;if(agendaSlide!=null&&agendaSlide.Tags[PresentationAgenda.SchemaTag]=="1")PresentationAgenda.RefreshSelected(Application);else if(!ChapterAgenda.RefreshAll(Application))throw new InvalidOperationException("Insert a managed chapter or contents slide first.");break;
                    case "annotation-refresh": ChartAnnotations.RefreshSelected(Application);break;
                    case "connections": ServiceConnections.ShowConfiguration();break;
                    case "series-settings": NativeSeriesControls.Show(Application);break;
                    case "export-pdf": NativePdfExport.Show(Application);break;
                    case "axis-settings": NativeChartControls.ShowAxis(Application);break;
                    case "match-scales": NativeChartControls.MatchScales(Application);break;
                    case "map-excel": MappedExcelData.Choose(Application);break;
                    case "relink-selected": MappedLinkOperations.RelinkSelected(Application);break;
                    case "data-manager": using(var manager=new DataLinkManager(Application,(shape,link)=>{Application.StartNewUndoEntry();MappedLinkOperations.Refresh(shape,link);}))manager.ShowDialog();break;
                    case "percent-area": new DemoTools(Application).InsertChart(Office.XlChartType.xlAreaStacked100);break;
                    case "quality-review": PresentationQuality.ShowInspection(Application);break;
                    default:
                        if(id.StartsWith("text-",StringComparison.Ordinal)){NativeFormatting.Text(Application,id);break;}
                        if(id.StartsWith("color-",StringComparison.Ordinal)){if(!NativeFormatting.Color(Application,id))return;break;}
                        throw new InvalidOperationException("Unknown workflow command.");
                }
                inspector.Text="Completed workflow command: "+id+". Use PowerPoint Undo to revert document edits.";
            });
        }
        internal void RunDemo(string id)
        {
            var tool = Array.Find(DemoTools.Catalog, item => item.Id == id);
            if (tool == null) return;
            Run(() => { new DemoTools(Application).Execute(tool); inspector.Text = "Completed: " + tool.Name; });
        }
        internal void OpenLayoutControls(){Run(()=>{layoutTimer.Start();if(layoutControls==null||layoutControls.IsDisposed)layoutControls=LibraryLayoutLinks.Controls(Application);layoutControls.Show();layoutControls.BringToFront();});}
        internal void OpenWorkspace() { Run(() => { inspector.Text = compatibilityStatus + Environment.NewLine + Environment.NewLine + "Link a Data Source, then choose a chart to begin."; ShowSelection(); }); }
        internal void Inspect() { Run(() => { host.InspectAndRecover(true); ShowSelection(); }); }
        internal void ShowCompatibility()
        {
            if (host == null || busy) return;
            try { EnsurePane(); inspector.Text = compatibilityStatus; }
            catch (COMException) { } // No document/window; try again after opening a presentation.
            catch (InvalidOperationException) { }
        }
        private void ShowSelection()
        {
            if (inspector == null) return;
            var ids = host.GetSelectedObjectIds();
            if (ids.Count != 1) return;
            var document = host.ReadChartMetadata(ids[0]);
            inspector.Text += Environment.NewLine + "Selected chart: " + document.ChartId +
                Environment.NewLine + "Revision: " + document.Revision +
                Environment.NewLine + "Categories: " + document.Categories.Count + "; series: " + document.Series.Count;
        }
        private void SelectionChanged(PowerPoint.Selection selection)
        {
            if(contextProperties!=null)contextProperties.Request();
            RunHostEvent(()=>{if(ribbon!=null)ribbon.Refresh();ShowSelection();if(agendaTimer!=null){agendaTimer.Stop();agendaTimer.Start();}});
        }
        private void RunHostEvent(Action action)
        {
            if (busy || host==null || !compatibilityAllowed) return;
            busy=true;
            try { OfficeEventBoundary.Run(()=>{using(var lease=PowerPointIntelligence.Host.OfficeMutationLease.TryEnter(System.Diagnostics.Process.GetCurrentProcess().Id)){if(lease!=null)action();}}, message => Trace.WriteLine(message)); }
            finally { busy=false; }
        }
        private void Opened(PowerPoint.Presentation presentation)
        { RunHostEvent(() => { paneNeedsRefresh=true; if(ribbon!=null)ribbon.Refresh(); }); }
        private void Activated(PowerPoint.Presentation presentation, PowerPoint.DocumentWindow window)
        {
            // Do not create/dispose docked controls or rewrite metadata from
            // Office's window activation callback. Rebind on the next user action.
            RunHostEvent(() => { if (paneWindow!=window.HWND) paneNeedsRefresh=true; if(ribbon!=null)ribbon.Refresh(); });
        }
        private void Closed(PowerPoint.Presentation presentation)
        { paneNeedsRefresh=true;if(contextProperties!=null)contextProperties.Hide(); RunHostEvent(() => { if(ribbon!=null)ribbon.Refresh(); }); }
        private void BeforeSave(PowerPoint.Presentation presentation, ref bool cancel)
        {
            // Only operate on the active window's presentation. Never inspect the wrong deck.
            RunHostEvent(() => {
                if (Application.Presentations.Count > 0 && presentation == Application.ActivePresentation)
                { if(Application.ActivePresentation.CustomXMLParts.SelectByNamespace(ChapterAgendaCodec.Namespace).Count>0)ChapterAgenda.RefreshAll(Application,false,true);LibraryLayoutLinks.Reflow(Application);host.InspectAndRecover(true); }
            });
        }
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        { return ribbon ?? (ribbon = new IntelligenceRibbon(this)); }
        private void InternalStartup() { Startup += OnStartup; Shutdown += OnShutdown; }
    }

    [ComVisible(true)]
    public sealed class IntelligenceRibbon : Office.IRibbonExtensibility
    {
        private readonly ThisAddIn addin;
        private Office.IRibbonUI ribbonUI;
        internal bool IsLoaded { get { return ribbonUI != null; } }
        public IntelligenceRibbon(ThisAddIn addin) { this.addin = addin; }
        public string GetCustomUI(string ribbonID) { return DemoRibbon.Create(); }
        public object LoadImage(string imageId) { return RibbonIcons.Load(imageId); }
        public void InsertRenderedShapes(Office.IRibbonControl control) { addin.InsertRenderedShapes(); }
        public void InsertSample(Office.IRibbonControl control) { addin.Insert(); }
        public void LinkDataSource(Office.IRibbonControl control) { addin.LinkDataSource(control.Tag); }
        public bool GetDataSourceEnabled(Office.IRibbonControl control) { return addin.DataSourceCommandsEnabled; }
        public void InspectChart(Office.IRibbonControl control) { addin.Inspect(); }
        public void ShowCompatibility(Office.IRibbonControl control) { addin.ShowCompatibility(); }
        public void ShowWorkspace(Office.IRibbonControl control) { addin.OpenWorkspace(); }
        public void ChangeTemplatePalette(Office.IRibbonControl control) { addin.ChangeTemplatePalette(); }
        public void ChangeTemplateAlignment(Office.IRibbonControl control) { addin.ChangeTemplateAlignment(); }
        public void ShowLibraryLayout(Office.IRibbonControl control) { addin.OpenLayoutControls(); }
        public void ShowTemplateLibrary(Office.IRibbonControl control) { addin.OpenTemplateLibrary(); }
        public void ShowContextProperties(Office.IRibbonControl control) { addin.ShowContextProperties(); }
        public void ShowDemo(Office.IRibbonControl control) { addin.OpenDemo(); }
        public bool GetWorkflowEnabled(Office.IRibbonControl control) { return addin.WorkflowEnabled(control.Tag); }
        public void RunWorkflowCommand(Office.IRibbonControl control) { addin.RunWorkflow(control.Tag); }
        public void RunProductivityTool(Office.IRibbonControl control) { addin.RunProductivity(control.Tag); }
        public bool GetElementEnabled(Office.IRibbonControl control) { return addin.ElementEnabled(control.Tag); }
        public void RunElementCommand(Office.IRibbonControl control) { addin.RunElement(control.Tag); }
        public void RunDemoTool(Office.IRibbonControl control) { addin.RunDemo(control.Tag); }
        public bool GetCommandEnabled(Office.IRibbonControl control) { return addin.CommandsEnabled; }
        public void OnLoad(Office.IRibbonUI ui) { ribbonUI = ui; }
        internal void Refresh() { if (ribbonUI != null) ribbonUI.Invalidate(); }
    }
}
