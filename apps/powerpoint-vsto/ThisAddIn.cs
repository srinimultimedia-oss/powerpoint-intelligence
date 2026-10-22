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
        private bool compatibilityAllowed;
        private string compatibilityStatus = "Office compatibility has not been checked.";
        private IntelligenceRibbon ribbon;
        internal bool CommandsEnabled { get { return host != null && compatibilityAllowed; } }

        private void OnStartup(object sender, EventArgs e)
        {
            host = new PowerPointHost(Application);
            ConfigureCompatibility();
            Application.WindowSelectionChange += SelectionChanged;
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
            Application.WindowSelectionChange -= SelectionChanged;
            Application.PresentationOpen -= Opened;
            Application.PresentationBeforeSave -= BeforeSave;
            Application.WindowActivate -= Activated;
            Application.PresentationClose -= Closed;
            RemovePane();
        }
        private void EnsurePane()
        {
            if (pane != null) return;
            var control = new UserControl { Dock = DockStyle.Fill };
            inspector = new Label { Dock = DockStyle.Fill, AutoSize = false, Padding = new Padding(12), Text = "Select an Intelligence chart." };
            control.Controls.Add(inspector);
            pane = CustomTaskPanes.Add(control, "PowerPoint Intelligence", Application.ActiveWindow);
            pane.Width = 330;
            pane.Visible = true;
            host.Status = text => inspector.Text = compatibilityStatus + Environment.NewLine + Environment.NewLine + text;
        }
        private void RemovePane()
        {
            if (pane == null) return;
            CustomTaskPanes.Remove(pane);
            pane = null; inspector = null;
            if (host != null) host.Status = delegate { };
        }
        private void Run(Action action)
        {
            if (busy || host == null) return;
            busy = true;
            try
            {
                EnsurePane();
                if (!compatibilityAllowed) { inspector.Text = compatibilityStatus; return; }
                action();
            }
            catch (Exception) { if (inspector != null) inspector.Text = "Operation unavailable. Check slide selection, document permissions and metadata. Native objects were not replaced during inspection."; }
            finally { busy = false; }
        }
        internal void Insert() { Run(() => host.InsertEditableChart(MetadataCodec.Sample())); }
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
            var ids = host.GetSelectedObjectIds();
            if (ids.Count != 1) return;
            var document = host.ReadChartMetadata(ids[0]);
            inspector.Text += Environment.NewLine + "Selected chart: " + document.ChartId +
                Environment.NewLine + "Revision: " + document.Revision +
                Environment.NewLine + "Categories: " + document.Categories.Count + "; series: " + document.Series.Count;
        }
        private void SelectionChanged(PowerPoint.Selection selection)
        {
            Run(() => { host.InspectAndRecover(false); ShowSelection(); });
        }
        private void Opened(PowerPoint.Presentation presentation) { Run(() => host.InspectAndRecover(false)); }
        private void Activated(PowerPoint.Presentation presentation, PowerPoint.DocumentWindow window)
        { if (busy) return; RemovePane(); Run(() => host.InspectAndRecover(false)); }
        private void Closed(PowerPoint.Presentation presentation) { RemovePane(); }
        private void BeforeSave(PowerPoint.Presentation presentation, ref bool cancel)
        {
            // Only operate on the active window's presentation. Never inspect the wrong deck.
            if (Application.Presentations.Count > 0 && presentation == Application.ActivePresentation)
                Run(() => host.InspectAndRecover(true));
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
        public IntelligenceRibbon(ThisAddIn addin) { this.addin = addin; }
        public string GetCustomUI(string ribbonID)
        {
            return "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnLoad'><ribbon><tabs>" +
                "<tab id='ppiTab' label='Intelligence'><group id='ppiS0' label='S0 Feasibility'>" +
                "<button id='ppiInsert' label='Insert Sample Chart' size='large' onAction='InsertSample' getEnabled='GetCommandEnabled'/>" +
                "<button id='ppiInspect' label='Inspect / Recover' onAction='InspectChart' getEnabled='GetCommandEnabled'/>" +
                "<button id='ppiCompatibility' label='Office Compatibility' onAction='ShowCompatibility'/>" +
                "</group></tab></tabs></ribbon></customUI>";
        }
        public void InsertSample(Office.IRibbonControl control) { addin.Insert(); }
        public void InspectChart(Office.IRibbonControl control) { addin.Inspect(); }
        public void ShowCompatibility(Office.IRibbonControl control) { addin.ShowCompatibility(); }
        public bool GetCommandEnabled(Office.IRibbonControl control) { return addin.CommandsEnabled; }
        public void OnLoad(Office.IRibbonUI ui) { ribbonUI = ui; }
        internal void Refresh() { if (ribbonUI != null) ribbonUI.Invalidate(); }
    }
}
