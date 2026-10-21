// VSTO host bootstrap. Keep host-item name in sync with ThisAddIn.Designer.xml.
using System;
using Microsoft.Office.Tools;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointIntelligence.AddIn
{
    [Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObject(0)]
    public sealed partial class ThisAddIn : AddInBase
    {
        internal PowerPoint.Application Application;
        internal CustomTaskPaneCollection CustomTaskPanes;
        private static Microsoft.Office.Tools.Factory factory;

        public ThisAddIn(Microsoft.Office.Tools.Factory factory, IServiceProvider serviceProvider)
            : base(factory, serviceProvider, "AddIn", "ThisAddIn")
        { ThisAddIn.factory = factory; }

        protected override void Initialize()
        {
            base.Initialize();
            Application = GetHostItem<PowerPoint.Application>(typeof(PowerPoint.Application), "Application");
            CustomTaskPanes = factory.CreateCustomTaskPaneCollection(null, null, "CustomTaskPanes", "CustomTaskPanes", this);
            InternalStartup();
        }
        protected override void FinishInitialization()
        { OnStartup(); }
        protected override void InitializeDataBindings() { BeginInitialization(); EndInitialization(); }
        protected override void OnShutdown()
        {
            base.OnShutdown();
            if (CustomTaskPanes != null) CustomTaskPanes.Dispose();
        }
    }
}
