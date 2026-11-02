using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointIntelligence.AddIn
{
    [ComVisible(true), Guid("EF4062B2-5E48-4DA0-8904-1AB37241ADCB"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IDemoAutomation
    {
        string GetVersion();
        string GetBuildInfo();
        string GetRibbonXml();
        bool GetRibbonLoaded();
        string InsertChartJson(string json);
        string GetContextPropertiesState();
    }
    [ComVisible(true), Guid("D3E4D258-43A9-4981-824C-28F4CA20675C"), ClassInterface(ClassInterfaceType.None)]
    public sealed class DemoAutomation : IDemoAutomation
    {
        private readonly PowerPoint.Application app;
        private readonly int threadId;
        private readonly Func<bool> ribbonLoaded;
        private readonly Func<string> contextProperties;
        internal DemoAutomation(PowerPoint.Application app,Func<bool> ribbonLoaded=null,Func<string> contextProperties=null){this.contextProperties=contextProperties??(()=>"<properties visible=\"false\" kind=\"None\" />");this.app=app;this.ribbonLoaded=ribbonLoaded??(()=>false);threadId=Thread.CurrentThread.ManagedThreadId;}
        public string GetVersion(){return "demo-v1 / S0 ChartDocument JSON";}
        public string GetBuildInfo(){return DemoRibbon.BuildMarker+"; tools="+DemoTools.Catalog.Length+"; mvid="+typeof(DemoAutomation).Module.ModuleVersionId+"; assembly="+typeof(DemoAutomation).Assembly.Location;}
        public string GetRibbonXml(){return DemoRibbon.Create();}
        public bool GetRibbonLoaded(){return ribbonLoaded();}
        // Read-only observation is an immutable UI-published snapshot; no Office/WinForms calls occur here.
        public string GetContextPropertiesState(){return contextProperties();}
        public string InsertChartJson(string json)
        {
            if(Thread.CurrentThread.ManagedThreadId!=threadId||Thread.CurrentThread.GetApartmentState()!=ApartmentState.STA)throw new InvalidOperationException("API calls require the Office UI STA.");
            if(String.IsNullOrWhiteSpace(json)||json.Length>Host.MetadataCodec.MaxCharacters)throw new InvalidDataException("Invalid JSON size.");
            Host.ChartDocument model;
            using(var stream=new MemoryStream(Encoding.UTF8.GetBytes(json)))model=(Host.ChartDocument)new DataContractJsonSerializer(typeof(Host.ChartDocument)).ReadObject(stream);
            Host.MetadataCodec.Validate(model);model.ChartId=Guid.NewGuid().ToString("D");model.Revision=0;
            new Host.PowerPointHost(app).InsertEditableChart(model);return model.ChartId;
        }
    }
}
