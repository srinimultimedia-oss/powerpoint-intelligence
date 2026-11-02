using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace PowerPointIntelligence.AddIn
{
    internal static class DemoRibbon
    {
        internal const string BuildMarker = "PPI-DEMO-RIBBON-20261029-v17-piidx7-typed-series";
        internal static string Create()
        {
            var text=new StringBuilder();
            using(var xml=XmlWriter.Create(text,new XmlWriterSettings{OmitXmlDeclaration=true}))
            {
                xml.WriteStartElement("customUI","http://schemas.microsoft.com/office/2009/07/customui");xml.WriteAttributeString("onLoad","OnLoad");xml.WriteAttributeString("loadImage","LoadImage");
                xml.WriteStartElement("ribbon");xml.WriteStartElement("tabs");xml.WriteStartElement("tab");xml.WriteAttributeString("id","ppiTab");xml.WriteAttributeString("label","Intelligence");xml.WriteAttributeString("keytip","PI");
                Group(xml,"ppiStructure","Structure");
                xml.WriteStartElement("menu");xml.WriteAttributeString("id","ppiAgenda");xml.WriteAttributeString("label","Agenda");xml.WriteAttributeString("size","large");xml.WriteAttributeString("image",RibbonIcons.Key("agenda-insert",32));
                Workflow(xml,"agenda-insert","Chapter","agenda");Workflow(xml,"agenda-contents","Table of contents","agenda");Workflow(xml,"agenda-settings","Manage Agenda","agenda");Workflow(xml,"agenda-split","Split agenda here","agenda");Workflow(xml,"agenda-refresh","Update agenda","agenda");xml.WriteEndElement();
                Tools(xml,"Slides");xml.WriteEndElement();
                Group(xml,"ppiElementsGroup","Elements");
                Button(xml,"ppiTemplateLibrary","Library","ShowTemplateLibrary",null,true,"Browse native previews and insert into the current presentation.","library-browser");
                Tools(xml,"Charts",true,"charts-gallery");
                Button(xml,"ppiRenderShapes","Editable shape columns","InsertRenderedShapes",null,false,"Linked editable shapes.");
                Button(xml,"ppiInsert","Linked column chart","InsertSample",null,false);
                Button(xml,"ppiContextProperties","Properties","ShowContextProperties",null,false,"Show the selection-aware floating property box.","element-edit");Tools(xml,"Elements");ElementMenu(xml);ProductivityMenu(xml);Tools(xml,"Chart options");
                Workflow(xml,"percent-area","100% stacked area","area");Workflow(xml,"series-settings","Series type and axis","combo");Workflow(xml,"axis-settings","Value axis settings","axes");Workflow(xml,"match-scales","Match value scales","axes");Workflow(xml,"annotation-refresh","Refresh selected annotation","value-line");xml.WriteEndElement();
                Group(xml,"ppiAI","AI");Workflow(xml,"connections","Configure services","category-data");xml.WriteEndElement();
                Group(xml,"ppiData","Data");
                xml.WriteStartElement("menu");xml.WriteAttributeString("id","ppiDataSource");xml.WriteAttributeString("label","Data Source");xml.WriteAttributeString("size","large");
                xml.WriteAttributeString("image",RibbonIcons.Key("category-data",32));xml.WriteAttributeString("getEnabled","GetDataSourceEnabled");
                foreach(var option in DataSources.Options)Button(xml,"ppiDataSource_"+option.Id,option.Label,"LinkDataSource",option.Id,false,"Choose a source file and save its reference with this presentation.",option.Icon);
                xml.WriteEndElement();Workflow(xml,"relink-selected","Relink selected chart","link-excel");Workflow(xml,"map-excel","Link Excel Range","link-excel");Workflow(xml,"data-manager","Data Link Manager","category-data");Tools(xml,"Data");xml.WriteEndElement();
                Group(xml,"ppiText","Text");Tools(xml,"Text");
                foreach(var mode in new[]{"left","center","right","justify","top","middle","bottom","horizontal","downward","upward"})Workflow(xml,"text-"+mode,"Text "+mode,"category-text");xml.WriteEndElement();
                Group(xml,"ppiLayout","Layout");Tools(xml,"Arrange");Button(xml,"ppiLibraryLayout","Layout Links","ShowLibraryLayout",null,false,null,"layout-links");
                Button(xml,"ppiTemplateAlignment","Template Alignment","ChangeTemplateAlignment",null,false,null,"template-alignment");xml.WriteEndElement();
                Group(xml,"ppiColor","Color");foreach(var kind in new[]{"fill","outline","text"})Workflow(xml,"color-"+kind,"Choose "+kind+" color","theme");Button(xml,"ppiTemplatePalette","Template Colors","ChangeTemplatePalette",null,false,null,"template-palette");xml.WriteEndElement();
                Group(xml,"ppiQuality","Quality");Workflow(xml,"quality-review","Presentation review","check-sync");Button(xml,"ppiInspect","Check & Sync","InspectChart",null,true);xml.WriteEndElement();
                Group(xml,"ppiCollaboration","Collaboration");Workflow(xml,"export-pdf","Export PDF","export-slide");Button(xml,"ppiCreateTour","Generate Linked Report","RunDemoTool","demo-tour",true);xml.WriteEndElement();
                Group(xml,"ppiMore","More");Tools(xml,"Start here");Tools(xml,"Assets");Tools(xml,"Automation");
                Button(xml,"ppiOpenWorkspace","Open Workspace","ShowWorkspace",null,false);Button(xml,"ppiDemo","Explore Demo Tools","ShowDemo",null,false);Button(xml,"ppiCompatibility","Office Compatibility","ShowCompatibility",null,false);xml.WriteEndElement();
                xml.WriteEndElement();xml.WriteEndElement();xml.WriteEndElement();xml.WriteEndElement();
            }
            return text.ToString();
        }
        private static void ProductivityMenu(XmlWriter xml)
        {
            xml.WriteStartElement("menu");xml.WriteAttributeString("id","ppiProductivityTools");xml.WriteAttributeString("label","Tools");xml.WriteAttributeString("image",RibbonIcons.Key("category-automation",16));
            string[][] commands={new[]{"symbol-picker","Insert symbol…","symbol"},new[]{"decimal-safe","Switch decimal symbols…","decimal"},new[]{"raster-copy","Create picture copy (retain original)","image"},new[]{"chart-workbooks","Chart workbooks…","slide-workbook"},new[]{"data-links","Data links…","category-data"},new[]{"save-selected","Save selected slides…","save-slides"},new[]{"selected-style","Apply selected-object style…","style-file"},new[]{"review","Review presentation…","quality-review"},new[]{"help","Element tools help","compatibility"}};
            foreach(var c in commands)Button(xml,"ppiProductivity_"+c[0],c[1],"RunProductivityTool",c[0],false,"Independent local tools. No automatic email or vendor object conversion.",c[2]);xml.WriteEndElement();
        }
        private static void ElementMenu(XmlWriter xml)
        {
            xml.WriteStartElement("menu");xml.WriteAttributeString("id","ppiElementControls");xml.WriteAttributeString("label","Element Controls");xml.WriteAttributeString("image",RibbonIcons.Key("element-edit",16));
            string[][] commands={
                new[]{"pentagon","Right pentagon","pentagon"},new[]{"chevron","Chevron","chevron"},new[]{"connector","Universal connector","universal-connector"},
                new[]{"edit","Edit selected element","element-edit"},new[]{"fit","Fit text to unlocked edges","text"},
                new[]{"convert-text","Convert to Text Box","text"},new[]{"convert-pentagon","Convert to Right Pentagon","pentagon"},new[]{"convert-chevron","Convert to Chevron","chevron"},new[]{"convert-rounded","Convert to Rounded Rectangle","rounded"},
                new[]{"duplicate-left","Duplicate left","left"},new[]{"duplicate-right","Duplicate right","right"},new[]{"duplicate-up","Duplicate above","top"},new[]{"duplicate-down","Duplicate below","bottom"},
                new[]{"indicator-next","Increase / next indicator state","harvey"},new[]{"indicator-previous","Decrease / previous state","harvey"},new[]{"indicator-keys","Indicator keyboard controls","checkbox"},new[]{"convert-harvey","Convert indicator to Harvey ball","harvey"},new[]{"convert-checkbox","Convert indicator to Checkbox","checkbox"},
                new[]{"grid-create","New process grid","process"},new[]{"grid-row","Add grid row","table"},new[]{"grid-column","Add grid column","table"},new[]{"grid-remove-row","Hide grid row","table"},new[]{"grid-remove-column","Hide grid column","table"},new[]{"grid-left","Move grid column left","left"},new[]{"grid-right","Move grid column right","right"},new[]{"grid-stripe-rows","Stripe grid rows","theme"},new[]{"grid-stripe-columns","Stripe grid columns","theme"},new[]{"grid-resize","Resize process grid","auto-layout"},new[]{"refresh","Refresh element geometry","check-sync"}};
            foreach(var c in commands)Button(xml,"ppiElement_"+c[0],c[1],"RunElementCommand",c[0],false,"Editable native objects. Select the relevant managed element first.",c[2]);xml.WriteEndElement();
        }
        private static void Workflow(XmlWriter xml,string id,string label,string icon){Button(xml,"ppiWorkflow_"+id,label,"RunWorkflowCommand",id,false,"Native command; select the relevant slide or object first.",id);}
        private static void Tools(XmlWriter xml,string category,bool large=false,string icon=null)
        {
            var tools=DemoTools.Catalog.Where(t=>t.Category==category&&t.Id!="demo-tour").ToArray();if(tools.Length==0)return;
            xml.WriteStartElement("menu");xml.WriteAttributeString("id","ppiCategory"+new String(category.Where(Char.IsLetterOrDigit).ToArray()));xml.WriteAttributeString("label",category);
            xml.WriteAttributeString("image",RibbonIcons.Key(icon??("category-"+category.ToLowerInvariant().Replace(" ","")),large?32:16));if(large){xml.WriteAttributeString("size","large");xml.WriteAttributeString("keytip","CH");xml.WriteAttributeString("supertip","Choose a chart type and build an editable chart from the linked data source.");}
            foreach(var tool in tools)Button(xml,"ppiTool_"+tool.Id,tool.Name,"RunDemoTool",tool.Id,false,tool.Description);xml.WriteEndElement();
        }
        private static void Group(XmlWriter xml,string id,string label)
        {
            string icon;
            switch(id){
                case "ppiStructure":icon="agenda-insert";break;
                case "ppiElementsGroup":icon="category-charts";break;
                case "ppiAI":icon="connections";break;
                case "ppiData":icon="category-data";break;
                case "ppiText":icon="category-text";break;
                case "ppiLayout":icon="auto-layout";break;
                case "ppiColor":icon="theme";break;
                case "ppiQuality":icon="quality-review";break;
                case "ppiCollaboration":icon="save-slides";break;
                case "ppiMore":icon="group-more";break;
                default:throw new InvalidOperationException("Ribbon group requires an explicit collapsed-state icon: "+id);
            }
            xml.WriteStartElement("group");xml.WriteAttributeString("id",id);xml.WriteAttributeString("label",label);
            xml.WriteAttributeString("image",RibbonIcons.Key(icon,32));
        }
        private static void Button(XmlWriter xml,string id,string label,string action,string tag,bool large,string description=null,string iconId=null)
        {
            xml.WriteStartElement("button");xml.WriteAttributeString("id",id);xml.WriteAttributeString("label",label);xml.WriteAttributeString("onAction",action);
            string icon=iconId??tag??(id=="ppiRenderShapes"?"editable-columns":id=="ppiInsert"?"managed-column":id=="ppiInspect"?"check-sync":id=="ppiOpenWorkspace"?"workspace":id=="ppiDemo"?"explore":"compatibility");
            xml.WriteAttributeString("image",RibbonIcons.Key(icon,large?32:16));
            if(tag!=null)xml.WriteAttributeString("tag",tag);if(large)xml.WriteAttributeString("size","large");
            if(action=="RunWorkflowCommand")xml.WriteAttributeString("getEnabled","GetWorkflowEnabled");
            if(action=="RunElementCommand")xml.WriteAttributeString("getEnabled","GetElementEnabled");
            if(action=="RunProductivityTool")xml.WriteAttributeString("getEnabled","GetDataSourceEnabled");
            if(action=="RunDemoTool"||action=="InsertSample"||action=="InsertRenderedShapes"||action=="InspectChart")xml.WriteAttributeString("getEnabled","GetCommandEnabled");
            if(id=="ppiTemplateAlignment")xml.WriteAttributeString("keytip","TA");
            if(action=="ChangeTemplateAlignment")xml.WriteAttributeString("getEnabled","GetDataSourceEnabled");
            if(id=="ppiTemplatePalette")xml.WriteAttributeString("keytip","CP");
            if(action=="ChangeTemplatePalette")xml.WriteAttributeString("getEnabled","GetDataSourceEnabled");
            if(action=="ShowContextProperties")xml.WriteAttributeString("getEnabled","GetDataSourceEnabled");
            if(action=="ShowTemplateLibrary"||action=="ShowLibraryLayout")xml.WriteAttributeString("getEnabled","GetDataSourceEnabled");
            if(action=="LinkDataSource")xml.WriteAttributeString("getEnabled","GetDataSourceEnabled");
            if(description!=null)xml.WriteAttributeString("supertip",description);xml.WriteEndElement();
        }
    }
}
