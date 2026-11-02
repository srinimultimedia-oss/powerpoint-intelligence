using System;
using System.IO;
using System.Linq;
using System.Xml;
using System.Drawing;
using System.Drawing.Imaging;
using PowerPointIntelligence.AddIn;
internal static class DemoContractTests
{
    [STAThread]
    public static int Main(string[] args)
    {
        try
        {
            if(args.Length==3&&args[0]=="--gallery-snapshot"){GalleryTests.Snapshot(args[1],args[2]);return 0;}
            if(args.Length==3&&args[0]=="--fidelity-replay"){TemplateFidelityTests.Replay(args[1],args[2]);return 0;}
            TemplateFidelityTests.Run();LibraryBrowserTests.Run();GalleryTests.Run(); PiidxGalleryTests.Run();
            var tools=DemoTools.Catalog;
            NativeFormattingTests.Run();PresentationQualityTests.Run();NativeChartControlTests.Run();ExcelRangeMappingTests.Run();DataSourceTests.Run(args[1]);ChartAnnotationTests.Run();PresentationAgendaTests.Run();ChapterAgendaTests.Run();ContextPropertiesTests.Run(args[1]);ContextPanelModelTests.Run();ContextFormattingTests.Run();ContextChartFormattingTests.Run();XmlPartUpdateTests.Run();ProductivityToolTests.Run();IndicatorElementTests.Run();UniversalConnectorTests.Run();TextShapeElementTests.Run();ProcessFlowTests.Run();ServiceConnectionTests.Run(args[1]);DataLinkManagerTests.Run();
            var cancelled=new DemoInputs { Prompt=(title,initial)=>null, OpenFile=filter=>null, SaveFile=(filter,name)=>null, Folder=()=>null };
            foreach(var id in new[]{"symbol","font","font-size","language","chart-style","chart-title","axes","cagr","round-data","excel-cagr","image","library","template","export-slide","import-json","style-file","library-browser","link-excel"})
            {
                tools.Single(t=>t.Id==id).Execute(new DemoTools(null,cancelled));
                Console.WriteLine("PASS cancellation exits before document mutation: "+id);
            }
            foreach(var input in new[]{new[]{"font-size","NaN"},new[]{"font-size","Infinity"},new[]{"font-size","201"},new[]{"round-data","11"},new[]{"language","bad"},new[]{"cagr","100,121,0"},new[]{"cagr","NaN,121,2"},new[]{"excel-cagr","NaN,121,2,E2"}})
            {
                bool rejected=false;
                try{tools.Single(t=>t.Id==input[0]).Execute(new DemoTools(null,new DemoInputs{Prompt=(title,initial)=>input[1]}));}
                catch(InvalidOperationException){rejected=true;}
                if(!rejected)throw new Exception("Invalid input was accepted: "+input[0]);
                Console.WriteLine("PASS invalid input rejected before COM: "+input[0]);
            }
            if(tools.Length!=90||tools.Select(t=>t.Id).Distinct().Count()!=tools.Length)throw new Exception("Catalog count/unique IDs failed.");
            var document=new XmlDocument();document.LoadXml(DemoRibbon.Create());
            var ns=new XmlNamespaceManager(document.NameTable);ns.AddNamespace("r","http://schemas.microsoft.com/office/2009/07/customui");
            var workflowGroups=document.SelectNodes("//r:tab[@id='ppiTab']/r:group",ns).Cast<XmlElement>().Select(e=>e.GetAttribute("label")).ToArray();
            if(String.Join(",",workflowGroups)!="Structure,Elements,AI,Data,Text,Layout,Color,Quality,Collaboration,More")throw new Exception("Workflow group order changed.");
            foreach(var id in new[]{"agenda-insert","agenda-contents","agenda-settings","agenda-split","agenda-refresh","annotation-refresh","percent-area","connections","text-left","text-center","text-right","text-justify","map-excel","relink-selected","data-manager","axis-settings","match-scales","quality-review","text-top","text-middle","text-bottom","text-horizontal","text-downward","text-upward","color-fill","color-outline","color-text","series-settings","export-pdf"})
                if(document.SelectNodes("//r:button[@tag='"+id+"' and @onAction='RunWorkflowCommand' and @getEnabled='GetWorkflowEnabled']",ns).Count!=1)throw new Exception("Missing workflow binding: "+id);
            if(document.SelectNodes("//r:button[@id='ppiContextProperties' and @onAction='ShowContextProperties' and @getEnabled='GetDataSourceEnabled']",ns).Count!=1)throw new Exception("Contextual Properties ribbon entry missing.");
            foreach(var command in ElementCommands.Ids)if(document.SelectNodes("//r:button[@tag='"+command+"' and @onAction='RunElementCommand' and @getEnabled='GetElementEnabled']",ns).Count!=1)throw new Exception("Missing element binding: "+command);
            foreach(var command in ProductivityTools.Ids)if(document.SelectNodes("//r:button[@tag='"+command+"' and @onAction='RunProductivityTool' and @getEnabled='GetDataSourceEnabled']",ns).Count!=1)throw new Exception("Missing productivity binding: "+command);
            if(ElementCommands.Ids.Distinct().Count()!=ElementCommands.Ids.Length)throw new Exception("Duplicate element command identity.");
            if(document.SelectNodes("//r:tab[@id='ppiTab']/r:group",ns).Count<4)throw new Exception("Ribbon structure/namespace failed.");
            if(document.SelectNodes("//r:button[@id='ppiRenderShapes' and @onAction='InsertRenderedShapes' and @getEnabled='GetCommandEnabled']",ns).Count!=1)throw new Exception("Native renderer ribbon binding missing or duplicated.");
            if(document.SelectNodes("//r:button[@id='ppiTemplateLibrary' and @label='Library' and @onAction='ShowTemplateLibrary' and @getEnabled='GetDataSourceEnabled']",ns).Count!=1)throw new Exception("Template library ribbon binding missing.");
            if(document.SelectNodes("//r:button[@id='ppiTemplateLibrary' and @size='large']/following-sibling::*[1][self::r:menu and @id='ppiCategoryCharts' and @label='Charts' and @size='large' and @image='ppi_charts-gallery_32' and @keytip='CH']",ns).Count!=1)throw new Exception("Large Charts gallery must sit immediately beside Library.");
            if(document.SelectNodes("//r:menu[@id='ppiCategoryCharts']",ns).Count!=1)throw new Exception("Charts menu must not be duplicated.");
            if(document.SelectNodes("//r:button[@id='ppiTemplatePalette' and @onAction='ChangeTemplatePalette' and @keytip='CP']",ns).Count!=1)throw new Exception("Template palette shortcut missing.");
            if(document.SelectNodes("//r:button[@id='ppiTemplateAlignment' and @onAction='ChangeTemplateAlignment' and @keytip='TA']",ns).Count!=1)throw new Exception("Template alignment shortcut missing.");
            foreach(var option in DataSources.Options)
                if(document.SelectNodes("//r:menu[@id='ppiDataSource']/r:button[@tag='"+option.Id+"' and @onAction='LinkDataSource' and @getEnabled='GetDataSourceEnabled']",ns).Count!=1)throw new Exception("Missing Data Source menu option: "+option.Id);
            var ids=document.SelectNodes("//*[@id]").Cast<XmlElement>().Select(e=>e.GetAttribute("id")).ToArray();
            if(ids.Distinct().Count()!=ids.Length)throw new Exception("Duplicate ribbon ID.");
            foreach(var tool in tools)
            {
                if(tool.Execute==null||String.IsNullOrWhiteSpace(tool.Description))throw new Exception("Missing command implementation/description: "+tool.Id);
                var nodes=document.SelectNodes("//r:button[@tag='"+tool.Id+"' and @onAction='RunDemoTool']",ns);
                if(nodes.Count!=1)throw new Exception("Tool not exposed exactly once on ribbon: "+tool.Id);
                var scenario=File.ReadAllText("tests/s0/DemoScenario.cs");
                bool asserted=scenario.Contains("case \""+tool.Id+"\":")||scenario.Contains("{\""+tool.Id+"\",Office.XlChartType.")||tool.Id=="combo"&&scenario.Contains("if(id==\"combo\")Check")||tool.Id=="demo-tour"&&scenario.Contains("Deck.Slides.Count==8");
                if(!asserted)throw new Exception("Feature assertion missing: "+tool.Id);
                Console.WriteLine("PASS ribbon command binding: "+tool.Id);
            }
            var callbacks=document.SelectNodes("//@onAction | //@onLoad | //@getEnabled | //@loadImage").Cast<XmlAttribute>().Select(a=>a.Value).Distinct();
            string source=File.ReadAllText(args[0]);
            foreach(var callback in callbacks)if(!System.Text.RegularExpressions.Regex.IsMatch(source,@"public\s+\w+\s+"+callback+@"\("))throw new Exception("Missing callback: "+callback);
            if(FeatureInventory.Items.Select(f=>f.Id).Distinct().Count()!=103)throw new Exception("Inventory identities failed.");
            Directory.CreateDirectory(args[1]);document.Save(Path.Combine(args[1],"demo-ribbon.xml"));
            VerifyIcons(document,ns,args[1]);
            File.WriteAllLines(Path.Combine(args[1],"demo-tool-ids.txt"),tools.Select(t=>t.Id));
            Console.WriteLine("PASS compiled ribbon XML, callbacks, all 90 tool bindings and 103 inventory identities. Office behavior NOT tested.");return 0;
        }
        catch(Exception ex){Console.Error.WriteLine(ex);return 1;}
    }
    private static void VerifyIcons(XmlDocument document,XmlNamespaceManager ns,string directory)
    {
        if(document.DocumentElement.GetAttribute("loadImage")!="LoadImage")throw new Exception("Missing custom image callback.");
        var controls=document.SelectNodes("//r:button | //r:menu | //r:group",ns).Cast<XmlElement>().ToArray();
        int columns=5,cellWidth=240,cellHeight=70;
        using(var sheet=new Bitmap(columns*cellWidth,40+((controls.Length+columns-1)/columns)*cellHeight))
        using(var graphics=Graphics.FromImage(sheet))
        using(var font=new Font("Segoe UI",10,GraphicsUnit.Pixel))
        using(var title=new Font("Segoe UI",17,FontStyle.Bold,GraphicsUnit.Pixel))
        {
            graphics.Clear(Color.White);graphics.DrawString("Intelligence ribbon • original icons • 32 / 16 px on light and dark",title,Brushes.Black,8,8);
            for(int i=0;i<controls.Length;i++)
            {
                var control=controls[i];string key=control.GetAttribute("image");
                if(String.IsNullOrWhiteSpace(key)||control.HasAttribute("imageMso")||control.HasAttribute("getImage"))throw new Exception("Invalid icon binding: "+control.GetAttribute("id"));
                string icon=key.Substring(4,key.LastIndexOf('_')-4);
                if(!key.EndsWith((control.LocalName=="group"||control.GetAttribute("size")=="large")?"_32":"_16",StringComparison.Ordinal))throw new Exception("Wrong icon size: "+key);
                object picture=RibbonIcons.Load(key);
                if(picture==null||!Object.ReferenceEquals(picture,RibbonIcons.Load(key)))throw new Exception("Image callback/cache failed: "+key);
                int x=(i%columns)*cellWidth,y=40+(i/columns)*cellHeight;
                graphics.FillRectangle(Brushes.WhiteSmoke,x,y,cellWidth-2,cellHeight-2);
                using(var dark=new SolidBrush(Color.FromArgb(32,32,32)))graphics.FillRectangle(dark,x+107,y,100,39);
                foreach(int size in new[]{16,32})
                {
                    using(var bitmap=RibbonIcons.Render(icon,size))
                    {
                        if(bitmap==null||bitmap.Width!=size||bitmap.Height!=size||bitmap.PixelFormat!=PixelFormat.Format32bppArgb)throw new Exception("Missing icon: "+icon);
                        int visible=0,transparent=0;
                        for(int py=0;py<size;py++)for(int px=0;px<size;px++) { if(bitmap.GetPixel(px,py).A>0)visible++;else transparent++; }
                        if(visible<8||transparent<8)throw new Exception("Blank/opaque icon: "+icon);
                        int offset=size==32?8:54;
                        graphics.DrawImageUnscaled(bitmap,x+offset,y+3);
                        graphics.DrawImageUnscaled(bitmap,x+offset+107,y+3);
                    }
                }
                graphics.DrawString(control.GetAttribute("label"),font,Brushes.Black,new RectangleF(x+5,y+41,cellWidth-10,25));
            }
            sheet.Save(Path.Combine(directory,"ribbon-icons.png"),ImageFormat.Png);
        }
        var groups=document.SelectNodes("//r:tab[@id='ppiTab']/r:group",ns).Cast<XmlElement>().ToArray();
        if(groups.Length!=10)throw new Exception("Expected ten icon-bearing ribbon groups.");
        var groupHashes=new System.Collections.Generic.HashSet<string>();
        foreach(var group in groups)
        {
            string key=group.GetAttribute("image");if(!key.StartsWith("ppi_",StringComparison.Ordinal)||!key.EndsWith("_32",StringComparison.Ordinal))throw new Exception("Collapsed group is missing its 32 px icon: "+group.GetAttribute("id"));
            string icon=key.Substring(4,key.LastIndexOf('_')-4);
            using(var bitmap=RibbonIcons.Render(icon,32))using(var stream=new MemoryStream())using(var sha=System.Security.Cryptography.SHA256.Create())
            {bitmap.Save(stream,ImageFormat.Png);if(!groupHashes.Add(Convert.ToBase64String(sha.ComputeHash(stream.ToArray()))))throw new Exception("Collapsed groups share identical artwork: "+group.GetAttribute("id"));}
        }
        Console.WriteLine("PASS all ten collapsed ribbon groups have distinct 32 px images through the loadImage callback.");
        // Workflow operations need dedicated semantic artwork, not a shared category glyph.
        var workflowControls=document.SelectNodes("//r:button[@onAction='RunWorkflowCommand']",ns).Cast<XmlElement>().ToArray();
        foreach(int size in new[]{16,32})
        {
            var hashes=new System.Collections.Generic.HashSet<string>();
            foreach(var control in workflowControls)
            {
                string key=control.GetAttribute("image");string icon=key.Substring(4,key.LastIndexOf('_')-4);
                if(icon!=control.GetAttribute("tag"))throw new Exception("Workflow uses a generic icon: "+control.GetAttribute("id"));
                using(var rendered=RibbonIcons.Render(icon,size))using(var bytes=new MemoryStream())using(var sha=System.Security.Cryptography.SHA256.Create())
                {rendered.Save(bytes,ImageFormat.Png);if(!hashes.Add(Convert.ToBase64String(sha.ComputeHash(bytes.ToArray()))))throw new Exception("Visually duplicated workflow icon: "+icon+" at "+size+"px");}
            }
        }
        foreach(var pair in new[]{new[]{"template-palette","template-alignment"},new[]{"library-browser","template-palette"},new[]{"source-excel","source-csv"},new[]{"column","stacked-column"},new[]{"bar","stacked-bar"},new[]{"editable-columns","managed-column"},new[]{"charts-gallery","category-charts"},new[]{"charts-gallery","library-browser"}})
        {
            using(var first=RibbonIcons.Render(pair[0],16))using(var second=RibbonIcons.Render(pair[1],16))
            {bool differs=false;for(int y=0;y<16;y++)for(int x=0;x<16;x++)if(first.GetPixel(x,y)!=second.GetPixel(x,y))differs=true;if(!differs)throw new Exception("Semantically different commands share artwork: "+String.Join("/",pair));}
        }
        var focus=document.SelectNodes("//r:group | //r:button[@onAction='RunWorkflowCommand' or @onAction='LinkDataSource' or @id='ppiTemplateLibrary' or @id='ppiTemplatePalette' or @id='ppiTemplateAlignment' or @id='ppiRenderShapes'] | //r:menu[@id='ppiCategoryCharts']",ns).Cast<XmlElement>().ToArray();
        using(var preview=new Bitmap(1000,44+((focus.Length+3)/4)*82))using(var g=Graphics.FromImage(preview))using(var font=new Font("Segoe UI",9))
        {
            g.Clear(Color.White);g.DrawString("Intelligence · New ribbon icons · 32 px and 16 px · light / dark",font,Brushes.Black,12,10);
            for(int i=0;i<focus.Length;i++)
            {
                int x=i%4*250,y=44+i/4*82;string key=focus[i].GetAttribute("image"),icon=key.Substring(4,key.LastIndexOf('_')-4);
                g.FillRectangle(Brushes.WhiteSmoke,x+2,y,244,76);using(var dark=new SolidBrush(Color.FromArgb(32,32,32)))g.FillRectangle(dark,x+126,y+3,116,42);
                foreach(int size in new[]{16,32})using(var bitmap=RibbonIcons.Render(icon,size)){int offset=size==32?12:62;g.DrawImageUnscaled(bitmap,x+offset,y+7);g.DrawImageUnscaled(bitmap,x+offset+120,y+7);}
                g.DrawString(focus[i].GetAttribute("label"),font,Brushes.Black,new RectangleF(x+8,y+49,235,25));
            }
            preview.Save(Path.Combine(directory,"ribbon-icons-focus.png"),ImageFormat.Png);
        }
        Console.WriteLine("PASS all "+workflowControls.Length+" workflow icons are distinct at 16/32 px; template/data/chart actions have dedicated bindings.");
        if(RibbonIcons.Load("ppi_unknown_32")!=null||RibbonIcons.Load("ppi_column_512")!=null||RibbonIcons.Load(null)!=null)throw new Exception("Unknown image request not rejected.");
        Console.WriteLine("PASS original icons for all "+controls.Length+" ribbon buttons/menus/groups, 16/32 px rendering, transparency and cached OLE image callback.");
    }
}

