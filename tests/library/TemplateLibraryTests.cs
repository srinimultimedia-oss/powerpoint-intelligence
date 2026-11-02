using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using PowerPointIntelligence.AddIn;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
using Office=Microsoft.Office.Core;
internal static class TemplateLibraryTests
{
    private static void Check(bool value,string name){if(!value)throw new Exception(name);}
    private static string Signature(IEnumerable<TemplateNode> nodes)
    {
        string value=String.Join("|",nodes.Select(n=>n.Kind+":"+String.Join(",",new[]{n.X,n.Y,n.W,n.H}.Select(x=>x.ToString("R",CultureInfo.InvariantCulture)))+":"+(n.Points==null?"":String.Join(",",n.Points.Select(x=>x.ToString("R",CultureInfo.InvariantCulture))))));
        using(var hash=SHA256.Create())return BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(value))).Replace("-","");
    }
    [STAThread] public static int Main(string[] args)
    {
        try
        {
            TemplateFidelityTests.Run();
            string output=Path.GetFullPath(args[0]);Directory.CreateDirectory(output);
            if(args.Contains("--source-originals")){OriginalRoundtrip(output);return 0;}
            Check(TemplateCatalog.Items.Count==520,"520 templates required");Check(TemplateCatalog.Items.Select(t=>t.Id).Distinct().Count()==520,"Unique template IDs");Check(TemplateCatalog.Items.Select(t=>t.Category).Distinct().Count()==26,"26 families required");
            Check(TemplatePalettes.Items.Count==20&&TemplatePalettes.Items.All(p=>p.Colors.Length==10),"20 palettes with 10 colors each");
            Check(TemplatePalettes.Get().Name=="Executive Slate"&&TemplatePalettes.Get().Colors[0]=="#34495E"&&TemplatePalettes.Get().Colors[9]=="#F5F7F8","First supplied palette is default");
            bool bad=false;try {TemplatePalettes.Get("unknown");}catch(InvalidDataException){bad=true;}Check(bad,"Unknown palettes rejected");
            foreach(var palette in TemplatePalettes.Items){foreach(var color in palette.Colors)Check(TemplatePalettes.Contrast(TemplatePalettes.TextColor(color,palette),color)>=4.5,"Readable foreground against palette surface");using(var image=TemplatePreview.Render(TemplateCatalog.Items[92],960,540,null,palette.Id))image.Save(Path.Combine(output,"palette-"+palette.Id+".png"),ImageFormat.Png);}
            foreach(var h in TemplateAlignment.Horizontal)foreach(var v in TemplateAlignment.Vertical)TemplateAlignment.Validate(h,v);
            bool invalidAlignment=false;try {TemplateAlignment.Validate("invalid","Middle");}catch(InvalidDataException){invalidAlignment=true;}Check(invalidAlignment,"Invalid alignment rejected");
            var centered=TemplateAlignment.Prepare(new List<TemplateNode>{new TemplateNode {Kind="rect",X=100,Y=100,W=200,H=100,Stage=1},new TemplateNode {Kind="text",X=110,Y=110,W=180,H=20,Stage=1}});
            Check(centered[1].Y==104&&centered[1].H==92&&centered[1].X==108,"Single shape labels fit parent interior");
            var multi=TemplateAlignment.Prepare(new List<TemplateNode>{new TemplateNode {Kind="rect",X=100,Y=100,W=200,H=100,Stage=1},new TemplateNode {Kind="text",X=110,Y=110,W=180,H=20,Stage=1},new TemplateNode {Kind="text",X=110,Y=150,W=180,H=30,Stage=1}});
            Check(multi[1].Y==110&&multi[2].Y==150,"Multi-field cards retain separate field positions");
            foreach(var h in TemplateAlignment.Horizontal)foreach(var v in TemplateAlignment.Vertical)using(var image=TemplatePreview.Render(TemplateCatalog.Items[92],960,540,null,null,h,v))image.Save(Path.Combine(output,"alignment-"+h+"-"+v+".png"),ImageFormat.Png);
            var signatures=new Dictionary<string,string>();var rows=new List<string>{"id\tcategory\tname\tfields\tnodes\tstructureSha256"};
            var duplicates=new List<string>();
            foreach(var template in TemplateCatalog.Items)
            {
                var nodes=template.Build();TemplateCatalog.Validate(nodes);Check(nodes.Any(n=>n.Key=="title")&&nodes.Count(n=>n.Kind=="text")>=3,"Editable fields required: "+template.Id);
                string signature=Signature(nodes);if(signatures.ContainsKey(signature))duplicates.Add(template.Id+" duplicates "+signatures[signature]);else signatures.Add(signature,template.Id);
                foreach(var palette in TemplatePalettes.Items){var themed=template.Build(palette.Id);Check(Signature(themed)==signature,"Palette changed layout");Check(themed.Where(n=>n.Kind=="text").Select(n=>n.Text).SequenceEqual(nodes.Where(n=>n.Kind=="text").Select(n=>n.Text)),"Palette changed fields");}
                Check(signature==Signature(template.Build()),"Nondeterministic layout "+template.Id);
                using(var preview=TemplatePreview.Render(template,384,216)){Check(preview.Width==384,"Preview dimensions");preview.Save(Path.Combine(output,template.Id+".png"),ImageFormat.Png);}
                rows.Add(template.Id+"\t"+template.Category+"\t"+template.Name+"\t"+nodes.Count(n=>n.Kind=="text")+"\t"+nodes.Count+"\t"+signature);
            }
            File.WriteAllLines(Path.Combine(output,"catalog.tsv"),rows);File.WriteAllLines(Path.Combine(output,"duplicate-structures.txt"),duplicates);
            Check(duplicates.Count==0,"Structurally duplicate templates: "+String.Join("; ",duplicates));
            Check(TemplateCatalog.Search("Decisions",null).Count()==20,"Category filtering");Check(TemplateCatalog.Search(null,"procurement").Any(),"Keyword search");Check(!TemplateCatalog.Search(null,"no-such-template-123456").Any(),"No-match search");
            Check(TemplateCatalog.Search(null,"  VENDOR ").Count()==TemplateCatalog.Search(null,"vendor").Count(),"Search normalization");
            foreach(var template in TemplateCatalog.Items.Where(t=>t.Variant==12))using(var preview=TemplatePreview.Render(template,960,540))preview.Save(Path.Combine(output,"featured-"+template.Family.ToString("D2")+".png"),ImageFormat.Png);
            ContactSheets(output);
            using(var browser=new TemplateLibrary(null)){Check(browser.Text.Contains("520"),"Browser catalog count");Check(browser.Controls.Count>0,"Browser controls absent");}
            Console.WriteLine("PASS 20 supplied palettes, default and text contrast, 10400 layout/palette combinations preserve geometry and fields; 520 deterministic layouts, unique IDs and geometry signatures, editable fields, bounds, 520 rendered previews, search and 26 contact sheets.");
            if(args.Contains("--interactive-office"))OfficeRoundtrip(output);
            else Console.WriteLine("SKIP Office insertion, native editability, animation and save/reopen without --interactive-office.");
            return 0;
        }
        catch(Exception ex){Console.Error.WriteLine(ex);return 1;}
    }
    private static void ContactSheets(string output)
    {
        for(int f=0;f<26;f++)using(var sheet=new Bitmap(1600,1320))using(var g=Graphics.FromImage(sheet))using(var font=new Font("Segoe UI",11,GraphicsUnit.Pixel))using(var heading=new Font("Segoe UI",24,FontStyle.Bold,GraphicsUnit.Pixel))
        {
            g.Clear(Color.FromArgb(226,232,240));g.DrawString(TemplateCatalog.Categories[f]+" · 20 structural configurations",heading,Brushes.Black,16,12);
            foreach(var t in TemplateCatalog.Items.Where(t=>t.Family==f)){int x=16+t.Variant%4*396,y=64+t.Variant/4*248;using(var image=Image.FromFile(Path.Combine(output,t.Id+".png")))g.DrawImage(image,x,y,384,216);g.DrawString(t.Name,font,Brushes.Black,x,y+218);}
            sheet.Save(Path.Combine(output,"family-"+f.ToString("D2")+".png"),ImageFormat.Png);
        }
    }
    private static PowerPoint.Shape Field(PowerPoint.Shapes shapes,string key)
    {foreach(PowerPoint.Shape shape in shapes){var found=FieldOrNull(shape,key);if(found!=null)return found;}throw new Exception("Editable field missing: "+key);}
    private static PowerPoint.Shape FieldOrNull(PowerPoint.Shape shape,string key)
    {
        if(String.Equals(shape.Tags[TemplateInserter.FieldTag],key,StringComparison.OrdinalIgnoreCase))return shape;
        if(shape.Type==Office.MsoShapeType.msoGroup)foreach(PowerPoint.Shape child in shape.GroupItems){var found=FieldOrNull(child,key);if(found!=null)return found;}
        return null;
    }
    private static void OfficeRoundtrip(string output)
    {
        PowerPoint.Application app=null;PowerPoint.Presentation owned=null;
        try
        {
            app=new PowerPoint.Application();app.Visible=Office.MsoTriState.msoTrue;
            if(app.Presentations.Count>0)throw new InvalidOperationException("Close user presentations before Office template validation.");
            owned=app.Presentations.Add(Office.MsoTriState.msoTrue);owned.PageSetup.SlideWidth=960;owned.PageSetup.SlideHeight=540;owned.Slides.Add(1,PowerPoint.PpSlideLayout.ppLayoutBlank).Select();
            foreach(var template in TemplateCatalog.Items)
            {
                bool builds=template.Variant!=0;string transition=template.Variant%3==0?"None":template.Variant%3==1?"Fade":"Wipe";string style=template.Variant%3==0?"Fade":template.Variant%3==1?"Wipe":"Fly in";
                var slide=TemplateInserter.Insert(app,template,null,builds,transition,style);Check(app.Presentations.Count==1,"Insertion opened another document");
                Check(Field(slide.Shapes,"title").TextFrame.VerticalAnchor==Office.MsoVerticalAnchor.msoAnchorMiddle&&Field(slide.Shapes,"title").TextFrame.TextRange.ParagraphFormat.Alignment==PowerPoint.PpParagraphAlignment.ppAlignCenter,"Default text not centered");
                Field(slide.Shapes,"title").TextFrame.TextRange.Text="Edited "+template.Id;
                string contentKey=template.Build().First(t=>t.Kind=="text"&&t.Key!="title"&&t.Key!="subtitle").Key;Field(slide.Shapes,contentKey).TextFrame.TextRange.Text="Editable content "+template.Id;
                slide.Select();TemplateInserter.ChangeAlignment(app,"Right","Bottom");Check(Field(slide.Shapes,"title").TextFrame.VerticalAnchor==Office.MsoVerticalAnchor.msoAnchorBottom,"Alignment command not applied");TemplateInserter.ChangeAlignment(app,"Center","Middle");
                int shapeCount=slide.Shapes.Count,effectCount=slide.TimeLine.MainSequence.Count;float titleX=Field(slide.Shapes,"title").Left;TemplateInserter.ChangePalette(app,TemplatePalettes.Items[template.Variant].Id);Check(slide.Shapes.Count==shapeCount&&slide.TimeLine.MainSequence.Count==effectCount&&Field(slide.Shapes,"title").Left==titleX,"Recolor changed objects, animations or layout");Check(Field(slide.Shapes,"title").TextFrame.TextRange.Text=="Edited "+template.Id,"Recolor changed text");
                Check((slide.TimeLine.MainSequence.Count>0)==builds,"Missing editable animation or unexpected animation in static template");Check(slide.SlideShowTransition.EntryEffect==(transition=="None"?PowerPoint.PpEntryEffect.ppEffectNone:transition=="Fade"?PowerPoint.PpEntryEffect.ppEffectFadeSmoothly:PowerPoint.PpEntryEffect.ppEffectWipeRight),"Missing native transition");
            }
            var existing=owned.Slides[1];existing.Select();int beforeSlides=owned.Slides.Count;int beforeBackground=existing.Background.Fill.ForeColor.RGB;
            var objectTemplate=TemplateCatalog.Items[0];TemplateInserter.Insert(app,objectTemplate,null,false,"Wipe",objectsOnly:true,destinationTheme:true);var firstGroup=app.ActiveWindow.Selection.ShapeRange[1];
            Check(owned.Slides.Count==beforeSlides&&existing.Background.Fill.ForeColor.RGB==beforeBackground,"Object-only insertion changed slide count or background");
            Check(Field(existing.Shapes,"title").TextFrame.TextRange.Font.Color.ObjectThemeColor==Office.MsoThemeColorIndex.msoThemeColorText1,"Destination theme text role missing");
            TemplateInserter.Insert(app,objectTemplate,null,false,"None",objectsOnly:true);var secondGroup=app.ActiveWindow.Selection.ShapeRange[1];int firstColor=FieldOrNull(firstGroup,"title").TextFrame.TextRange.Font.Color.RGB;
            TemplateInserter.ChangePalette(app,"03");Check(FieldOrNull(firstGroup,"title").TextFrame.TextRange.Font.Color.RGB==firstColor,"Recolor crossed template instance boundary");
            FieldOrNull(secondGroup,"title").TextFrame.TextRange.Text="Object-only edited";string objectInstance=secondGroup.Tags["PPI_TEMPLATE_INSTANCE"];
            existing.Select();var anchor=existing.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle,50,50,100,40);var follower=existing.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle,180,50,80,40);anchor.Select();follower.Select(Office.MsoTriState.msoFalse);LibraryLayoutLinks.Set(app,"Keep horizontal gap");anchor.Width+=20;LibraryLayoutLinks.Reflow(app);Check(Math.Abs(follower.Left-200)<.1,"Native layout gap did not follow resized anchor");
            string file=Path.Combine(output,"library-all-520-roundtrip.pptx");owned.SaveAs(file,PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);owned.Close();owned=null;owned=app.Presentations.Open(file);
            Check(owned.Slides.Count==521,"Slides lost on reopen");
            var restoredGroup=owned.Slides[1].Shapes.Cast<PowerPoint.Shape>().Single(s=>s.Tags["PPI_TEMPLATE_INSTANCE"]==objectInstance);Check(FieldOrNull(restoredGroup,"title").TextFrame.TextRange.Text=="Object-only edited","Object-only group edits lost after reopen");FieldOrNull(restoredGroup,"title").TextFrame.TextRange.Text="Object-only remains editable";
            var linked=owned.Slides[1].Shapes.Cast<PowerPoint.Shape>().Single(s=>!String.IsNullOrEmpty(s.Tags["PPI_LAYOUT_X"]));owned.Slides[1].Select();float oldLeft=linked.Left;var leader=owned.Slides[1].Shapes.Cast<PowerPoint.Shape>().Single(s=>!String.IsNullOrEmpty(s.Tags["PPI_LAYOUT_ANCHOR"]));leader.Width+=10;LibraryLayoutLinks.Reflow(app);Check(Math.Abs(linked.Left-oldLeft-10)<.1,"Layout links did not survive save/reopen");
            foreach(PowerPoint.Slide slide in owned.Slides){string id=slide.Tags[TemplateInserter.TemplateTag];if(String.IsNullOrEmpty(id))continue;Check(Field(slide.Shapes,"title").TextFrame.TextRange.Text=="Edited "+id.ToLowerInvariant(),"Edits lost after reopen");var source=TemplateCatalog.Items.Single(t=>String.Equals(t.Id,id,StringComparison.OrdinalIgnoreCase));Check(slide.Tags["PPI_TEMPLATE_PALETTE"]==TemplatePalettes.Items[source.Variant].Id,"Palette selection lost after reopen");string contentKey=source.Build().First(t=>t.Kind=="text"&&t.Key!="title"&&t.Key!="subtitle").Key;Check(Field(slide.Shapes,contentKey).TextFrame.TextRange.Text=="Editable content "+id.ToLowerInvariant(),"Grouped content edit lost after reopen");Check((slide.TimeLine.MainSequence.Count>0)==(source.Variant!=0),"Animations lost or added after reopen");Field(slide.Shapes,"subtitle").TextFrame.TextRange.Text="Still editable";}
            Console.WriteLine("PASS Office 520-template insertion, same presentation, editable titles and native animations after save/reopen.");
        }
        finally {if(owned!=null){owned.Saved=Office.MsoTriState.msoTrue;owned.Close();}if(app!=null&&app.Presentations.Count==0)app.Quit();}
    }
    internal static string Pixels(Bitmap bitmap)
    {
        // Compare decoded pixels, not a Graphics draw that can apply image DPI scaling.
        var bytes=new byte[8+bitmap.Width*bitmap.Height*4];Array.Copy(BitConverter.GetBytes(bitmap.Width),0,bytes,0,4);Array.Copy(BitConverter.GetBytes(bitmap.Height),0,bytes,4,4);int offset=8;
        for(int y=0;y<bitmap.Height;y++)for(int x=0;x<bitmap.Width;x++){var color=bitmap.GetPixel(x,y);bytes[offset++]=color.R;bytes[offset++]=color.G;bytes[offset++]=color.B;bytes[offset++]=color.A;}
        using(var hash=SHA256.Create())return BitConverter.ToString(hash.ComputeHash(bytes));
    }
    private static PowerPoint.Shape EditableText(PowerPoint.Shape shape)
    {
        if(shape.Type==Office.MsoShapeType.msoGroup){foreach(PowerPoint.Shape child in shape.GroupItems){var found=EditableText(child);if(found!=null)return found;}}
        if(shape.HasTextFrame==Office.MsoTriState.msoTrue&&shape.TextFrame.HasText==Office.MsoTriState.msoTrue)return shape;return null;
    }
    private static string RenderHash(PowerPoint.Slide slide,string file,int height)
    {slide.Export(file,"PNG",720,height);using(var bitmap=new Bitmap(file))return Pixels(bitmap);}
    private static string Geometry(PowerPoint.Slide slide,string output)
    {
        var rows=new List<string>();Action<PowerPoint.Shape,string> visit=null;visit=(shape,path)=>{string text="";if(shape.HasTextFrame==Office.MsoTriState.msoTrue&&shape.TextFrame.HasText==Office.MsoTriState.msoTrue)using(var hash=SHA256.Create())text=BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(shape.TextFrame.TextRange.Text)));rows.Add(path+"\t"+shape.Type+"\t"+String.Join("\t",new[]{shape.Left,shape.Top,shape.Width,shape.Height,shape.Rotation}.Select(x=>x.ToString("R",CultureInfo.InvariantCulture)))+"\t"+text);if(shape.Type==Office.MsoShapeType.msoGroup)for(int i=1;i<=shape.GroupItems.Count;i++)visit(shape.GroupItems[i],path+"."+i);};for(int i=1;i<=slide.Shapes.Count;i++)visit(slide.Shapes[i],i.ToString(CultureInfo.InvariantCulture));File.WriteAllLines(output,rows);return String.Join("\n",rows);
    }
    private static void OriginalRoundtrip(string output)
    {
        string root=TemplateSources.InstalledOriginals;if(root==null)throw new InvalidOperationException("Installed original template files were not found; source fidelity was not tested.");
        var files=new[]{Path.Combine(root,"Diagrams","Cubes","Cubes.potx"),Path.Combine(root,"Diagrams","Pyramids, Core Elements","Pyramids, Core Elements.potx")};var assets=files.SelectMany(f=>LocalSlideIndex.ReadFile(f)).ToArray();
        var hashes=files.ToDictionary(f=>f,f=>{using(var input=File.OpenRead(f))using(var sha=SHA256.Create())return BitConverter.ToString(sha.ComputeHash(input));});PowerPoint.Application app=null;PowerPoint.Presentation deck=null;
        var pixelHashes=new List<string>();var nativeGeometry=new List<string>();var nativeFormatting=new List<string>();var sourceEffects=new List<int>();var heights=new List<int>();var fidelityFailures=new List<string>();var fidelityRows=new List<string>{"Index,Stage,ExactPixels,ExactGeometry,ExactNativeFormatting,RepeatSourcePixels,RepeatSourceWithinTolerance,WithinTolerance,ChangedPixels,MaxChannelDelta,AlphaChanges,NonEdgeChanges,AllowedPixels"};
        try{
            app=new PowerPoint.Application();app.Visible=Office.MsoTriState.msoTrue;if(app.Presentations.Count!=0)throw new InvalidOperationException("Save and close user presentations and previous test fixtures before original-source validation.");
            deck=app.Presentations.Add(Office.MsoTriState.msoTrue);deck.PageSetup.SlideWidth=assets[0].Width;deck.PageSetup.SlideHeight=assets[0].Height;deck.Slides.Add(1,PowerPoint.PpSlideLayout.ppLayoutBlank).Select();
            Console.WriteLine("Original fidelity runner bitness="+(Environment.Is64BitProcess?"x64":"x86")+"; Office="+app.Version);
            for(int i=0;i<assets.Length;i++){
                deck.Slides[deck.Slides.Count].Select();var item=assets[i];Console.WriteLine("STEP source-original "+(i+1)+"/"+assets.Length+" "+Path.GetFileName(item.File)+" slide "+item.Slide);
                using(var session=new SourceTemplateSession(app)){
                    using(var expected=session.Preview(item)){pixelHashes.Add(Pixels(expected));expected.Save(Path.Combine(output,"original-"+i+"-source.png"),ImageFormat.Png);heights.Add(expected.Height);}
                    var source=session.Source(item);string geometry=Geometry(source,Path.Combine(output,"original-"+i+"-source-geometry.tsv"));nativeGeometry.Add(geometry);nativeFormatting.Add(TemplateFidelity.Formatting(source,Path.Combine(output,"original-"+i+"-source-format.txt")));int shapeCount=source.Shapes.Count,effects=source.TimeLine.MainSequence.Count;var transition=source.SlideShowTransition.EntryEffect;sourceEffects.Add(effects);
                    var slide=session.InsertOriginal(item);Check(deck.Slides.Count==i+2,"Original inserted in wrong destination");Check(slide.Shapes.Count==shapeCount&&slide.TimeLine.MainSequence.Count==effects&&slide.SlideShowTransition.EntryEffect==transition,"Source objects, builds or transition changed");
                    bool exactPixels=RenderHash(slide,Path.Combine(output,"original-"+i+"-inserted.png"),heights[i])==pixelHashes[i];bool exactGeometry=Geometry(slide,Path.Combine(output,"original-"+i+"-inserted-geometry.tsv"))==geometry;bool exactFormat=TemplateFidelity.Formatting(slide,Path.Combine(output,"original-"+i+"-inserted-format.txt"))==nativeFormatting[i];bool repeatSource=RenderHash(session.Source(item),Path.Combine(output,"original-"+i+"-source-repeat.png"),heights[i])==pixelHashes[i];var raster=TemplateFidelity.Files(Path.Combine(output,"original-"+i+"-source.png"),Path.Combine(output,"original-"+i+"-inserted.png"));var repeatRaster=TemplateFidelity.Files(Path.Combine(output,"original-"+i+"-source.png"),Path.Combine(output,"original-"+i+"-source-repeat.png"));fidelityRows.Add(i+",insert,"+exactPixels+","+exactGeometry+","+exactFormat+","+repeatSource+","+repeatRaster.Accepted+","+raster.Accepted+","+raster.ChangedPixels+","+raster.MaxChannelDelta+","+raster.AlphaChanges+","+raster.NonEdgeChanges+","+raster.AllowedPixels);File.WriteAllLines(Path.Combine(output,"original-fidelity.csv"),fidelityRows);if(!raster.Accepted||!repeatRaster.Accepted||!exactGeometry||!exactFormat)fidelityFailures.Add("Insertion fidelity "+i+": exactPixels="+exactPixels+", tolerance="+raster.Accepted+", geometry="+exactGeometry+", formatting="+exactFormat+", repeated source="+repeatSource);Check(slide.Tags["PPI_LIBRARY_FORMATTING"]=="KEEP_SOURCE","Source provenance missing");
                }
                Check(app.Presentations.Count==1,"Owned hidden source was not closed");Console.WriteLine((fidelityFailures.Any(f=>f.StartsWith("Insertion fidelity "+i+":"))?"FAIL":"PASS")+" original fidelity and owned-source cleanup "+i);
            }
            string path=Path.Combine(output,"originals-native-roundtrip.pptx");deck.SaveAs(path,PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);deck.Close();deck=null;deck=app.Presentations.Open(path);
            for(int i=0;i<assets.Length;i++){var slide=deck.Slides[i+2];bool exact=RenderHash(slide,Path.Combine(output,"original-"+i+"-reopened.png"),heights[i])==pixelHashes[i];var raster=TemplateFidelity.Files(Path.Combine(output,"original-"+i+"-source.png"),Path.Combine(output,"original-"+i+"-reopened.png"));bool geometry=Geometry(slide,Path.Combine(output,"original-"+i+"-reopened-geometry.tsv"))==nativeGeometry[i],format=TemplateFidelity.Formatting(slide,Path.Combine(output,"original-"+i+"-reopened-format.txt"))==nativeFormatting[i];fidelityRows.Add(i+",reopen,"+exact+","+geometry+","+format+",,,"+raster.Accepted+","+raster.ChangedPixels+","+raster.MaxChannelDelta+","+raster.AlphaChanges+","+raster.NonEdgeChanges+","+raster.AllowedPixels);File.WriteAllLines(Path.Combine(output,"original-fidelity.csv"),fidelityRows);if(!raster.Accepted||!geometry||!format)fidelityFailures.Add("Reopen fidelity "+i+": tolerance="+raster.Accepted+", geometry="+geometry+", formatting="+format);Check(slide.TimeLine.MainSequence.Count==sourceEffects[i]&&slide.Tags["PPI_LIBRARY_ASSET_VERSION"]==assets[i].Signature,"Original builds/provenance lost on reopen");}
            var text=deck.Slides[2].Shapes.Cast<PowerPoint.Shape>().Select(EditableText).FirstOrDefault(s=>s!=null);Check(text!=null,"No native editable text found");text.TextFrame.TextRange.Text="Original text edit survives reopen";deck.Save();deck.Close();deck=null;deck=app.Presentations.Open(path);Check(deck.Slides[2].Shapes.Cast<PowerPoint.Shape>().Select(EditableText).Any(s=>s!=null&&s.TextFrame.TextRange.Text=="Original text edit survives reopen"),"Native text edit did not persist");
            foreach(string file in files)using(var input=File.OpenRead(file))using(var sha=SHA256.Create())Check(BitConverter.ToString(sha.ComputeHash(input))==hashes[file],"Read-only original source changed");
            Check(fidelityFailures.Count==0,"Original fidelity incomplete: "+String.Join("; ",fidelityFailures));Console.WriteLine("PASS original-source owner-approved bounded raster comparison and exact native geometry/formatting before/after insertion and reopen, editable text, native builds, metadata, source hashes; slides="+assets.Length);
        }finally{if(deck!=null){deck.Saved=Office.MsoTriState.msoTrue;deck.Close();}if(app!=null&&app.Presentations.Count==0)app.Quit();}
    }

}
