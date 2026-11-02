using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using PowerPointIntelligence.AddIn;
internal static class GalleryTests
{
    private static int passed;
    private static void Check(bool value,string label){if(!value)throw new Exception("Gallery: "+label);passed++;Console.WriteLine("PASS Gallery "+label);}
    internal static void Run()
    {
        string assembly=Path.Combine(Path.GetTempPath(),"ppi-library-origin"),local=Path.Combine(assembly,"templates"),vendor=Path.Combine(Path.GetTempPath(),"ppi-vendor-library"),copied=Path.Combine(Path.GetTempPath(),"ppi-copied-library");
        Check(TemplateSources.ResolveFolder(assembly,copied,vendor,p=>p==copied||p==vendor)==copied,"configured local copy wins over vendor install");Check(TemplateSources.ResolveFolder(assembly,"templates",vendor,p=>p==local)==local,"relative template reference resolves beside the assembly");Check(TemplateSources.ResolveFolder(assembly,null,vendor,p=>p==local||p==vendor)==local,"adjacent template folder wins without explicit configuration");Check(TemplateSources.ResolveFolder(assembly,null,vendor,p=>p==vendor)==null,"unconfigured runtime never discovers a vendor installation");Check(TemplateSources.ResolveFolder(assembly,null,vendor,p=>false)==null,"absent unconfigured source permits procedural gallery");bool missing=false;try{TemplateSources.ResolveFolder(assembly,copied,vendor,p=>p==vendor);}catch(DirectoryNotFoundException){missing=true;}Check(missing,"missing explicit local copy never silently falls back to vendor");
        Check(TemplateSources.ReadFolderSetting("<configuration><appSettings><add key='TemplateLibraryFolder' value='templates'/></appSettings></configuration>")=="templates","runtime config reads template folder setting");bool duplicate=false;try{TemplateSources.ReadFolderSetting("<configuration><appSettings><add key='TemplateLibraryFolder' value='a'/><add key='TemplateLibraryFolder' value='b'/></appSettings></configuration>");}catch(InvalidDataException){duplicate=true;}Check(duplicate,"ambiguous runtime template references rejected");bool dtd=false;try{TemplateSources.ReadFolderSetting("<!DOCTYPE configuration [<!ENTITY x SYSTEM 'file:///never-read'>]><configuration/>");}catch(System.Xml.XmlException){dtd=true;}Check(dtd,"runtime template setting prohibits DTD and external entities");
        string deployedConfig=typeof(TemplateSources).Assembly.Location+".config";if(File.Exists(deployedConfig)){string configured=TemplateSources.ReadFolderSetting(File.ReadAllText(deployedConfig));if(!String.IsNullOrWhiteSpace(configured)){string actual=TemplateSources.InstalledOriginals;Check(actual==Path.GetFullPath(configured),"compiled runtime loads the configured copied-library reference");var packages=LocalSlideIndex.Discover(actual,true);var slides=packages.SelectMany(LocalSlideIndex.ReadFile).ToArray();Check(packages.Length>0&&slides.Length>0&&slides.All(a=>LocalSlideIndex.IsWithin(actual,a.File)),"runtime gallery indexes templates from the configured local copy");Console.WriteLine("Runtime template source: packages="+packages.Length+"; slides="+slides.Length+"; configured local copy.");}}
        string root=Path.Combine(Path.GetTempPath(),"ppi-gallery-fixture");var items=Enumerable.Range(1,29).Select(i=>new LocalSlideAsset {File=Path.Combine(root,"Collection.potx"),Slide=i,Signature="gallery-fixture-"+i,Title="Fixture "+i,Text="Fixture",Author="",Layout="",Master="",Width=960,Height=540}).ToArray();
        Check(LibraryGallery.Page(items,0).Count()==24&&LibraryGallery.Page(items,1).Count()==5,"pages bounded at 24 cards");Check(LibraryGallery.Caption(items[0])=="Collection · 1","short source identity replaces full text dump");
        int rendered=0;using(var form=new LocalLibraryForm(null,previewProvider:asset=>{rendered++;var image=new Bitmap(160,90);using(var g=Graphics.FromImage(image))g.Clear(Color.SteelBlue);return image;})){
            form.ShowAssets(root,items);var list=(ListView)form.Controls.Find("slideGallery",true).Single();var handle=list.Handle;
            Check(list.View==View.LargeIcon&&list.LargeImageList!=null,"native large-image gallery wired");Check(list.Items.Count==24&&((LocalSlideAsset)list.Items[0].Tag).Slide==1,"first page binds correct assets");Check(!form.Controls.Find("previousGalleryPage",true).Single().Enabled,"previous disabled on first page");
            for(int i=0;i<LibraryGallery.PageSize+2;i++)form.RenderNextPreview();Check(rendered>0&&list.Items.Cast<ListViewItem>().Any(i=>i.ImageKey!="loading"),"queued previews populate real image slots");
            form.MovePage(1);Check(list.Items.Count==5&&((LocalSlideAsset)list.Items[0].Tag).Slide==25,"next page retains source identity");Check(!form.Controls.Find("nextGalleryPage",true).Single().Enabled,"next disabled on last page");form.MovePage(-1);Check(list.Items.Count==24,"previous returns first page");
            int before=rendered;for(int i=0;i<4;i++)form.RenderNextPreview();Check(rendered-before<4,"previous-page thumbnails reused from cache");
        }
        Console.WriteLine("PASS "+passed+" thumbnail gallery UI/paging/cache assertions; PowerPoint rendering not invoked.");
    }
    internal static void Snapshot(string output,string captures)
    {
        Application.EnableVisualStyles();
        string root=TemplateSources.InstalledOriginals;if(root==null)throw new Exception("Original source files absent");
        var items=LocalSlideIndex.ReadFile(Path.Combine(root,"Diagrams","Cubes","Cubes.potx")).Concat(LocalSlideIndex.ReadFile(Path.Combine(root,"Diagrams","Pyramids, Core Elements","Pyramids, Core Elements.potx")).Take(4)).ToArray();
        using(var form=new LocalLibraryForm(null,originals:true,previewProvider:asset=>new Bitmap(Path.Combine(captures,"original-"+Array.IndexOf(items,asset)+"-source.png")))){
            form.ShowInTaskbar=false;form.StartPosition=FormStartPosition.Manual;form.Location=new Point(-2400,-2400);form.ShowAssets(root,items);form.Show();Application.DoEvents();for(int i=0;i<items.Length+3;i++)form.RenderNextPreview();Application.DoEvents();using(var image=new Bitmap(form.Width,form.Height)){form.DrawToBitmap(image,new Rectangle(0,0,image.Width,image.Height));image.Save(output,System.Drawing.Imaging.ImageFormat.Png);}form.Close();
        }
        Console.WriteLine("Gallery UI snapshot uses seven captured native source renders; deployed PowerPoint UI not tested.");
    }
}
