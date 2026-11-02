using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using PowerPointIntelligence.AddIn;
using PowerPointIntelligence.Library;
internal static class PiidxGalleryTests
{
    static void Wait(Func<bool> ready,string name){DateTime deadline=DateTime.UtcNow.AddSeconds(15);while(DateTime.UtcNow<deadline){Application.DoEvents();if(ready())return;Thread.Sleep(10);}throw new Exception("PIIDX gallery timeout: "+name);}
    internal static void Run()
    {
        string output=Path.Combine("artifacts","demo-contracts","piidx-ui-"+Guid.NewGuid().ToString("N")),sources=Path.Combine(output,"sources");Directory.CreateDirectory(sources);
        for(int i=0;i<13;i++)PiidxSourceTests.Fixture(Path.Combine(sources,"source"+i.ToString("00")+".pptx"));
        var catalog=PiidxCatalog.OpenOrCreate(Path.Combine(output,"gallery.piidx"));string id=catalog.Register(sources,"UI fixture",true);catalog.Refresh(id,CancellationToken.None);int rendered=0;int uiThread=Thread.CurrentThread.ManagedThreadId;
        using(var form=new LocalLibraryForm(null,previewProvider:a=>{if(Thread.CurrentThread.ManagedThreadId!=uiThread)throw new Exception("Preview provider left the UI thread.");rendered++;return new Bitmap(720,405);},catalogOverride:catalog))
        {
            form.Show();var list=(ListView)form.Controls.Find("slideGallery",true).Single();Wait(()=>list.Items.Count==24,"first indexed page");
            if(!form.Controls.Find("nextGalleryPage",true).Single().Enabled)throw new Exception("PIIDX next page disabled.");
            form.MovePage(1);Wait(()=>list.Items.Count==2,"second indexed page");
            var asset=(LocalSlideAsset)list.Items[0].Tag;if(asset.IndexedReference==null)throw new Exception("Indexed source identity was not bound to the gallery row.");
            var query=(TextBox)form.Controls.Find("librarySearch",true).Single();query.Text="source01";Wait(()=>list.Items.Count==2&&((LocalSlideAsset)list.Items[0].Tag).File.EndsWith("source01.pptx"),"SQL filename filtering");
            query.Text="term_not_in_the_index";Wait(()=>list.Items.Count==0,"empty indexed result");query.Text="";Wait(()=>list.Items.Count==24,"return to page one");
            try{Wait(()=>rendered>0,"preview provider invoked on UI thread");}catch(Exception ex){throw new Exception(ex.Message+"; status="+form.Controls.Find("libraryStatus",true).Single().Text+"; rendering="+typeof(LocalLibraryForm).GetField("rendering",System.Reflection.BindingFlags.NonPublic|System.Reflection.BindingFlags.Instance).GetValue(form)+"; state="+String.Join(";",new[]{"previewProvider","previewTimer","pending","selectedPreviewPending","indexing"}.Select(n=>{var value=typeof(LocalLibraryForm).GetField(n,System.Reflection.BindingFlags.NonPublic|System.Reflection.BindingFlags.Instance).GetValue(form);var timer=value as System.Windows.Forms.Timer;var queue=value as System.Collections.ICollection;return n+"="+(timer!=null?timer.Enabled.ToString():queue!=null?queue.Count.ToString():value==null?"null":value.ToString());})),ex);}
            string legacy=Path.Combine(Path.GetDirectoryName(catalog.IndexPath),"PowerPointIntelligence","library-index-v1.xml");if(File.Exists(legacy))throw new Exception("UI wrote an obsolete XML index.");
            form.Hide();
        }
        Console.WriteLine("PASS 8 integrated PIIDX gallery assertions: native WinForms catalog binding, identity, paging, query/empty states and preview dispatch; preview pixels are test-provider output, not Office rendering evidence.");
    }
}