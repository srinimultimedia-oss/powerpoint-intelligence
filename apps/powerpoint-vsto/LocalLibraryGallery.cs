using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Threading.Tasks;
using PowerPointIntelligence.Library;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
namespace PowerPointIntelligence.AddIn
{
    internal static class LibraryGallery
    {
        internal const int PageSize=24;
        internal static string Caption(LocalSlideAsset asset)
        {string name=Path.GetFileNameWithoutExtension(asset.File);return (name.Length>46?name.Substring(0,43)+"…":name)+" · "+asset.Slide.ToString(CultureInfo.InvariantCulture);}
        internal static IEnumerable<LocalSlideAsset> Page(IList<LocalSlideAsset> assets,int page)
        {if(page<0)throw new ArgumentOutOfRangeException(nameof(page));return assets.Skip(page*PageSize).Take(PageSize);}
    }
    internal sealed partial class LocalLibraryForm:Form
    {
        private readonly PowerPoint.Application app;
        private readonly SourceTemplateSession sourceSession;
        private readonly Func<LocalSlideAsset,Bitmap> previewProvider;
        private readonly bool originalMode;
        private readonly System.Threading.CancellationTokenSource cancellation=new System.Threading.CancellationTokenSource();
        private readonly List<LocalSlideAsset> assets=new List<LocalSlideAsset>();
        private List<LocalSlideAsset> filtered=new List<LocalSlideAsset>();
        private readonly TextBox folder=new TextBox {Width=430,ReadOnly=true};
        private readonly TextBox query=new TextBox {Name="librarySearch",Width=220,AccessibleName="Search slide text and filenames"},author=new TextBox {Width=110,AccessibleName="Author filter"},layout=new TextBox {Width=120,AccessibleName="Layout or master filter"};
        private readonly ComboBox aspect=new ComboBox {DropDownStyle=ComboBoxStyle.DropDownList,Width=80},sort=new ComboBox {DropDownStyle=ComboBoxStyle.DropDownList,Width=130},size=new ComboBox {DropDownStyle=ComboBoxStyle.DropDownList,Width=90,AccessibleName="Thumbnail size"};
        private readonly CheckBox duplicates=new CheckBox {Name="libraryHideDuplicates",Text="Hide exact duplicates",AutoSize=true},destination=new CheckBox {Text="Use destination design",AutoSize=true},recursive=new CheckBox {Text="Include subfolders",AutoSize=true},remember=new CheckBox {Text="Remember index locally",AutoSize=true};
        private readonly DateTimePicker since=new DateTimePicker {ShowCheckBox=true,Checked=false,Width=150};
        private readonly ListView gallery=new ListView {Name="slideGallery",AccessibleName="Slide template thumbnails",Dock=DockStyle.Fill,View=View.LargeIcon,MultiSelect=false,HideSelection=false,ShowItemToolTips=true,LabelWrap=true,BackColor=Color.White,BorderStyle=BorderStyle.None};
        private readonly ImageList images=new ImageList {ColorDepth=ColorDepth.Depth32Bit,ImageSize=new Size(208,129)};
        private readonly TreeView categories=new TreeView {Dock=DockStyle.Left,Width=230,HideSelection=false,BorderStyle=BorderStyle.FixedSingle};
        private readonly PictureBox largePreview=new PictureBox {Dock=DockStyle.Top,Height=235,SizeMode=PictureBoxSizeMode.Zoom,BackColor=Color.WhiteSmoke};
        private readonly RichTextBox details=new RichTextBox {Dock=DockStyle.Fill,ReadOnly=true,BorderStyle=BorderStyle.None};
        private readonly Panel detailPanel=new Panel {Dock=DockStyle.Right,Width=350,Padding=new Padding(8)};
        private readonly FlowLayoutPanel toolbar=new FlowLayoutPanel {Dock=DockStyle.Top,Height=105,Padding=new Padding(8),AutoScroll=true};
        private readonly Label status=new Label {Name="libraryStatus",Dock=DockStyle.Bottom,Height=42,Padding=new Padding(8)},pageLabel=new Label {AutoSize=true,Padding=new Padding(8)};
        private readonly Timer previewTimer=new Timer {Interval=80};
        private readonly Queue<ListViewItem> pending=new Queue<ListViewItem>();
        private readonly Dictionary<string,Bitmap> cache=new Dictionary<string,Bitmap>();
        private readonly Queue<string> cacheOrder=new Queue<string>();
        private string selectedCategory="";
        private int page;
        private bool indexing,rendering,changingPage,selectedPreviewPending,ready;
        private LocalSlideAsset SelectedAsset {get{return gallery.SelectedItems.Count==0?null:gallery.SelectedItems[0].Tag as LocalSlideAsset;}}
        internal LocalLibraryForm(PowerPoint.Application app,string initialFolder=null,bool originals=false,Func<LocalSlideAsset,Bitmap> previewProvider=null,PiidxCatalog catalogOverride=null,LibraryRuntimeProfile profileOverride=null)
        {
            workspaceToken=cancellation.Token;workspaceMode=app!=null||catalogOverride!=null;catalog=catalogOverride;profile=profileOverride;
            this.app=app;originalMode=originals;sourceSession=app==null?null:new SourceTemplateSession(app);this.previewProvider=previewProvider??(sourceSession==null?(Func<LocalSlideAsset,Bitmap>)null:sourceSession.Preview);
            Text=originals?"Original template gallery · PowerPoint Intelligence":"Slide template gallery · PowerPoint Intelligence";Width=1440;Height=860;MinimumSize=new Size(1080,650);StartPosition=FormStartPosition.CenterScreen;Font=new Font("Segoe UI",9);
            gallery.LargeImageList=images;detailPanel.Controls.Add(details);detailPanel.Controls.Add(largePreview);
            var browse=new Button {Text="Choose folder",AutoSize=true};var refresh=new Button {Text="Refresh",AutoSize=true};var insert=new Button {Text="Insert selected slide",AutoSize=true};
            aspect.Items.AddRange(new object[]{"All","16:9","4:3"});aspect.SelectedIndex=0;sort.Items.AddRange(new object[]{"File / slide","Title","Newest first"});sort.SelectedIndex=0;size.Items.AddRange(new object[]{"Small","Medium","Large"});size.SelectedIndex=1;
            toolbar.Controls.AddRange(new Control[]{folder,browse,refresh,new Label {Text="Search",AutoSize=true},query,new Label {Text="Author",AutoSize=true},author,new Label {Text="Layout/master",AutoSize=true},layout,aspect,since,duplicates,sort,recursive,destination,remember,insert});
            var controls=new FlowLayoutPanel {Dock=DockStyle.Top,Height=40,Padding=new Padding(6,2,6,2)};var previous=new Button {Name="previousGalleryPage",Text="Previous",AutoSize=true};var next=new Button {Name="nextGalleryPage",Text="Next",AutoSize=true};var showDetails=new CheckBox {Text="Large preview",Checked=true,AutoSize=true,Padding=new Padding(8)};
            controls.Controls.AddRange(new Control[]{previous,pageLabel,next,new Label {Text="Thumbnail size",AutoSize=true,Padding=new Padding(10,7,0,0)},size,showDetails});
            var center=new Panel {Dock=DockStyle.Fill,Padding=new Padding(8)};center.Controls.Add(gallery);center.Controls.Add(controls);Controls.Add(center);Controls.Add(categories);Controls.Add(detailPanel);Controls.Add(status);Controls.Add(toolbar);
            browse.Click+=(s,e)=>{using(var picker=new FolderBrowserDialog {Description="Select the folder containing original .pptx/.potx files."})if(picker.ShowDialog()==DialogResult.OK){folder.Text=picker.SelectedPath;RegisterWorkspaceFolder(picker.SelectedPath);}};refresh.Click+=(s,e)=>RefreshIndex();insert.Click+=(s,e)=>Insert();gallery.ItemActivate+=(s,e)=>Insert();
            query.TextChanged+=(s,e)=>Filter();author.TextChanged+=(s,e)=>Filter();layout.TextChanged+=(s,e)=>Filter();aspect.SelectedIndexChanged+=(s,e)=>Filter();sort.SelectedIndexChanged+=(s,e)=>Filter();since.ValueChanged+=(s,e)=>Filter();duplicates.CheckedChanged+=(s,e)=>Filter();recursive.CheckedChanged+=(s,e)=>{if(ready&&!rebinding)RefreshIndex();};
            categories.AfterSelect+=(s,e)=>{selectedCategory=Convert.ToString(e.Node.Tag,CultureInfo.InvariantCulture);if(!rebinding)Filter();};previous.Click+=(s,e)=>MovePage(-1);next.Click+=(s,e)=>MovePage(1);showDetails.CheckedChanged+=(s,e)=>detailPanel.Visible=showDetails.Checked;
            size.SelectedIndexChanged+=(s,e)=>{int width=size.SelectedIndex==0?160:size.SelectedIndex==2?256:208;images.ImageSize=new Size(width,width*9/16+12);ShowPage();};
            gallery.SelectedIndexChanged+=(s,e)=>{if(changingPage)return;UpdateDetails();selectedPreviewPending=SelectedAsset!=null;previewTimer.Start();};previewTimer.Tick+=(s,e)=>RenderNextPreview();
            remember.Enabled=false;remember.Text="PIIDX catalog";remember.Checked=true;
            FormClosing+=(s,e)=>{if(rendering){e.Cancel=true;status.Text="Finishing the current preview. Close again in a moment.";return;}previewTimer.Stop();pending.Clear();try{if(sourceSession!=null)sourceSession.Dispose();}catch(System.Runtime.InteropServices.COMException ex){e.Cancel=true;status.Text="Cannot close the owned hidden source (0x"+ex.HResult.ToString("X8")+"). Retry closing the gallery.";}};
            FormClosed+=(s,e)=>cancellation.Cancel();
            if(initialFolder!=null){folder.Text=initialFolder;recursive.Checked=true;}
            if(originals){destination.Checked=false;destination.Enabled=false;}
            ConfigureWorkspace();ready=true;Shown+=(s,e)=>{if(workspaceMode)InitializeWorkspace(initialFolder);else if(pending.Count>0)previewTimer.Start();};

        }
        internal void ShowAssets(string root,IEnumerable<LocalSlideAsset> items)
        {
            // Create the native list before populating it; parent-handle creation can raise tree selection events.
            var nativeHandle=gallery.Handle;
            injected=true;folder.Text=root;assets.Clear();assets.AddRange(items);selectedCategory="";categories.Nodes.Clear();var all=categories.Nodes.Add("All original files ("+assets.Count+")");all.Tag="";
            var nodes=new Dictionary<string,TreeNode>(StringComparer.OrdinalIgnoreCase);nodes[Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar)]=all;
            foreach(var asset in assets){string cursor=Path.GetDirectoryName(asset.File);var parts=new Stack<string>();while(!nodes.ContainsKey(cursor)){parts.Push(cursor);cursor=Path.GetDirectoryName(cursor);if(cursor==null)break;}while(parts.Count>0){string path=parts.Pop();TreeNode parent;var node=(nodes.TryGetValue(Path.GetDirectoryName(path),out parent)?parent:all).Nodes.Add(Path.GetFileName(path));node.Tag=path;nodes[path]=node;}}
            foreach(var pair in nodes.Where(p=>p.Value!=all))pair.Value.Text=Path.GetFileName(pair.Key)+" ("+assets.Count(a=>LocalSlideIndex.IsWithin(pair.Key,a.File))+")";all.Expand();categories.SelectedNode=all;Filter();
        }
        private void Filter()
        {
            if(workspaceMode&&!injected){LoadWorkspacePage(true);return;}
            var matches=LocalSlideIndex.Filter(assets.Where(a=>String.IsNullOrEmpty(selectedCategory)||LocalSlideIndex.IsWithin(selectedCategory,a.File)),query.Text,author.Text,layout.Text,(string)aspect.SelectedItem,since.Checked?(DateTime?)since.Value.Date:null,duplicates.Checked);
            if(sort.SelectedIndex==1)matches=matches.OrderBy(a=>a.Title,StringComparer.CurrentCultureIgnoreCase);else if(sort.SelectedIndex==2)matches=matches.OrderByDescending(a=>a.Modified);else matches=matches.OrderBy(a=>a.File,StringComparer.CurrentCultureIgnoreCase).ThenBy(a=>a.Slide);filtered=matches.ToList();page=0;ShowPage();
        }
        internal void MovePage(int delta)
        {int next=page+delta;if(next<0||next*LibraryGallery.PageSize>=(workspaceMode&&!injected?totalMatched:filtered.Count))return;page=next;if(workspaceMode&&!injected)LoadWorkspacePage(false);else ShowPage();}
        private string Key(LocalSlideAsset asset){return asset.Signature+"@"+images.ImageSize.Width.ToString(CultureInfo.InvariantCulture);}
        private Bitmap Frame(Bitmap slide,string message=null)
        {
            var result=new Bitmap(images.ImageSize.Width,images.ImageSize.Height);using(var g=Graphics.FromImage(result)){
                g.Clear(Color.FromArgb(243,246,250));g.FillRectangle(Brushes.White,3,3,result.Width-7,result.Height-7);
                if(slide!=null){float scale=Math.Min((result.Width-12)/(float)slide.Width,(result.Height-12)/(float)slide.Height);float w=slide.Width*scale,h=slide.Height*scale;g.InterpolationMode=InterpolationMode.HighQualityBicubic;g.DrawImage(slide,new RectangleF((result.Width-w)/2,(result.Height-h)/2,w,h),new RectangleF(0,0,slide.Width,slide.Height),GraphicsUnit.Pixel);}
                else using(var font=new Font("Segoe UI",9))using(var format=new StringFormat {Alignment=StringAlignment.Center,LineAlignment=StringAlignment.Center})g.DrawString(message??"Loading preview…",font,Brushes.SlateGray,new RectangleF(8,8,result.Width-16,result.Height-16),format);
                using(var pen=new Pen(Color.FromArgb(213,221,232)))g.DrawRectangle(pen,3,3,result.Width-7,result.Height-7);
            }return result;
        }
        private void ShowPage()
        {
            previewTimer.Stop();pending.Clear();changingPage=true;gallery.BeginUpdate();
            try{gallery.Items.Clear();images.Images.Clear();using(var placeholder=Frame(null))images.Images.Add("loading",placeholder);foreach(var asset in workspaceMode&&!injected?filtered:LibraryGallery.Page(filtered,page)){string key=Key(asset);Bitmap cached;if(cache.TryGetValue(key,out cached)&&!images.Images.ContainsKey(key))images.Images.Add(key,cached);var item=new ListViewItem(LibraryGallery.Caption(asset)) {Tag=asset,ImageKey=images.Images.ContainsKey(key)?key:"loading",ToolTipText=asset.Title+"\n"+asset.File+"\nDouble-click to insert this slide."};gallery.Items.Add(item);if(item.ImageKey=="loading")pending.Enqueue(item);}}
            finally{gallery.EndUpdate();changingPage=false;}
            var previous=Controls.Find("previousGalleryPage",true).FirstOrDefault();var next=Controls.Find("nextGalleryPage",true).FirstOrDefault();if(previous!=null)previous.Enabled=page>0;if(next!=null)next.Enabled=(page+1)*LibraryGallery.PageSize<(workspaceMode&&!injected?totalMatched:filtered.Count);
            long count=workspaceMode&&!injected?totalMatched:filtered.Count;
            pageLabel.Text=count==0?"No matching slides":(page+1)+" / "+Math.Ceiling(count/(double)LibraryGallery.PageSize)+" · "+count+" slides";
            if(gallery.Items.Count>0){gallery.Items[0].Selected=true;selectedPreviewPending=true;}else selectedPreviewPending=false;UpdateDetails();status.Text=count+" matching slides. Select a thumbnail to preview; double-click to insert.";if(pending.Count>0||selectedPreviewPending)previewTimer.Start();
        }
        private void UpdateDetails()
        {
            var item=SelectedAsset;details.Text=item==null?"Select a slide thumbnail.":LibraryGallery.Caption(item)+"\nAuthor: "+item.Author+"\nLayout: "+item.Layout+"\nModified (UTC): "+item.Modified.ToString("u")+"\n\n"+item.Text;
            if(largePreview.Image!=null){largePreview.Image.Dispose();largePreview.Image=null;}
        }
        private void RememberThumbnail(LocalSlideAsset asset,Bitmap bitmap)
        {
            string key=Key(asset);if(!cache.ContainsKey(key)){cache[key]=Frame(bitmap);cacheOrder.Enqueue(key);while(cache.Count>96){string oldest=cacheOrder.Dequeue();cache[oldest].Dispose();cache.Remove(oldest);}}
            if(!images.Images.ContainsKey(key))images.Images.Add(key,cache[key]);foreach(ListViewItem item in gallery.Items)if(Key((LocalSlideAsset)item.Tag)==key)item.ImageKey=key;
        }
        internal async void RenderNextPreview()
        {
            if(rendering||IsDisposed)return;previewTimer.Stop();rendering=true;LocalSlideAsset renderingAsset=null;
            try{
                if(selectedPreviewPending){selectedPreviewPending=false;var selected=SelectedAsset;renderingAsset=selected;if(selected!=null&&previewProvider!=null){var bitmap=await WorkspacePreview(selected);await WorkspaceUi();if(bitmap==null)return;if(SelectedAsset==selected){if(largePreview.Image!=null)largePreview.Image.Dispose();largePreview.Image=bitmap;RememberThumbnail(selected,bitmap);}else bitmap.Dispose();}return;}
                while(pending.Count>0){var item=pending.Dequeue();if(item.ListView!=gallery)continue;var asset=(LocalSlideAsset)item.Tag;renderingAsset=asset;string key=Key(asset);if(cache.ContainsKey(key)){item.ImageKey=key;continue;}if(previewProvider==null)break;using(var bitmap=await WorkspacePreview(asset)){await WorkspaceUi();if(bitmap!=null&&!IsDisposed)RememberThumbnail(asset,bitmap);}return;}
            }catch(OperationCanceledException){}catch(Exception ex) when(ex is IOException||ex is InvalidDataException||ex is UnauthorizedAccessException||ex is InvalidOperationException||ex is System.Runtime.InteropServices.COMException){await WorkspaceUi();if(IsDisposed)return;status.Text="Preview unavailable: "+ex.Message;if(renderingAsset!=null){string errorKey=Key(renderingAsset)+"-error";if(!images.Images.ContainsKey(errorKey))using(var failure=Frame(null,"Preview unavailable"))images.Images.Add(errorKey,failure);foreach(ListViewItem item in gallery.Items)if(Key((LocalSlideAsset)item.Tag)==Key(renderingAsset)){item.ImageKey=errorKey;item.ToolTipText="Preview unavailable: "+ex.Message;}}}
            finally{await WorkspaceUi();rendering=false;if(!IsDisposed&&(pending.Count>0||selectedPreviewPending))previewTimer.Start();}
        }
        private void RefreshIndex(){if(workspaceMode)RefreshWorkspace(false);}
        private async void Insert()
        {
            if(rendering){status.Text="Finishing the current preview. Try insertion again in a moment.";return;}var item=SelectedAsset;if(item==null||app==null)return;previewTimer.Stop();PowerPoint.Slide inserted=null;
            try{if(item.IndexedReference!=null){await Task.Run(()=>catalog.VerifySource(item.IndexedReference,workspaceToken));await WorkspaceUi();if(IsDisposed||SelectedAsset!=item)throw new InvalidOperationException("Selection changed before insertion.");}if(originalMode){sourceSession.InsertOriginal(item);status.Text="Inserted the original slide with its source design.";return;}var deck=sourceSession.Destination;var active=sourceSession.TargetSlide;if(item.IndexedReference==null){var current=LocalSlideIndex.ReadFile(item.File).Single(a=>a.Slide==item.Slide);if(current.Signature!=item.Signature)throw new InvalidOperationException("Source changed. Refresh before inserting.");}app.StartNewUndoEntry();if(deck.Slides.InsertFromFile(item.File,active.SlideIndex,item.Slide,item.Slide)!=1)throw new InvalidOperationException("PowerPoint did not insert exactly one slide.");inserted=deck.Slides[active.SlideIndex+1];inserted.Tags.Add("PPI_LIBRARY_ASSET_SCHEMA","1");inserted.Tags.Add("PPI_LIBRARY_ASSET_VERSION",item.Signature);inserted.Tags.Add("PPI_LIBRARY_SOURCE_SLIDE",item.Slide.ToString(CultureInfo.InvariantCulture));if(destination.Checked)inserted.Design=active.Design;inserted.Select();status.Text="Inserted into the current presentation.";}
            catch(Exception ex) when(ex is IOException||ex is InvalidDataException||ex is InvalidOperationException||ex is System.Runtime.InteropServices.COMException){await WorkspaceUi();if(IsDisposed)return;bool restored=true;if(inserted!=null)try{inserted.Delete();}catch{restored=false;}status.Text="Insertion failed: "+ex.Message+(restored?"":" Cleanup was incomplete. Use PowerPoint Undo.");}
            finally{if(pending.Count>0||selectedPreviewPending)previewTimer.Start();}
        }
        protected override void Dispose(bool disposing)
        {
            if(disposing){CloseWorkspace();previewTimer.Dispose();images.Dispose();if(largePreview.Image!=null){largePreview.Image.Dispose();largePreview.Image=null;}foreach(var bitmap in cache.Values)bitmap.Dispose();cache.Clear();cancellation.Dispose();}base.Dispose(disposing);
        }
    }
}
