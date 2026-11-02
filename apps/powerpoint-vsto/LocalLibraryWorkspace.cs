using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using PowerPointIntelligence.Library;
namespace PowerPointIntelligence.AddIn
{
    internal sealed partial class LocalLibraryForm
    {
        readonly int workspaceUiThread=Thread.CurrentThread.ManagedThreadId;
        struct UiAwaiter : INotifyCompletion
        {
            readonly LocalLibraryForm owner;
            internal UiAwaiter(LocalLibraryForm owner){this.owner=owner;}
            public UiAwaiter GetAwaiter(){return this;}
            public bool IsCompleted {get{return owner.IsDisposed||Thread.CurrentThread.ManagedThreadId==owner.workspaceUiThread;}}
            public void GetResult(){}
            public void OnCompleted(Action continuation)
            {
                try{if(!owner.IsDisposed&&owner.IsHandleCreated){owner.BeginInvoke(continuation);return;}}catch(InvalidOperationException){}
                ThreadPool.QueueUserWorkItem(_=>continuation());
            }
        }
        UiAwaiter WorkspaceUi(){return new UiAwaiter(this);}
        readonly bool workspaceMode;
        readonly CancellationToken workspaceToken;
        bool injected,rebinding,pendingWorkspaceReload;
        PiidxCatalog catalog;
        LibraryRuntimeProfile profile;
        LibraryRefreshMonitor monitor;
        CancellationTokenSource workCancellation,viewCancellation;
        long totalMatched;
        int viewRevision;
        readonly ComboBox libraries=new ComboBox{DropDownStyle=ComboBoxStyle.DropDownList,Width=180,Name="librarySelector"};
        readonly ComboBox presentations=new ComboBox{DropDownStyle=ComboBoxStyle.DropDownList,Width=180,Name="presentationSelector"};
        readonly ComboBox searchMode=new ComboBox{DropDownStyle=ComboBoxStyle.DropDownList,Width=100,Name="librarySearchMode"};
        readonly TextBox tagFilter=new TextBox{Width=150,AccessibleName="Exact indexed tag filter"};
        readonly Button cancelIndex=new Button{Text="Cancel indexing",AutoSize=true,Enabled=false};
        readonly System.Windows.Forms.Timer workspaceTimer=new System.Windows.Forms.Timer{Interval=500};
        sealed class LibraryChoice
        {
            internal LibraryRegistration Registration;
            public override string ToString(){return Registration==null?"All libraries":Registration.Name;}
        }
        string SelectedLibraryId {get{var item=libraries.SelectedItem as LibraryChoice;return item==null||item.Registration==null?null:item.Registration.Id;}}
        bool Managed(string root){return profile!=null&&profile.Sources.Any(s=>s.Managed&&String.Equals(Path.GetFullPath(s.Root).TrimEnd('\\'),Path.GetFullPath(root).TrimEnd('\\'),StringComparison.OrdinalIgnoreCase));}
        void ConfigureWorkspace()
        {
            if(!workspaceMode)return;toolbar.Height=145;sort.Items.Add("Relevance");sort.SelectedIndex=3;
            searchMode.Items.AddRange(Enum.GetNames(typeof(LibrarySearchMode)));searchMode.SelectedIndex=0;
            var remove=new Button{Text="Remove library",AutoSize=true,Name="removeLibrary"};var rebuild=new Button{Text="Rebuild",AutoSize=true};
            var designed=new Button{Text="Designed templates",AutoSize=true};designed.Click+=(s,e)=>{if(app!=null)using(var templates=new TemplateLibrary(app))templates.ShowDialog(this);};
            toolbar.Controls.Add(designed);toolbar.Controls.AddRange(new Control[]{libraries,presentations,searchMode,new Label{Text="Tag",AutoSize=true},tagFilter,remove,rebuild});
            var container=Controls.Find("previousGalleryPage",true).First().Parent;container.Controls.Add(cancelIndex);
            cancelIndex.Click+=(s,e)=>{if(workCancellation!=null)workCancellation.Cancel();};
            rebuild.Click+=(s,e)=>RefreshWorkspace(true);remove.Click+=(s,e)=>RemoveWorkspaceLibrary();
            libraries.SelectedIndexChanged+=async(s,e)=>{if(rebinding||catalog==null)return;await WorkspaceNavigation();await WorkspaceUi();if(!IsDisposed)Filter();};
            presentations.SelectedIndexChanged+=(s,e)=>{if(!rebinding)Filter();};searchMode.SelectedIndexChanged+=(s,e)=>Filter();tagFilter.TextChanged+=(s,e)=>Filter();
            workspaceTimer.Tick+=async(s,e)=>{if(pendingWorkspaceReload&&!indexing&&!rendering&&catalog!=null){pendingWorkspaceReload=false;await WorkspaceNavigation();await WorkspaceUi();if(!IsDisposed)Filter();}};workspaceTimer.Start();
        }
        async void InitializeWorkspace(string initialFolder)
        {
            indexing=true;toolbar.Enabled=false;status.Text="Opening PIIDX library catalog…";
            try
            {
                string assembly=typeof(TemplateSources).Assembly.Location;var existing=catalog;var settings=profile;
                var initialized=await Task.Run(()=>
                {
                    var selected=settings??LibraryRuntimeProfile.Load(existing==null?assembly+".config":Path.Combine(Path.GetDirectoryName(existing.IndexPath),"no-app.config"),Path.GetDirectoryName(assembly),existing==null?Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData):Path.GetDirectoryName(existing.IndexPath));
                    var index=existing??PiidxCatalog.OpenOrCreate(selected.IndexPath);var warnings=new List<string>();
                    foreach(var source in selected.Sources)
                    {
                        try{string id=index.Register(source.Root,source.Name,source.Recursive);if(!index.Libraries().Single(l=>l.Id==id).LastRefreshUtc.HasValue)index.Refresh(id,workspaceToken);}
                        catch(Exception ex)when(ex is IOException||ex is UnauthorizedAccessException||ex is System.Xml.XmlException){warnings.Add("Unavailable source: "+source.Root+" ("+ex.GetType().Name+")");}
                    }
                    if(initialFolder!=null&&!selected.Sources.Any(s=>String.Equals(Path.GetFullPath(s.Root),Path.GetFullPath(initialFolder),StringComparison.OrdinalIgnoreCase))){string id=index.Register(initialFolder,Path.GetFileName(initialFolder.TrimEnd('\\')),true);if(!index.Libraries().Single(l=>l.Id==id).LastRefreshUtc.HasValue)index.Refresh(id,workspaceToken);}
                    try{selected.MigrateLegacy(index,workspaceToken);}catch(Exception ex)when(ex is IOException||ex is UnauthorizedAccessException||ex is System.Xml.XmlException){warnings.Add("Legacy registration preserved; migration unavailable ("+ex.GetType().Name+").");}
                    return Tuple.Create(index,selected,warnings);
                });
                await WorkspaceUi();if(IsDisposed||Disposing)return;catalog=initialized.Item1;profile=initialized.Item2;
                await WorkspaceNavigation();await WorkspaceUi();if(IsDisposed)return;monitor=new LibraryRefreshMonitor(catalog,profile.RefreshSeconds,profile.WatchRootLimit,notice=>
                {
                    if(IsDisposed||Disposing||!IsHandleCreated)return;
                    try{BeginInvoke((Action)(()=>{if(IsDisposed)return;if(notice.ErrorCode!=null)status.Text="Automatic refresh unavailable ("+notice.ErrorCode+"). Previous catalog retained.";else if(notice.Result!=null&&(notice.Result.FilesUpdated>0||notice.Result.MissingFiles>0))pendingWorkspaceReload=true;}));}catch(InvalidOperationException){}
                });
                monitor.RequestAll();Filter();if(initialized.Item3.Count>0)status.Text=String.Join(" · ",initialized.Item3);
            }
            catch(OperationCanceledException){}
            catch(Exception ex)when(ex is IOException||ex is UnauthorizedAccessException||ex is System.Xml.XmlException||ex is ArgumentException){await WorkspaceUi();if(!IsDisposed)status.Text="PIIDX library unavailable: "+ex.Message;}
            finally{await WorkspaceUi();indexing=false;if(!IsDisposed)toolbar.Enabled=true;}
        }
        async Task WorkspaceNavigation(string selectId=null)
        {
            if(catalog==null)return;string chosen=selectId??SelectedLibraryId;
            var data=await Task.Run(()=>Tuple.Create(catalog.Libraries(),catalog.Categories(chosen,workspaceToken),catalog.Presentations(chosen,workspaceToken)));
            await WorkspaceUi();if(IsDisposed||Disposing)return;rebinding=true;
            try
            {
                libraries.Items.Clear();libraries.Items.Add(new LibraryChoice());foreach(var library in data.Item1)libraries.Items.Add(new LibraryChoice{Registration=library});
                libraries.SelectedIndex=chosen==null?0:Math.Max(0,libraries.Items.Cast<LibraryChoice>().ToList().FindIndex(l=>l.Registration!=null&&l.Registration.Id==chosen));
                var selected=libraries.SelectedItem as LibraryChoice;folder.Text=selected==null||selected.Registration==null?catalog.IndexPath:selected.Registration.Root;recursive.Enabled=selected!=null&&selected.Registration!=null&&!Managed(selected.Registration.Root);if(selected!=null&&selected.Registration!=null)recursive.Checked=selected.Registration.Recursive;
                presentations.Items.Clear();presentations.Items.Add("All presentations");foreach(var presentation in data.Item3)presentations.Items.Add(presentation);presentations.SelectedIndex=0;
                selectedCategory="";categories.Nodes.Clear();var all=categories.Nodes.Add("All indexed slides");all.Tag="";var nodes=new Dictionary<string,TreeNode>{{"",all}};
                foreach(var group in data.Item2){string path="";TreeNode parent=all;foreach(string part in group.Category.Split(new[]{'/'},StringSplitOptions.RemoveEmptyEntries)){path=path.Length==0?part:path+"/"+part;TreeNode next;if(!nodes.TryGetValue(path,out next)){next=parent.Nodes.Add(part);next.Tag=path;nodes[path]=next;}parent=next;}if(parent!=all)parent.Text=Path.GetFileName(group.Category)+" ("+group.Slides+")";}
                all.Expand();categories.SelectedNode=all;
            }
            finally{rebinding=false;}
        }
        async void LoadWorkspacePage(bool reset)
        {
            if(!ready||catalog==null||rebinding||IsDisposed||injected)return;if(reset)page=0;int revision=++viewRevision;
            if(viewCancellation!=null){viewCancellation.Cancel();viewCancellation.Dispose();}viewCancellation=CancellationTokenSource.CreateLinkedTokenSource(workspaceToken);var token=viewCancellation.Token;
            var presentation=presentations.SelectedItem as LibraryPresentationSummary;
            var request=new LibrarySearchRequest{LibraryId=SelectedLibraryId,PresentationId=presentation==null?null:presentation.Id,Query=query.Text,Author=author.Text,Layout=layout.Text,Tag=tagFilter.Text,Category=String.IsNullOrEmpty(selectedCategory)?null:selectedCategory,CategoryRecursive=true,HideDuplicates=duplicates.Checked,ModifiedSinceUtc=since.Checked?(DateTime?)since.Value.Date.ToUniversalTime():null,Aspect=aspect.SelectedIndex==1?LibraryAspect.Wide:aspect.SelectedIndex==2?LibraryAspect.Standard:LibraryAspect.All,Mode=(LibrarySearchMode)Math.Max(0,searchMode.SelectedIndex),Sort=sort.SelectedIndex==1?LibrarySortMode.Title:sort.SelectedIndex==2?LibrarySortMode.Newest:sort.SelectedIndex==3?LibrarySortMode.Relevance:LibrarySortMode.Source,Offset=checked(page*LibraryGallery.PageSize),Limit=LibraryGallery.PageSize};
            try
            {
                var result=await Task.Run(()=>catalog.Search(request,token));await WorkspaceUi();if(IsDisposed||Disposing||revision!=viewRevision||token.IsCancellationRequested)return;
                totalMatched=result.Total;filtered=result.Slides.Select(r=>new LocalSlideAsset{File=r.SourcePath,Slide=r.Position,Title=r.Title,Text=r.Text,Author=r.Author,Layout=r.Layout,Master=r.Master,Width=(float)r.Width,Height=(float)r.Height,Modified=r.ModifiedUtc,Signature=r.SlideFingerprint,IndexedReference=r}).ToList();ShowPage();
            }
            catch(OperationCanceledException){}
            catch(Exception ex)when(ex is IOException||ex is ArgumentException){await WorkspaceUi();if(!IsDisposed&&revision==viewRevision)status.Text="Indexed search unavailable: "+ex.Message;}
        }
        async void RegisterWorkspaceFolder(string root)
        {
            if(catalog==null||indexing)return;await WorkspaceOperation(()=>{string id=catalog.Register(root,Path.GetFileName(root.TrimEnd('\\')),true);catalog.Refresh(id,workCancellation.Token);return id;});
        }
        async void RemoveWorkspaceLibrary()
        {
            var selected=libraries.SelectedItem as LibraryChoice;if(selected==null||selected.Registration==null){status.Text="Select one registered library to remove.";return;}if(Managed(selected.Registration.Root)){status.Text="This library is configured by the deployment. Update its configured source definition to remove it.";return;}
            string id=selected.Registration.Id;await WorkspaceOperation(()=>{catalog.Remove(id);return (string)null;});
        }
        async void RefreshWorkspace(bool force)
        {
            if(catalog==null||indexing)return;string selected=SelectedLibraryId;bool recurse=recursive.Checked;
            await WorkspaceOperation(()=>
            {
                var sources=catalog.Libraries().Where(l=>selected==null||l.Id==selected).ToArray();foreach(var source in sources){if(selected!=null&&!Managed(source.Root))catalog.Register(source.Root,source.Name,recurse);catalog.Refresh(source.Id,workCancellation.Token,null,force);}return selected;
            });
        }
        async Task WorkspaceOperation(Func<string> operation)
        {
            if(indexing)return;indexing=true;previewTimer.Stop();pending.Clear();toolbar.Enabled=false;cancelIndex.Enabled=true;workCancellation=CancellationTokenSource.CreateLinkedTokenSource(workspaceToken);status.Text="Updating PIIDX catalog; previous committed data is retained until completion…";
            try{string selected=await Task.Run(operation);await WorkspaceUi();if(IsDisposed||Disposing)return;await WorkspaceNavigation(selected);await WorkspaceUi();if(IsDisposed)return;if(monitor!=null)monitor.Reset();Filter();}
            catch(OperationCanceledException){await WorkspaceUi();if(!IsDisposed)status.Text="Indexing cancelled. Previous catalog retained.";}
            catch(Exception ex)when(ex is IOException||ex is UnauthorizedAccessException||ex is System.Xml.XmlException||ex is ArgumentException){await WorkspaceUi();if(!IsDisposed)status.Text="Library update failed: "+ex.Message;}
            finally{await WorkspaceUi();indexing=false;if(workCancellation!=null){workCancellation.Dispose();workCancellation=null;}if(!IsDisposed){toolbar.Enabled=true;cancelIndex.Enabled=false;}}
        }
        async Task<Bitmap> WorkspacePreview(LocalSlideAsset asset)
        {
            if(asset.IndexedReference==null||catalog==null)return previewProvider==null?null:previewProvider(asset);
            var reference=asset.IndexedReference;int width=720,height=(int)Math.Round(720*asset.Height/asset.Width);byte[] bytes=null;var token=workspaceToken;
            await Task.Run(()=>{catalog.VerifySource(reference,token);try{bytes=catalog.Preview(reference.SlideId,width,height,token);}catch(InvalidDataException){bytes=null;}});
            await WorkspaceUi();if(IsDisposed||Disposing)return null;if(bytes!=null){using(var stream=new MemoryStream(bytes))using(var image=Image.FromStream(stream))return new Bitmap(image);}
            if(previewProvider==null)return null;var bitmap=previewProvider(asset);
            try
            {
                using(var stream=new MemoryStream()){bitmap.Save(stream,ImageFormat.Png);bytes=stream.ToArray();}
                await Task.Run(()=>catalog.PutPreview(reference,bytes,width,height,token));await WorkspaceUi();if(IsDisposed||Disposing){bitmap.Dispose();return null;}return bitmap;
            }
            catch(OperationCanceledException){bitmap.Dispose();throw;}
            catch(Exception ex)when(ex is IOException||ex is InvalidOperationException){await WorkspaceUi();if(!IsDisposed)status.Text="Native preview displayed; cache could not be saved ("+ex.GetType().Name+").";return bitmap;}
        }
        void CloseWorkspace()
        {
            if(monitor!=null){monitor.Dispose();monitor=null;}workspaceTimer.Dispose();if(workCancellation!=null)workCancellation.Cancel();if(viewCancellation!=null){viewCancellation.Cancel();viewCancellation.Dispose();viewCancellation=null;}cancellation.Cancel();
        }
    }
}