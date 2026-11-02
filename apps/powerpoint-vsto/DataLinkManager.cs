using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office=Microsoft.Office.Core;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
namespace PowerPointIntelligence.AddIn
{
    internal enum LocalLinkFileState { Available,Missing,Unavailable,RemoteNotChecked }
    internal sealed class LocalLinkFileStatus
    {
        internal LocalLinkFileState State;
        internal DateTime? LastWriteUtc;
        internal string Message;
    }
    internal sealed class LinkObjectReference
    {
        internal int SlideId,SlideIndex;
        internal int[] ShapePath;
        internal string ShapeName;
        internal IReadOnlyList<DataSourceLink> Links;
        internal int ReferenceCount,InvalidReferences;
        internal string Identity {get{return SlideId.ToString(CultureInfo.InvariantCulture)+":"+String.Join("/",ShapePath.Select(i=>i.ToString(CultureInfo.InvariantCulture)));}}
    }
    internal sealed class LocalLinkSource
    {
        internal DataSourceLink Link;
        internal LocalLinkFileStatus File;
        internal bool Registered,Conflict;
        internal string Status;
        internal IReadOnlyList<LinkObjectReference> Objects;
        internal bool CanRefresh(LinkObjectReference item,bool catalogComplete)
        {return catalogComplete&&!Conflict&&File.State==LocalLinkFileState.Available&&item!=null&&item.ReferenceCount==1&&item.InvalidReferences==0&&item.Links.Count==1&&DataLinkInventory.Same(item.Links[0],Link);}
    }
    internal sealed class LocalLinkInventory
    {
        internal IReadOnlyList<LocalLinkSource> Sources;
        internal IReadOnlyList<string> Diagnostics;
        internal bool CatalogComplete;
        internal IReadOnlyList<LinkObjectReference> UnresolvedObjects;
    }
    internal static class DataLinkInventory
    {
        internal const int MaximumSources=256,MaximumObjects=5000,MaximumReferencesPerObject=16,MaximumDepth=32;
        internal static bool Same(DataSourceLink a,DataSourceLink b)
        {return a!=null&&b!=null&&String.Equals(a.Id,b.Id,StringComparison.OrdinalIgnoreCase)&&String.Equals(a.Format,b.Format,StringComparison.Ordinal)&&String.Equals(a.Path,b.Path,StringComparison.OrdinalIgnoreCase)&&a.LinkedAtUtc==b.LinkedAtUtc&&String.Equals(a.Mapping==null?null:a.Mapping.Key,b.Mapping==null?null:b.Mapping.Key,StringComparison.Ordinal);}
        internal static IReadOnlyList<DataSourceLink> DecodeReferences(IEnumerable<string> xml,out int count,out int invalid)
        {
            if(xml==null)throw new ArgumentNullException(nameof(xml));var links=new List<DataSourceLink>();count=invalid=0;
            foreach(string item in xml){if(++count>MaximumReferencesPerObject)throw new InvalidDataException("An object exceeds 16 source references; inventory stopped without modification.");try{links.Add(DataSources.Decode(item));}catch(Exception ex) when(ex is InvalidDataException||ex is IOException||ex is ArgumentException){invalid++;}}
            return links.AsReadOnly();
        }
        private static DataSourceLink Copy(DataSourceLink link){return DataSources.Decode(DataSources.Encode(link));}
        private static string Key(DataSourceLink link){return link.Id.ToUpperInvariant()+"\n"+link.Format+"\n"+link.Path.ToUpperInvariant()+"\n"+link.LinkedAtUtc.ToString("o",CultureInfo.InvariantCulture)+"\n"+(link.Mapping==null?"":link.Mapping.Key);}
        internal static LocalLinkInventory Build(IEnumerable<DataSourceLink> registered,IEnumerable<LinkObjectReference> objects,Func<string,LocalLinkFileStatus> probe,bool catalogComplete=true,IEnumerable<string> diagnostics=null)
        {
            if(registered==null||objects==null||probe==null)throw new ArgumentNullException("Source inventory inputs are required.");
            var catalog=registered.Take(MaximumSources+1).Select(Copy).ToArray();var bindings=objects.Take(MaximumObjects+1).ToArray();var notes=(diagnostics??Enumerable.Empty<string>()).ToList();
            if(catalog.Length>MaximumSources||bindings.Length>MaximumObjects)throw new InvalidDataException("Source inventory exceeds its bounded source/object limit.");
            if(bindings.Any(o=>o==null||o.SlideId<1||o.SlideIndex<1||o.ShapePath==null||o.ShapePath.Length<1||o.ShapePath.Length>MaximumDepth||o.ShapePath.Any(i=>i<1)||o.Links==null||o.ReferenceCount<o.Links.Count||o.ReferenceCount>MaximumReferencesPerObject||o.InvalidReferences<0||o.InvalidReferences+o.Links.Count!=o.ReferenceCount)||bindings.Select(o=>o.Identity).Distinct().Count()!=bindings.Length)throw new InvalidDataException("Invalid source-object associations.");
            var sources=catalog.Concat(bindings.SelectMany(o=>o.Links).Select(Copy)).GroupBy(Key).Select(g=>g.First()).ToArray();if(sources.Length>MaximumSources)throw new InvalidDataException("Distinct source inventory exceeds 256 entries.");
            var duplicates=new HashSet<string>(catalog.GroupBy(l=>l.Id,StringComparer.OrdinalIgnoreCase).Where(g=>g.Count()>1).Select(g=>g.Key),StringComparer.OrdinalIgnoreCase);
            var conflicting=new HashSet<string>(sources.GroupBy(l=>l.Id,StringComparer.OrdinalIgnoreCase).Where(g=>g.Select(Key).Distinct().Count()>1).Select(g=>g.Key),StringComparer.OrdinalIgnoreCase);
            var rows=new List<LocalLinkSource>();foreach(var link in sources){var file=probe(link.Path);if(file==null)throw new InvalidDataException("File-status provider returned no result.");var associated=bindings.Where(o=>o.Links.Any(l=>Same(l,link))).ToArray();bool conflict=duplicates.Contains(link.Id)||conflicting.Contains(link.Id);string state=file.State==LocalLinkFileState.Available?"Available when checked; refresh state unknown":file.State==LocalLinkFileState.Missing?"Source file missing":file.State==LocalLinkFileState.RemoteNotChecked?"Network source not probed":"Source metadata unavailable";
                if(conflict)state="Conflicting or duplicate source identity; refresh disabled";else if(!catalogComplete)state="Registry damaged/incomplete; refresh disabled; "+state;else if(associated.Any(o=>o.InvalidReferences>0||o.ReferenceCount!=1))state="Ambiguous/damaged object reference; "+state;
                bool isRegistered=catalog.Any(l=>Same(l,link));if(!isRegistered)state="Object-only reference; "+state;
                rows.Add(new LocalLinkSource {Link=link,Registered=isRegistered,Conflict=conflict,File=file,Status=state,Objects=Array.AsReadOnly(associated)});
            }
            int damaged=bindings.Count(o=>o.InvalidReferences>0),multiple=bindings.Count(o=>o.ReferenceCount>1);if(damaged>0)notes.Add(damaged+" objects contain damaged source references; their metadata was preserved.");if(multiple>0)notes.Add(multiple+" objects contain multiple references; refresh is disabled for them.");
            return new LocalLinkInventory {Sources=rows.AsReadOnly(),Diagnostics=notes.AsReadOnly(),CatalogComplete=catalogComplete,UnresolvedObjects=Array.AsReadOnly(bindings.Where(o=>o.Links.Count==0&&o.InvalidReferences>0).ToArray())};
        }
        internal static LocalLinkFileStatus ProbeLocal(string path)
        {
            // No workbook/CSV/JSON/XML contents are read, and no remote endpoint is contacted.
            if(path.StartsWith(@"\\",StringComparison.Ordinal))return new LocalLinkFileStatus {State=LocalLinkFileState.RemoteNotChecked,Message="UNC path not probed"};
            try{string root=Path.GetPathRoot(path);if(new DriveInfo(root).DriveType==DriveType.Network)return new LocalLinkFileStatus {State=LocalLinkFileState.RemoteNotChecked,Message="Mapped network drive not probed"};var attributes=File.GetAttributes(path);if((attributes&FileAttributes.Directory)!=0)return new LocalLinkFileStatus {State=LocalLinkFileState.Unavailable,Message="Source path is a directory"};return new LocalLinkFileStatus {State=LocalLinkFileState.Available,LastWriteUtc=File.GetLastWriteTimeUtc(path),Message="File metadata only; no refresh baseline"};}
            catch(FileNotFoundException){return new LocalLinkFileStatus {State=LocalLinkFileState.Missing,Message="File not found when checked"};}
            catch(DirectoryNotFoundException){return new LocalLinkFileStatus {State=LocalLinkFileState.Missing,Message="Source directory not found when checked"};}
            catch(Exception ex) when(ex is IOException||ex is UnauthorizedAccessException||ex is ArgumentException||ex is NotSupportedException||ex is System.Security.SecurityException){return new LocalLinkFileStatus {State=LocalLinkFileState.Unavailable,Message="Cannot read file metadata ("+ex.GetType().Name+")"};}
        }
        internal static LocalLinkInventory Capture(PowerPoint.Presentation presentation)
        {
            if(presentation==null)throw new ArgumentNullException(nameof(presentation));var diagnostics=new List<string>();var catalog=new List<DataSourceLink>();bool complete=true;if(presentation.CustomXMLParts.SelectByNamespace(DataSources.Namespace).Count>MaximumSources)throw new InvalidDataException("Source catalog exceeds 256 parts.");
            try{catalog.AddRange(DataSources.Read(presentation));}
            catch(Exception ex) when(ex is InvalidDataException||ex is IOException||ex is COMException){complete=false;diagnostics.Add("Presentation source registry could not be read as a unique valid catalog; individual records are shown where valid.");int parts=0;foreach(Office.CustomXMLPart part in presentation.CustomXMLParts.SelectByNamespace(DataSources.Namespace)){if(++parts>MaximumSources)throw new InvalidDataException("Source catalog exceeds 256 parts.");try{catalog.Add(DataSources.Decode(part.XML));}catch(Exception error) when(error is InvalidDataException||error is IOException||error is COMException){diagnostics.Add("A damaged presentation source record was preserved.");}}}
            if(catalog.Count>MaximumSources)throw new InvalidDataException("Source catalog exceeds 256 entries.");
            var objects=new List<LinkObjectReference>();int visited=0,partsVisited=0;long metadataCharacters=0;Action<PowerPoint.Shape,PowerPoint.Slide,int[]> visit=null;
            visit=(shape,slide,path)=>{if(++visited>MaximumObjects||path.Length>MaximumDepth)throw new InvalidDataException("Source-object traversal exceeds 5000 objects or 32 group levels.");var parts=new List<string>();int inaccessible=0;foreach(Office.CustomXMLPart part in shape.CustomerData){if(++partsVisited>10000)throw new InvalidDataException("Object metadata traversal exceeds 10000 parts.");try{if(part.NamespaceURI==DataSources.Namespace){if(parts.Count+inaccessible>=MaximumReferencesPerObject)throw new InvalidDataException("An object exceeds 16 source references.");string xml=part.XML;metadataCharacters+=xml.Length;if(metadataCharacters>16L*1024*1024)throw new InvalidDataException("Object source metadata exceeds the 16 MB inventory budget.");parts.Add(xml);}}catch(COMException){if(++inaccessible+parts.Count>MaximumReferencesPerObject)throw new InvalidDataException("An object exceeds 16 unreadable source references.");}}
                if(parts.Count+inaccessible>0){int count,invalid;var links=DecodeReferences(parts,out count,out invalid);objects.Add(new LinkObjectReference {SlideId=slide.SlideID,SlideIndex=slide.SlideIndex,ShapePath=path,ShapeName=shape.Name,Links=links,ReferenceCount=count+inaccessible,InvalidReferences=invalid+inaccessible});}
                if(shape.Type==Office.MsoShapeType.msoGroup)foreach(PowerPoint.Shape child in shape.GroupItems)visit(child,slide,path.Concat(new[]{child.Id}).ToArray());};
            foreach(PowerPoint.Slide slide in presentation.Slides)foreach(PowerPoint.Shape shape in slide.Shapes)visit(shape,slide,new[]{shape.Id});return Build(catalog,objects,ProbeLocal,complete,diagnostics);
        }
        internal static PowerPoint.Shape Find(PowerPoint.Presentation deck,LinkObjectReference reference)
        {
            if(deck==null)throw new ArgumentNullException(nameof(deck));if(reference==null||reference.SlideId<1||reference.ShapePath==null||reference.ShapePath.Length<1||reference.ShapePath.Length>MaximumDepth||reference.ShapePath.Any(i=>i<1))throw new InvalidDataException("Invalid object location.");var slide=deck.Slides.FindBySlideID(reference.SlideId);PowerPoint.Shape current=null;for(int depth=0;depth<reference.ShapePath.Length;depth++){var candidates=depth==0?slide.Shapes.Cast<PowerPoint.Shape>():current.Type==Office.MsoShapeType.msoGroup?current.GroupItems.Cast<PowerPoint.Shape>():Enumerable.Empty<PowerPoint.Shape>();current=candidates.SingleOrDefault(s=>s.Id==reference.ShapePath[depth]);if(current==null)throw new InvalidOperationException("Linked object moved or was deleted. Rescan the inventory.");}return current;
        }
    }
    internal sealed class DataLinkManager:Form
    {
        private readonly PowerPoint.Application app;
        private readonly PowerPoint.Presentation deck;
        private readonly Action<PowerPoint.Shape,DataSourceLink> refresh;
        private readonly DataGridView sources=Grid(),objects=Grid();
        private readonly Label status=new Label {Dock=DockStyle.Bottom,Height=70,Padding=new Padding(8)};
        private readonly Button update=new Button {Text="Refresh selected object",AutoSize=true};
        private LocalLinkInventory inventory;
        private bool loading;
        private static DataGridView Grid(){return new DataGridView {Dock=DockStyle.Fill,ReadOnly=true,AllowUserToAddRows=false,AllowUserToDeleteRows=false,MultiSelect=false,SelectionMode=DataGridViewSelectionMode.FullRowSelect,RowHeadersVisible=false,AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.Fill};}
        internal DataLinkManager(PowerPoint.Application app,Action<PowerPoint.Shape,DataSourceLink> refresh=null)
        {
            if(app==null||app.Presentations.Count==0)throw new InvalidOperationException("Open a presentation before inspecting data links.");this.app=app;deck=app.ActivePresentation;this.refresh=refresh;Text="Data sources and linked objects";Width=1150;Height=680;StartPosition=FormStartPosition.CenterScreen;
            foreach(string name in new[]{"Format","Source file","Status","Modified UTC","Objects"})sources.Columns.Add(name,name);sources.Columns[1].FillWeight=220;sources.Columns[2].FillWeight=220;foreach(string name in new[]{"Slide","Object","Association status"})objects.Columns.Add(name,name);objects.Columns[1].FillWeight=180;
            var split=new SplitContainer {Dock=DockStyle.Fill,Orientation=Orientation.Horizontal,SplitterDistance=290};split.Panel1.Controls.Add(sources);split.Panel2.Controls.Add(objects);var buttons=new FlowLayoutPanel {Dock=DockStyle.Top,Height=42,Padding=new Padding(5)};var rescan=new Button {Text="Rescan metadata",AutoSize=true};var locate=new Button {Text="Locate selected object",AutoSize=true};buttons.Controls.AddRange(new Control[]{rescan,locate,update});Controls.Add(split);Controls.Add(status);Controls.Add(buttons);
            sources.SelectionChanged+=(s,e)=>{if(!loading)ShowObjects();};objects.SelectionChanged+=(s,e)=>UpdateEnabled();rescan.Click+=(s,e)=>Rescan();locate.Click+=(s,e)=>Locate();update.Click+=(s,e)=>RefreshObject();Shown+=(s,e)=>Rescan();
        }
        private LocalLinkSource SelectedSource {get{return sources.SelectedRows.Count==0?null:sources.SelectedRows[0].Tag as LocalLinkSource;}}
        private LinkObjectReference SelectedObject {get{return objects.SelectedRows.Count==0?null:objects.SelectedRows[0].Tag as LinkObjectReference;}}
        private void Rescan()
        {
            try{inventory=DataLinkInventory.Capture(deck);loading=true;sources.Rows.Clear();foreach(var source in inventory.Sources){int row=sources.Rows.Add(source.Link.Format,source.Link.Path,source.Status,source.File.LastWriteUtc.HasValue?source.File.LastWriteUtc.Value.ToString("u"):"Unknown",source.Objects.Count);sources.Rows[row].Tag=source;}if(inventory.UnresolvedObjects.Count>0){int row=sources.Rows.Add("Unknown","(damaged object metadata)","Preserved; source cannot be inferred","Unknown",inventory.UnresolvedObjects.Count);sources.Rows[row].Tag=inventory.UnresolvedObjects;}loading=false;if(sources.Rows.Count>0)sources.Rows[0].Selected=true;ShowObjects();status.Text=inventory.Sources.Count+" sources. Metadata snapshot only; successful refresh time is unknown. "+String.Join(" ",inventory.Diagnostics);}
            catch(Exception ex) when(ex is InvalidDataException||ex is IOException||ex is COMException||ex is InvalidOperationException){inventory=null;loading=true;sources.Rows.Clear();objects.Rows.Clear();status.Text="Inventory unavailable: "+ex.Message;}
            finally{loading=false;UpdateEnabled();}
        }
        private void ShowObjects()
        {
            objects.Rows.Clear();var source=SelectedSource;var references=source!=null?source.Objects:sources.SelectedRows.Count>0?sources.SelectedRows[0].Tag as IReadOnlyList<LinkObjectReference>:null;if(references!=null)foreach(var item in references){int row=objects.Rows.Add(item.SlideIndex,item.ShapeName,item.InvalidReferences>0?"Damaged reference":item.ReferenceCount!=1?"Multiple references":"Single reference");objects.Rows[row].Tag=item;}if(objects.Rows.Count>0)objects.Rows[0].Selected=true;UpdateEnabled();
        }
        private void UpdateEnabled(){var source=SelectedSource;update.Enabled=refresh!=null&&inventory!=null&&source!=null&&source.CanRefresh(SelectedObject,inventory.CatalogComplete);}
        private void Locate()
        {
            try{var item=SelectedObject;if(item==null)return;if(app.SlideShowWindows.Count>0)throw new InvalidOperationException("Exit the slide show before locating objects.");var shape=DataLinkInventory.Find(deck,item);if(deck.Windows.Count==0)throw new InvalidOperationException("Open a visible window for this presentation.");deck.Windows[1].Activate();deck.Windows[1].ViewType=PowerPoint.PpViewType.ppViewNormal;deck.Slides.FindBySlideID(item.SlideId).Select();shape.Select(Office.MsoTriState.msoTrue);status.Text="Selected the linked object on slide "+deck.Slides.FindBySlideID(item.SlideId).SlideIndex+". No data source was changed.";}
            catch(Exception ex) when(ex is InvalidOperationException||ex is InvalidDataException||ex is COMException){status.Text="Locate failed: "+ex.Message;}
        }
        private void RefreshObject()
        {
            var source=SelectedSource;var item=SelectedObject;if(refresh==null||source==null||inventory==null||!source.CanRefresh(item,inventory.CatalogComplete))return;
            try{var checkedInventory=DataLinkInventory.Capture(deck);var checkedSource=checkedInventory.Sources.SingleOrDefault(s=>DataLinkInventory.Same(s.Link,source.Link));var checkedObject=checkedSource==null?null:checkedSource.Objects.SingleOrDefault(o=>o.Identity==item.Identity);if(checkedSource==null||!checkedSource.CanRefresh(checkedObject,checkedInventory.CatalogComplete))throw new InvalidOperationException("Source status or association changed. Rescan before refreshing.");var shape=DataLinkInventory.Find(deck,item);var xml=new List<string>();foreach(Office.CustomXMLPart part in shape.CustomerData)if(part.NamespaceURI==DataSources.Namespace)xml.Add(part.XML);int count,invalid;var live=DataLinkInventory.DecodeReferences(xml,out count,out invalid);if(count!=1||invalid!=0||!DataLinkInventory.Same(live[0],source.Link))throw new InvalidOperationException("Source metadata changed. Rescan before refreshing.");refresh(shape,DataSources.Decode(DataSources.Encode(source.Link)));Rescan();status.Text="Selected-object refresh completed. Sources were not rebound, moved or removed.";}
            catch(Exception ex) when(ex is InvalidOperationException||ex is InvalidDataException||ex is IOException||ex is COMException){status.Text="Refresh failed: "+ex.Message;}
        }
    }
}
