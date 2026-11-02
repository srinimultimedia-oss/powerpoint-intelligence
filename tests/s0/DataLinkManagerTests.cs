using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PowerPointIntelligence.AddIn;
internal static class DataLinkManagerTests
{
    private static int passed;
    private static void Check(bool condition,string name){if(!condition)throw new Exception("Link inventory: "+name);passed++;}
    private static void Reject(Action action,string name){bool rejected=false;try{action();}catch(Exception ex) when(ex is InvalidDataException||ex is ArgumentException){rejected=true;}Check(rejected,name);}
    private static DataSourceLink Link(string file,string id=null){return new DataSourceLink {Id=id??Guid.NewGuid().ToString("D"),Format="csv",Path=Path.GetFullPath(file),LinkedAtUtc=new DateTime(2020,1,1,0,0,0,DateTimeKind.Utc)};}
    private static LinkObjectReference Object(int id,IReadOnlyList<DataSourceLink> links,int invalid=0,int? count=null){return new LinkObjectReference {SlideId=100,SlideIndex=2,ShapePath=new[]{id},ShapeName="Linked object",Links=links,ReferenceCount=count??(links.Count+invalid),InvalidReferences=invalid};}
    private static LocalLinkFileStatus Available(string path){return new LocalLinkFileStatus {State=LocalLinkFileState.Available,LastWriteUtc=DateTime.UtcNow};}
    internal static void Run()
    {
        string root=Path.Combine(Path.GetTempPath(),"ppi-link-inventory-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(root);string file=Path.Combine(root,"source.csv");File.WriteAllText(file,"not valid CSV; inventory must not parse contents");var link=Link(file);var binding=Object(1,new[]{link});
        var snapshot=DataLinkInventory.Build(new[]{link},new[]{binding},Available);var source=snapshot.Sources.Single();Check(source.Registered&&source.Objects.Count==1&&source.Objects[0].Identity=="100:1","registered source has stable slide/object association");Check(source.CanRefresh(binding,snapshot.CatalogComplete),"single valid available association eligible for supplied refresh delegate");Check(source.Status.Contains("refresh state unknown"),"availability does not claim refresh freshness");Check(!source.CanRefresh(null,true),"missing object not refreshable");
        var local=DataLinkInventory.ProbeLocal(file);Check(local.State==LocalLinkFileState.Available&&local.LastWriteUtc.HasValue&&local.LastWriteUtc.Value.Kind==DateTimeKind.Utc,"only file metadata inspected");Check(DataLinkInventory.ProbeLocal(Path.Combine(root,"absent.csv")).State==LocalLinkFileState.Missing,"missing file status");Directory.CreateDirectory(Path.Combine(root,"folder.csv"));Check(DataLinkInventory.ProbeLocal(Path.Combine(root,"folder.csv")).State==LocalLinkFileState.Unavailable,"directory not accepted as file");Check(DataLinkInventory.ProbeLocal(@"\\server\share\file.csv").State==LocalLinkFileState.RemoteNotChecked,"network source not contacted automatically");
        var missing=DataLinkInventory.Build(new[]{link},new[]{binding},p=>new LocalLinkFileStatus {State=LocalLinkFileState.Missing});Check(!missing.Sources[0].CanRefresh(binding,true),"missing files not refreshable");
        var remote=DataLinkInventory.Build(new[]{link},new[]{binding},p=>new LocalLinkFileStatus {State=LocalLinkFileState.RemoteNotChecked});Check(!remote.Sources[0].CanRefresh(binding,true),"unchecked network sources not refreshable");
        var objectOnly=DataLinkInventory.Build(new DataSourceLink[0],new[]{binding},Available);Check(!objectOnly.Sources[0].Registered&&objectOnly.Sources[0].Status.StartsWith("Object-only reference"),"copied object source is reported without registering it");
        var duplicate=DataLinkInventory.Build(new[]{link,link},new[]{binding},Available);Check(duplicate.Sources.Count==1&&duplicate.Sources[0].Conflict&&!duplicate.Sources[0].CanRefresh(binding,true),"duplicate registration preserved and disables refresh");
        var other=Link(Path.Combine(root,"other.csv"),link.Id);var conflict=DataLinkInventory.Build(new[]{link,other},new[]{binding},Available);Check(conflict.Sources.Count==2&&conflict.Sources.All(s=>s.Conflict),"same identity different paths remain visible and ambiguous");Check(!conflict.Sources[0].CanRefresh(binding,true),"conflicting source cannot refresh");
        var multiple=Object(2,new[]{link,link});var ambiguous=DataLinkInventory.Build(new[]{link},new[]{multiple},Available);Check(ambiguous.Sources[0].Objects.Count==1&&!ambiguous.Sources[0].CanRefresh(multiple,true)&&ambiguous.Diagnostics.Count>0,"multiple object references are not silently deduplicated");
        int count,invalid;var decoded=DataLinkInventory.DecodeReferences(new[]{DataSources.Encode(link),"<broken"},out count,out invalid);Check(count==2&&invalid==1&&decoded.Count==1,"damaged and valid references reported separately");var damaged=Object(3,decoded,invalid,count);var damage=DataLinkInventory.Build(new[]{link},new[]{damaged},Available);Check(!damage.Sources[0].CanRefresh(damaged,true)&&damage.Diagnostics.Count>0,"mixed damaged object reference disables refresh");
        var corrupt=Object(4,new DataSourceLink[0],1);var unresolved=DataLinkInventory.Build(new DataSourceLink[0],new[]{corrupt},Available);Check(unresolved.Sources.Count==0&&unresolved.UnresolvedObjects.Count==1&&unresolved.Diagnostics.Count>0,"fully damaged objects remain visible as unresolved associations");
        var incomplete=DataLinkInventory.Build(new[]{link},new[]{binding},Available,false,new[]{"damaged registry"});Check(!incomplete.Sources[0].CanRefresh(binding,incomplete.CatalogComplete)&&incomplete.Sources[0].Status.Contains("Registry damaged"),"incomplete registry disables mutations");
        var nested=Object(5,new[]{link});nested.ShapePath=new[]{40,5};Check(DataLinkInventory.Build(new[]{link},new[]{nested},Available).Sources[0].Objects[0].Identity=="100:40/5","nested object path retained");
        var cloned=DataSources.Decode(DataSources.Encode(link));Check(DataLinkInventory.Same(link,cloned),"source identity round trip");cloned.Path=Path.Combine(root,"changed.csv");Check(!DataLinkInventory.Same(link,cloned),"changed path invalidates association");
        Reject(()=>DataLinkInventory.DecodeReferences(Enumerable.Repeat(DataSources.Encode(link),17),out count,out invalid),"per-object reference cap");
        Reject(()=>DataLinkInventory.Build(Enumerable.Range(1,257).Select(i=>Link(Path.Combine(root,i+".csv"))),new LinkObjectReference[0],Available),"catalog source cap");
        Reject(()=>DataLinkInventory.Build(new[]{link},new[]{binding,binding},Available),"duplicate object identity rejected");
        var deep=Object(6,new[]{link});deep.ShapePath=Enumerable.Repeat(1,33).ToArray();Reject(()=>DataLinkInventory.Build(new[]{link},new[]{deep},Available),"nested traversal depth cap");
        Reject(()=>DataLinkInventory.Build(new[]{link},Enumerable.Range(1,5001).Select(i=>Object(i,new[]{link})),Available),"object count cap");
        Check(File.ReadAllText(file)=="not valid CSV; inventory must not parse contents","source data was not modified");
        Console.WriteLine("PASS "+passed+" local data-link inventory/association policy assertions. Native inventory, locate and supplied refresh delegate remain unverified.");
    }
}
