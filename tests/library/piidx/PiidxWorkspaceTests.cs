using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using PowerPointIntelligence.Library;
public static class PiidxWorkspaceTests
{
    static int count;
    static void Check(bool value,string name){if(!value)throw new Exception("FAIL "+name);count++;}
    static void Throws<T>(Action action,string name)where T:Exception{try{action();}catch(T){count++;return;}throw new Exception("FAIL "+name);}
    public static string Run(string output)
    {
        count=0;Directory.CreateDirectory(output);string root=Path.Combine(output,"sources"),nested=Path.Combine(root,"Finance","Q1");Directory.CreateDirectory(nested);string one=Path.Combine(root,"one.pptx");PiidxSourceTests.Fixture(one);File.Copy(one,Path.Combine(root,"two.pptx"));PiidxSourceTests.Fixture(Path.Combine(nested,"report.potx"));
        var catalog=PiidxCatalog.OpenOrCreate(Path.Combine(output,"workspace.piidx"));string id=catalog.Register(root,"Workspace",true);catalog.Refresh(id,CancellationToken.None);
        Check(catalog.Search(new LibrarySearchRequest{HideDuplicates=true},CancellationToken.None).Total<6,"duplicate filtering executes in database");
        Check(catalog.Search(new LibrarySearchRequest{Query="one",HideDuplicates=true},CancellationToken.None).Total==2,"duplicate representative stays within filtered candidates");
        Check(catalog.Search(new LibrarySearchRequest{Category="Finance",CategoryRecursive=true},CancellationToken.None).Total==2,"category descendants included");
        Check(catalog.Search(new LibrarySearchRequest{Category="Finance"},CancellationToken.None).Total==0,"exact category excludes descendants");
        Check(catalog.Search(new LibrarySearchRequest{Aspect=LibraryAspect.Wide},CancellationToken.None).Total==6,"wide aspect filter");
        Check(catalog.Search(new LibrarySearchRequest{Aspect=LibraryAspect.Standard},CancellationToken.None).Total==0,"standard aspect filter");
        Check(catalog.Search(new LibrarySearchRequest{ModifiedSinceUtc=DateTime.UtcNow.AddDays(1)},CancellationToken.None).Total==0,"UTC date filter");
        Check(catalog.Search(new LibrarySearchRequest{Author="Test author",Sort=LibrarySortMode.Title},CancellationToken.None).Total==6,"author and title ordering");
        Check(catalog.Search(new LibrarySearchRequest{Query="résumé",HideDuplicates=true,Sort=LibrarySortMode.Relevance},CancellationToken.None).Total>0,"FTS rank and duplicate filter compose");
        Check(catalog.Categories(id,CancellationToken.None).Any(c=>c.Category=="Finance/Q1"&&c.Slides==2),"category navigation comes from index");
        Check(catalog.Presentations(id,CancellationToken.None).Count==3,"presentation navigation comes from index");
        string config=Path.Combine(output,"app.config"),external=Path.Combine(output,"libraries.xml");XNamespace ns="urn:powerpoint-intelligence:libraries:v1";
        new XDocument(new XElement(ns+"Libraries",new XAttribute("version","1"),new XElement(ns+"Library",new XAttribute("path",root),new XAttribute("name","Enterprise"),new XAttribute("recursive","true")))).Save(external);
        new XDocument(new XElement("configuration",new XElement("appSettings",new XElement("add",new XAttribute("key","LibrarySourcesConfig"),new XAttribute("value","libraries.xml")),new XElement("add",new XAttribute("key","LibraryIndexPath"),new XAttribute("value","local.piidx")),new XElement("add",new XAttribute("key","LibraryRefreshSeconds"),new XAttribute("value","5")),new XElement("add",new XAttribute("key","LibraryWatchRootLimit"),new XAttribute("value","0"))))).Save(config);
        var profile=LibraryRuntimeProfile.Load(config,output,output);Check(profile.IndexPath==Path.Combine(output,"local.piidx")&&profile.RefreshSeconds==5&&profile.WatchRootLimit==0,"bounded runtime settings load");
        Check(profile.Sources.Single().Managed&&profile.Sources[0].Root==root,"enterprise registration definitions load");
        Throws<System.Xml.XmlException>(()=>LibraryRuntimeProfile.ParseSources("<!DOCTYPE Libraries [<!ENTITY x SYSTEM 'file:///not-read'>]><Libraries/>",output),"enterprise DTD blocked");
        Throws<InvalidDataException>(()=>LibraryRuntimeProfile.ParseSources("<Libraries xmlns='urn:powerpoint-intelligence:libraries:v1' version='2'/>",output),"unknown enterprise schema rejected");
        Directory.CreateDirectory(Path.GetDirectoryName(profile.LegacyCachePath));new XDocument(new XElement(XName.Get("Library","urn:powerpoint-intelligence:local-library:v1"),new XAttribute("version","1"),new XAttribute("folder",root),new XElement(XName.Get("Slide","urn:powerpoint-intelligence:local-library:v1"),"NEVER_IMPORT_LEGACY_TEXT"))).Save(profile.LegacyCachePath);
        string before=PresentationSource.Fingerprint(profile.LegacyCachePath,CancellationToken.None);string migrated=profile.MigrateLegacy(catalog,CancellationToken.None);
        Check(migrated==id&&PresentationSource.Fingerprint(profile.LegacyCachePath,CancellationToken.None)==before,"legacy source registration migrates without changing original");
        Check(catalog.Search(new LibrarySearchRequest{Query="NEVER_IMPORT_LEGACY_TEXT"},CancellationToken.None).Total==0,"migration reindexes original sources instead of cached text");
        Check(profile.MigrateLegacy(catalog,CancellationToken.None)==null,"migration receipt avoids repeated import");
        using(var observed=new ManualResetEventSlim())using(var monitor=new LibraryRefreshMonitor(catalog,30,4,n=>{if(n.Result!=null&&n.Result.FilesUpdated>0)observed.Set();}))
        {
            PiidxSourceTests.Fixture(Path.Combine(root,"watched.pptx"));if(!observed.Wait(10000))throw new Exception("File watcher did not refresh the changed source.");
            Check(catalog.Search(new LibrarySearchRequest(),CancellationToken.None).Total==8,"file watcher refreshes catalog automatically");
        }
        using(var observed=new ManualResetEventSlim())using(var monitor=new LibraryRefreshMonitor(catalog,5,0,n=>{if(n.Result!=null&&n.Result.FilesUpdated>0)observed.Set();}))
        {
            PiidxSourceTests.Fixture(Path.Combine(root,"polled.pptx"));if(!observed.Wait(15000))throw new Exception("Polling fallback did not refresh the source.");
            Check(catalog.Search(new LibrarySearchRequest(),CancellationToken.None).Total==10,"polling works without watchers");
        }
        return count+" workspace configuration/navigation/migration/automatic-refresh assertions passed; no Office automation executed.";
    }
}