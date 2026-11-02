using PowerPointIntelligence.Security;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
namespace PowerPointIntelligence.Library
{
    internal sealed class LibrarySourceDefinition
    {
        internal string Root {get;set;}
        internal string Name {get;set;}
        internal bool Recursive {get;set;}
        internal bool Managed {get;set;}
    }
    internal sealed class LibraryRuntimeProfile
    {
        internal string IndexPath {get;private set;}
        internal string LegacyCachePath {get;private set;}
        internal int RefreshSeconds {get;private set;}
        internal int WatchRootLimit {get;private set;}
        internal readonly List<LibrarySourceDefinition> Sources=new List<LibrarySourceDefinition>();
        static XDocument Xml(string text)
        {
            using(var reader=XmlReader.Create(new StringReader(text),new XmlReaderSettings{DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=4000000}))return XDocument.Load(reader);
        }
        static string PathValue(string value,string baseDirectory){return PathPolicy.Configured(value,baseDirectory);}
        internal static LibraryRuntimeProfile Load(string configPath,string baseDirectory,string localAppData)
        {
            configPath=PathPolicy.ConfigFile(configPath);baseDirectory=PathPolicy.Configured(baseDirectory,Environment.CurrentDirectory);localAppData=PathPolicy.Configured(localAppData,Environment.CurrentDirectory);
            var result=new LibraryRuntimeProfile{IndexPath=Path.Combine(localAppData,"PowerPointIntelligence","Library","catalog.piidx"),LegacyCachePath=Path.Combine(localAppData,"PowerPointIntelligence","library-index-v1.xml"),RefreshSeconds=60,WatchRootLimit=32};
            if(!File.Exists(configPath))return result;
            if(new FileInfo(configPath).Length>4000000)throw new InvalidDataException("Library configuration exceeds limits.");
            var xml=Xml(File.ReadAllText(configPath));if(xml.Root==null||xml.Root.Name!="configuration")throw new InvalidDataException("Library configuration root is invalid.");
            var settings=xml.Root.Elements("appSettings").Elements("add").Where(e=>((string)e.Attribute("key")??"").StartsWith("Library",StringComparison.Ordinal)||(string)e.Attribute("key")=="TemplateLibraryFolder").ToArray();
            var values=new Dictionary<string,string>(StringComparer.Ordinal);
            foreach(var setting in settings){string key=(string)setting.Attribute("key");if(values.ContainsKey(key))throw new InvalidDataException("Duplicate library configuration setting.");if(!new[]{"TemplateLibraryFolder","LibraryIndexPath","LibrarySourcesConfig","LibraryRefreshSeconds","LibraryWatchRootLimit"}.Contains(key))throw new InvalidDataException("Unknown library configuration setting.");values[key]=(string)setting.Attribute("value")??"";}
            string value;
            if(values.TryGetValue("LibraryIndexPath",out value)&&!String.IsNullOrWhiteSpace(value))result.IndexPath=PathValue(value,baseDirectory);
            if(!String.Equals(Path.GetExtension(result.IndexPath),".piidx",StringComparison.OrdinalIgnoreCase))throw new InvalidDataException("Configured library index must use .piidx.");
            foreach(string key in new[]{"LibraryRefreshSeconds","LibraryWatchRootLimit"})if(values.TryGetValue(key,out value)&&value.Length>0){int parsed;if(!Int32.TryParse(value,NumberStyles.None,CultureInfo.InvariantCulture,out parsed)||(key=="LibraryRefreshSeconds"?(parsed<5||parsed>86400):(parsed<0||parsed>128)))throw new InvalidDataException("Library polling/watcher setting exceeds bounds.");if(key=="LibraryRefreshSeconds")result.RefreshSeconds=parsed;else result.WatchRootLimit=parsed;}
            if(values.TryGetValue("TemplateLibraryFolder",out value)&&!String.IsNullOrWhiteSpace(value))result.Sources.Add(new LibrarySourceDefinition{Root=PathValue(value,baseDirectory),Name="Templates",Recursive=true,Managed=true});
            if(values.TryGetValue("LibrarySourcesConfig",out value)&&!String.IsNullOrWhiteSpace(value))
            {
                string file=PathValue(value,baseDirectory);if(!File.Exists(file))throw new FileNotFoundException("Configured enterprise library definitions are unavailable.",file);
                if(new FileInfo(file).Length>4000000)throw new InvalidDataException("Enterprise library definitions exceed limits.");
                result.Sources.AddRange(ParseSources(File.ReadAllText(file),Path.GetDirectoryName(file)));
            }
            return result;
        }
        internal static IReadOnlyList<LibrarySourceDefinition> ParseSources(string text,string baseDirectory)
        {
            XNamespace ns="urn:powerpoint-intelligence:libraries:v1";var xml=Xml(text);var root=xml.Root;
            if(root==null||root.Name!=ns+"Libraries"||(string)root.Attribute("version")!="1")throw new InvalidDataException("Unknown enterprise library schema.");
            var result=new List<LibrarySourceDefinition>();var seen=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach(var entry in root.Elements())
            {
                if(entry.Name!=ns+"Library"||entry.Attributes().Any(a=>!a.IsNamespaceDeclaration&&!new[]{"path","name","recursive"}.Contains(a.Name.ToString())))throw new InvalidDataException("Unknown enterprise library field.");
                string path=(string)entry.Attribute("path"),name=(string)entry.Attribute("name"),recursive=(string)entry.Attribute("recursive")??"true";
                if(String.IsNullOrWhiteSpace(path)||String.IsNullOrWhiteSpace(name)||name.Length>200||(recursive!="true"&&recursive!="false")||result.Count>=1000)throw new InvalidDataException("Invalid enterprise library definition.");
                path=PathValue(path,baseDirectory);if(!seen.Add(path.TrimEnd(Path.DirectorySeparatorChar,Path.AltDirectorySeparatorChar)))throw new InvalidDataException("Enterprise library root is duplicated.");
                result.Add(new LibrarySourceDefinition{Root=path,Name=name,Recursive=recursive=="true",Managed=true});
            }
            return result;
        }
        internal static string LegacyFolder(string text)
        {
            using(var reader=XmlReader.Create(new StringReader(text),new XmlReaderSettings{DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=40000000}))
            {
                reader.MoveToContent();if(reader.LocalName!="Library"||reader.NamespaceURI!="urn:powerpoint-intelligence:local-library:v1"||reader.GetAttribute("version")!="1"||String.IsNullOrWhiteSpace(reader.GetAttribute("folder")))throw new InvalidDataException("Legacy library registration is unsupported.");
                return PathPolicy.Selected(reader.GetAttribute("folder"));
            }
        }

        internal string MigrateLegacy(PiidxCatalog catalog,CancellationToken cancellation)
        {
            if(!File.Exists(LegacyCachePath))return null;
            if(new FileInfo(LegacyCachePath).Length>40000000)throw new InvalidDataException("Legacy registration exceeds migration limits; original file preserved.");
            string hash=PresentationSource.Fingerprint(LegacyCachePath,cancellation),receipt=Path.Combine(Path.GetDirectoryName(catalog.IndexPath),"legacy-migration-"+hash+".xml");
            if(File.Exists(receipt))return null;
            string requested=LegacyFolder(File.ReadAllText(LegacyCachePath));string root=PathPolicy.ApprovedRoot(requested,Sources.Select(source=>source.Root).Concat(catalog.Libraries().Select(library=>library.Root)));if(!Directory.Exists(root))throw new DirectoryNotFoundException("Legacy library source folder is unavailable; original registration preserved.");
            string id=catalog.Register(root,Path.GetFileName(root.TrimEnd(Path.DirectorySeparatorChar))+" (migrated)",true);catalog.Refresh(id,cancellation);catalog.VerifyIntegrity();
            if(PresentationSource.Fingerprint(LegacyCachePath,cancellation)!=hash)throw new IOException("Legacy registration changed while migrating; no receipt written.");
            var doc=new XDocument(new XElement("Migration",new XAttribute("version","1"),new XAttribute("source",LegacyCachePath),new XAttribute("sourceSha256",hash),new XAttribute("libraryId",id),new XAttribute("index",catalog.IndexPath),new XAttribute("root",root),new XAttribute("utc",DateTime.UtcNow.ToString("o",CultureInfo.InvariantCulture))));
            using(var file=new FileStream(receipt,FileMode.CreateNew,FileAccess.Write))doc.Save(file);
            return id;
        }
    }
}