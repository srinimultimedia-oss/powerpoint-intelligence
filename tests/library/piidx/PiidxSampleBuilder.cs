using System;
using System.IO;
using System.Threading;
using PowerPointIntelligence.Library;
public static class PiidxSampleBuilder
{
    public static string Build(string output,string original)
    {
        Directory.CreateDirectory(output);string sources=Path.Combine(output,"sources");Directory.CreateDirectory(sources);string source=Path.Combine(sources,"sample-native.pptx");
        File.Copy(original,source,false);string hash=PresentationSource.Fingerprint(source,CancellationToken.None);
        var catalog=PiidxCatalog.OpenOrCreate(Path.Combine(output,"sample-native.piidx"));string library=catalog.Register(sources,"Previously generated native test presentation",true);var result=catalog.Refresh(library,CancellationToken.None);catalog.VerifyIntegrity();
        if(result.SlidesIndexed<1||catalog.Search(new LibrarySearchRequest(),CancellationToken.None).Total!=result.SlidesIndexed)throw new Exception("Native-source sample index is incomplete.");
        if(hash!=PresentationSource.Fingerprint(source,CancellationToken.None))throw new Exception("Sample source changed.");
        return "PASS sample-native.piidx; source is an owner-generated Office test deck; slides="+result.SlidesIndexed+"; sourceSHA256="+hash+". Existing source copied unchanged; no new Office execution or preservation acceptance is claimed.";
    }
}