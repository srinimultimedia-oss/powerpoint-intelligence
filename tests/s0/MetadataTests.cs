using System;
using System.IO;
using PowerPointIntelligence.Host;

internal static class MetadataTests
{
    private static int passed;
    private static void Test(string name, Action action) { action(); Console.WriteLine("PASS " + name); passed++; }
    private static void Reject(Action action)
    {
        try { action(); } catch (InvalidDataException) { return; } catch (InvalidOperationException) { return; }
        throw new Exception("Expected metadata rejection.");
    }
    private static void Equal(object expected, object actual)
    { if (!Equals(expected, actual)) throw new Exception("Round-trip mismatch."); }
    public static int Main()
    {
        try
        {
            Test("round trip preserves identity, revision, labels, signed and missing values", () => {
                var model = MetadataCodec.Sample(); model.Revision = 7;
                model.Categories[0] = "日本語 & <category>"; model.Series[0].Values[0] = -1.25;
                model.Series[1].Values[2] = null;
                var restored = MetadataCodec.Decode(MetadataCodec.Encode(model));
                Equal(model.ChartId, restored.ChartId); Equal(7, restored.Revision);
                Equal(model.Categories[0], restored.Categories[0]); Equal(-1.25, restored.Series[0].Values[0]);
                Equal(null, restored.Series[1].Values[2]); Equal("actual", restored.Series[0].Id);
            });
            Test("invalid GUID", () => { var m = MetadataCodec.Sample(); m.ChartId = "invalid"; Reject(() => MetadataCodec.Encode(m)); });
            Test("empty GUID", () => { var m = MetadataCodec.Sample(); m.ChartId = Guid.Empty.ToString(); Reject(() => MetadataCodec.Encode(m)); });
            Test("unknown schema", () => { var m = MetadataCodec.Sample(); m.SchemaVersion = 2; Reject(() => MetadataCodec.Encode(m)); });
            Test("unknown chart type", () => { var m = MetadataCodec.Sample(); m.ChartType = "pie"; Reject(() => MetadataCodec.Encode(m)); });
            Test("negative revision", () => { var m = MetadataCodec.Sample(); m.Revision = -1; Reject(() => MetadataCodec.Encode(m)); });
            Test("dimension mismatch", () => { var m = MetadataCodec.Sample(); m.Series[0].Values.Clear(); Reject(() => MetadataCodec.Encode(m)); });
            Test("duplicate series ID", () => { var m = MetadataCodec.Sample(); m.Series[0].Id = m.Series[1].Id; Reject(() => MetadataCodec.Encode(m)); });
            Test("NaN", () => { var m = MetadataCodec.Sample(); m.Series[0].Values[0] = Double.NaN; Reject(() => MetadataCodec.Encode(m)); });
            Test("infinity", () => { var m = MetadataCodec.Sample(); m.Series[0].Values[0] = Double.PositiveInfinity; Reject(() => MetadataCodec.Encode(m)); });
            Test("empty categories", () => { var m = MetadataCodec.Sample(); m.Categories.Clear(); Reject(() => MetadataCodec.Encode(m)); });
            Test("null series", () => { var m = MetadataCodec.Sample(); m.Series[0] = null; Reject(() => MetadataCodec.Encode(m)); });
            Test("missing metadata", () => Reject(() => MetadataCodec.Decode(null)));
            Test("malformed XML", () => Reject(() => MetadataCodec.Decode("<broken>")));
            Test("DTD and external entities disabled", () => Reject(() => MetadataCodec.Decode("<!DOCTYPE ChartDocument [<!ENTITY x SYSTEM 'file:///not-read'>]><ChartDocument xmlns='" + MetadataCodec.Namespace + "'>&x;</ChartDocument>")));
            Test("oversize bounded", () => Reject(() => MetadataCodec.Decode(new string('x', MetadataCodec.MaxCharacters + 1))));
            Test("unknown root namespace", () => Reject(() => MetadataCodec.Decode(MetadataCodec.Encode(MetadataCodec.Sample()).Replace(MetadataCodec.Namespace, "urn:unknown"))));
            Console.WriteLine("Metadata tests passed: " + passed); return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
}
