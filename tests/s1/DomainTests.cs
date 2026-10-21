using System;
using System.IO;
using System.Globalization;
using Newtonsoft.Json.Linq;
using PowerPointIntelligence.Domain;
using Legacy = PowerPointIntelligence.Host;

internal static class DomainTests
{
    private static int passed;
    private static string golden;
    private static void Test(string name, Action action) { action(); Console.WriteLine("PASS " + name); passed++; }
    private static void Require(bool condition) { if (!condition) throw new Exception("Assertion failed."); }
    private static void Reject(Action action)
    { try { action(); } catch (InvalidDataException) { return; } throw new Exception("Expected rejection."); }
    private static void Invalid(string name, Action<JObject> mutation)
    { Test(name, () => { var value = JObject.Parse(golden); mutation(value); Reject(() => ChartCodec.Decode(value.ToString())); }); }
    public static int Main(string[] args)
    {
        try
        {
            golden = File.ReadAllText(args[0]);
            Test("golden JSON round trip preserves every field", () => {
                var encoded = ChartCodec.Encode(ChartCodec.Decode(golden));
                Require(JToken.DeepEquals(JToken.Parse(golden), JToken.Parse(encoded)));
                Require(encoded == ChartCodec.Encode(ChartCodec.Decode(encoded)));
                File.WriteAllText(args[1], encoded);
            });
            Test("culture invariant wire format", () => {
                var prior = CultureInfo.CurrentCulture;
                try { CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
                    Require(ChartCodec.Encode(ChartCodec.Decode(golden)).Contains("-1.25")); }
                finally { CultureInfo.CurrentCulture = prior; }
            });
            Invalid("unknown version", m => m["schemaVersion"] = 99);
            Invalid("missing property", m => m.Remove("revision"));
            Invalid("unknown property", m => m["unknown"] = 1);
            Invalid("nested unknown property", m => m["layout"]["unknown"] = 1);
            Invalid("no numeric string coercion", m => m["revision"] = "4");
            Invalid("negative revision", m => m["revision"] = -1);
            Invalid("fractional revision", m => m["revision"] = 4.5);
            Invalid("revision overflow", m => m["revision"] = 2147483648L);
            Invalid("null dataset", m => m["dataset"] = null);
            Invalid("invalid GUID", m => m["chartId"] = "bad");
            Invalid("empty GUID", m => m["chartId"] = Guid.Empty.ToString());
            Invalid("unsupported chart type", m => m["chartType"] = "pie");
            Invalid("duplicate semantic identity", m => m["series"][0]["id"] = m["dataset"]["categories"][0]["id"].DeepClone());
            Invalid("category count mismatch", m => ((JArray)m["series"][0]["values"]).RemoveAt(0));
            Invalid("empty categories", m => m["dataset"]["categories"] = new JArray());
            Invalid("axis role duplicate", m => { m["axes"][1]["role"] = "category"; m["axes"][1]["scale"] = "category"; });
            Invalid("unsupported axis scale", m => m["axes"][1]["scale"] = "log");
            Invalid("zero width", m => m["layout"]["width"] = 0);
            Invalid("unsupported units", m => m["layout"]["unit"] = "px");
            Invalid("color count mismatch", m => ((JArray)m["style"]["seriesColors"]).RemoveAt(0));
            Invalid("invalid color", m => m["style"]["seriesColors"][0] = "red");
            Test("NaN rejected before serialization", () => { var m = ChartCodec.Decode(golden); m.Series[0].Values[0] = Double.NaN; Reject(() => ChartCodec.Encode(m)); });
            Test("integer-valued JSON number follows JSON Schema semantics", () => {
                Require(ChartCodec.Decode(golden.Replace("\"revision\": 4", "\"revision\": 4.0")).Revision == 4);
            });
            Test("infinite layout rejected", () => { var m = ChartCodec.Decode(golden); m.Layout.X = Double.PositiveInfinity; Reject(() => ChartCodec.Encode(m)); });
            Test("malformed JSON", () => Reject(() => ChartCodec.Decode("{")));
            Test("duplicate JSON property", () => Reject(() => ChartCodec.Decode(golden.Replace("\"revision\": 4", "\"revision\": 4, \"revision\": 5"))));
            Test("comments rejected", () => Reject(() => ChartCodec.Decode("/*comment*/" + golden)));
            Test("trailing document rejected", () => Reject(() => ChartCodec.Decode(golden + " {}")));
            Test("oversized metadata rejected", () => Reject(() => ChartCodec.Decode(new string('x', ChartCodec.MaxCharacters + 1))));
            Test("S0 migration is deterministic and nonmutating", () => {
                var old = Legacy.MetadataCodec.Sample(); old.Revision = 7; old.Series[1].Values[1] = null;
                var xml = Legacy.MetadataCodec.Encode(old);
                var defaults = ChartCodec.Decode(golden);
                var migrated = S0Migration.FromXml(xml, defaults.Layout, defaults.Style);
                Require(migrated.ChartId == old.ChartId && migrated.Revision == 7 && migrated.SchemaVersion == 2);
                Require(migrated.Series[0].Values[0] == 10 && migrated.Series[1].Values[1] == null);
                Require(ChartCodec.Encode(migrated) == ChartCodec.Encode(S0Migration.FromXml(xml, defaults.Layout, defaults.Style)));
                migrated.Layout.Width = 999; Require(defaults.Layout.Width == 600);
                Require(xml == Legacy.MetadataCodec.Encode(old));
            });
            Test("S0 migration refuses missing formatting inputs", () => Reject(() => S0Migration.FromXml(Legacy.MetadataCodec.Encode(Legacy.MetadataCodec.Sample()), null, null)));
            Test("S0 future schema preserved by rejection", () => {
                var xml = Legacy.MetadataCodec.Encode(Legacy.MetadataCodec.Sample()).Replace("<SchemaVersion>1</SchemaVersion>", "<SchemaVersion>99</SchemaVersion>");
                var defaults = ChartCodec.Decode(golden);
                Reject(() => S0Migration.FromXml(xml, defaults.Layout, defaults.Style));
            });
            Console.WriteLine("Domain tests passed: " + passed); return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex.GetType().Name + ": domain tests failed."); return 1; }
    }
}
