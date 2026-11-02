using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PowerPointIntelligence.Rendering;
using PowerPointIntelligence.Tests.Security;

internal static class SceneContractTests
{
    private static int passed;
    private static void Require(bool valid, string name) { if (!valid) throw new Exception(name); }
    private static void Test(string name, Action action) { action(); Console.WriteLine("PASS " + name); passed++; }
    private static void Reject(Action action)
    { try { action(); } catch (InvalidDataException) { return; } throw new Exception("Expected InvalidDataException."); }
    public static int Main(string[] args)
    {
        try
        {
            Console.WriteLine("Scene contract runner process=" + (Environment.Is64BitProcess ? "x64" : "x86"));
            string golden = File.ReadAllText(args[0]), output = args[1], corpus = args[2];
            Directory.CreateDirectory(output);
            Test("all six geometry types decode to whitelisted DTOs", () => {
                var model = SceneCodec.Decode(golden);
                Require(model.Nodes.Count == 6 && model.Nodes[0].Geometry is BoxGeometry && model.Nodes[2].Geometry is PathGeometry && model.Nodes[4].Geometry is SectorGeometry && model.Nodes[5].Geometry is TextGeometry, "geometry DTOs");
                Require(((TextGeometry)model.Nodes[5].Geometry).Text == "Δ Revenue 😀\nLine 2", "Unicode payload");
            });
            Test("golden hash verified without mutating payload", () => {
                var model = SceneCodec.Decode(golden); string before = SceneCodec.Encode(model);
                Require(SceneCodec.ComputeHash(model) == model.RenderHash, "hash"); Require(SceneCodec.Encode(model) == before, "mutation");
                File.WriteAllText(Path.Combine(output, "all-primitives.v1.json"), before);
            });
            Test("culture invariant wire and hash", () => {
                var previous = CultureInfo.CurrentCulture;
                try { CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR"); var model = SceneCodec.Decode(golden); Require(SceneCodec.ComputeHash(model) == model.RenderHash, "culture hash"); SceneCodec.Decode(SceneCodec.Encode(model)); }
                finally { CultureInfo.CurrentCulture = previous; }
            });
            Test("changed geometry and revision cannot reuse a hash", () => {
                var model = SceneCodec.Decode(golden); ((BoxGeometry)model.Nodes[0].Geometry).X++;
                Reject(() => SceneCodec.Encode(model)); model.RenderHash = SceneCodec.ComputeHash(model); SceneCodec.Decode(SceneCodec.Encode(model));
                model.ModelRevision++; Reject(() => SceneCodec.Encode(model));
            });
            Test("typed models cannot bypass structural validation", () => {
                var model = SceneCodec.Decode(golden); ((BoxGeometry)model.Nodes[0].Geometry).X = Double.NaN; Reject(() => SceneCodec.ComputeHash(model));
                model = SceneCodec.Decode(golden); model.Nodes[0].Paint.Opacity = Double.PositiveInfinity; Reject(() => SceneCodec.Encode(model));
                model = SceneCodec.Decode(golden); model.Nodes[0].Geometry.Type = "script"; Reject(() => SceneCodec.ComputeHash(model));
                model = SceneCodec.Decode(golden); model.Nodes[0].Paint = null; Reject(() => SceneCodec.Encode(model));
            });
            Test("unsealed drafts require a computed hash before encoding", () => {
                var model = SceneCodec.Decode(golden); model.RenderHash = null;
                Reject(() => SceneCodec.Encode(model)); model.RenderHash = SceneCodec.ComputeHash(model); SceneCodec.Decode(SceneCodec.Encode(model));
            });
            Test("size and nesting caps", () => {
                Reject(() => SceneCodec.Decode(new string(' ', SceneCodec.MaxCharacters + 1)));
                Reject(() => SceneCodec.Decode(new string('[', 33) + "0" + new string(']', 33)));
            });
            Test("C# producer independently seals changed doubles and Unicode", () => {
                var model = SceneCodec.Decode(golden); model.ModelRevision = 8;
                ((BoxGeometry)model.Nodes[0].Geometry).X = -0d;
                ((BoxGeometry)model.Nodes[1].Geometry).Y = -1.2345678901234567e-200;
                ((TextGeometry)model.Nodes[5].Geometry).Text = "مرحبا 😀 \"quoted\" <&>";
                ((SectorGeometry)model.Nodes[4].Geometry).Sweep = Double.Epsilon;
                model.RenderHash = SceneCodec.ComputeHash(model);
                File.WriteAllText(Path.Combine(output, "csharp-produced.v1.json"), SceneCodec.Encode(model));
            });
            var rows = new List<object>();
            var cases = JArray.Parse(File.ReadAllText(Path.Combine(corpus, "cases.json")));
            SceneArtifactPath.ValidateNames(output, cases.Select(item => (string)item["name"]));
            foreach (var item in cases)
            {
                string name = (string)item["name"], json = (string)item["json"]; bool expected = (bool)item["valid"];
                Test("shared fixture " + name, () => {
                    RenderResult model = null;
                    if (expected) { model = SceneCodec.Decode(json); File.WriteAllText(SceneArtifactPath.Resolve(output, name), SceneCodec.Encode(model)); }
                    else Reject(() => SceneCodec.Decode(json));
                    rows.Add(new { Name = name, ExpectedValid = expected, Result = "PASS" });
                });
            }
            File.WriteAllText(Path.Combine(output, "results.json"), JsonConvert.SerializeObject(rows, Formatting.Indented));
            Console.WriteLine("Scene contract tests passed: " + passed); return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex.GetType().Name + ": " + ex.Message); return 1; }
    }
}
