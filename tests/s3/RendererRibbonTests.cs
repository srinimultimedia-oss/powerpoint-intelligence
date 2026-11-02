using System;
using System.Linq;
using PowerPointIntelligence.AddIn;
using PowerPointIntelligence.Domain;
using PowerPointIntelligence.Rendering;

internal static class RendererRibbonTests
{
    private static int passed;
    private static void Require(bool condition, string name)
    { if (!condition) throw new Exception("Renderer ribbon: " + name); Console.WriteLine("PASS renderer ribbon " + name); passed++; }
    internal static void Run(string canonical, string scene)
    {
        var baseline = ClusteredProjection.Validate(canonical, scene);
        var first = RendererRibbonCommand.Prepare();
        var second = RendererRibbonCommand.Prepare();
        Require(first.Model.ChartId != second.Model.ChartId && first.Model.ChartId != baseline.Model.ChartId, "assigns fresh identity on each insertion");
        Require(first.Scene.ChartId == first.Model.ChartId && first.Scene.Nodes.All(n => n.Id.StartsWith(first.Model.ChartId + "/", StringComparison.Ordinal)), "remaps every semantic node consistently");
        Require(first.Scene.RenderHash == SceneCodec.ComputeHash(first.Scene) && first.Scene.RenderHash != second.Scene.RenderHash, "recomputes identity-dependent scene hash");
        // Normalize only chart identity to compare every other canonical field.
        first.Model.ChartId = baseline.Model.ChartId;
        Require(ChartCodec.Encode(first.Model) == ChartCodec.Encode(baseline.Model), "retains every canonical field, signed/missing values and Unicode");
        first.Scene.ChartId = baseline.Scene.ChartId;
        foreach (var node in first.Scene.Nodes) node.Id = baseline.Scene.ChartId + node.Id.Substring(node.Id.IndexOf('/'));
        first.Scene.RenderHash = SceneCodec.ComputeHash(first.Scene);
        Require(SceneCodec.Encode(first.Scene) == SceneCodec.Encode(baseline.Scene), "retains exact geometry, colors, revision and omission diagnostics");
        second.Model.Series[0].Values[0] = 999;
        second.Scene.Nodes.Clear();
        var third = RendererRibbonCommand.Prepare();
        Require(third.Model.Series[0].Values[0] == baseline.Model.Series[0].Values[0] && third.Scene.Nodes.Count == baseline.Scene.Nodes.Count, "isolates resource snapshots between invocations");
        bool rejected = false;
        try { RendererRibbonCommand.Insert(null); } catch (ArgumentNullException) { rejected = true; }
        Require(rejected, "rejects absent host before Office access");
        var linkedModel=ChartCodec.Decode(canonical);
        linkedModel.Series[0].Values[0]=137.5;
        linkedModel.Dataset.Categories[0].Label="Linked source category";
        var linked=RendererRibbonCommand.Prepare(linkedModel);
        Require(linked.Model.Series[0].Values[0]==137.5&&linked.Model.Dataset.Categories[0].Label=="Linked source category","linked model values and categories replace resource samples");
        Require(linked.Scene.RenderHash==SceneCodec.ComputeHash(linked.Scene)&&linked.Scene.Nodes.Count==baseline.Scene.Nodes.Count,"linked source generates validated signed and missing geometry");
        Console.WriteLine("Renderer ribbon preparation tests passed: " + passed + ". Installed ribbon/UI behavior not tested.");
    }
}
