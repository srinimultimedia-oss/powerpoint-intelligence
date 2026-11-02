using System.Collections.Generic;
using Newtonsoft.Json;

namespace PowerPointIntelligence.Rendering
{
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class RenderResult
    {
        [JsonProperty("schemaVersion", Required = Required.Always)] public int SchemaVersion { get; set; }
        [JsonProperty("chartId", Required = Required.Always)] public string ChartId { get; set; }
        [JsonProperty("modelRevision", Required = Required.Always)] public int ModelRevision { get; set; }
        [JsonProperty("unit", Required = Required.Always)] public string Unit { get; set; }
        [JsonProperty("plot", Required = Required.Always)] public ScenePlot Plot { get; set; }
        [JsonProperty("semanticScene", Required = Required.Always)] public List<SceneNode> Nodes { get; set; }
        [JsonProperty("diagnostics", Required = Required.Always)] public List<SceneDiagnostic> Diagnostics { get; set; }
        // An unsealed in-memory draft may have no hash; SceneCodec.Encode/Decode
        // still require a valid, verified hash before the DTO crosses a boundary.
        [JsonProperty("renderHash", Required = Required.AllowNull)] public string RenderHash { get; set; }
    }
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ScenePlot
    {
        [JsonProperty("x", Required = Required.Always)] public double X { get; set; }
        [JsonProperty("y", Required = Required.Always)] public double Y { get; set; }
        [JsonProperty("width", Required = Required.Always)] public double Width { get; set; }
        [JsonProperty("height", Required = Required.Always)] public double Height { get; set; }
        [JsonProperty("unit", Required = Required.Always)] public string Unit { get; set; }
    }
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class SceneNode
    {
        [JsonProperty("id", Required = Required.Always)] public string Id { get; set; }
        [JsonProperty("geometry", Required = Required.Always)] public SceneGeometry Geometry { get; set; }
        [JsonProperty("paint", Required = Required.Always)] public ScenePaint Paint { get; set; }
    }
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ScenePaint
    {
        [JsonProperty("fill", Required = Required.AllowNull)] public string Fill { get; set; }
        [JsonProperty("stroke", Required = Required.AllowNull)] public string Stroke { get; set; }
        [JsonProperty("strokeWidth", Required = Required.Always)] public double StrokeWidth { get; set; }
        [JsonProperty("opacity", Required = Required.Always)] public double Opacity { get; set; }
    }
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class SceneDiagnostic
    {
        [JsonProperty("code", Required = Required.Always)] public string Code { get; set; }
        [JsonProperty("nodeId", Required = Required.AllowNull)] public string NodeId { get; set; }
    }
    [JsonObject(MemberSerialization.OptIn)]
    public abstract class SceneGeometry
    { [JsonProperty("type", Required = Required.Always)] public string Type { get; set; } }
    [JsonObject(MemberSerialization.OptIn)]
    public class BoxGeometry : SceneGeometry
    {
        [JsonProperty("x", Required = Required.Always)] public double X { get; set; }
        [JsonProperty("y", Required = Required.Always)] public double Y { get; set; }
        [JsonProperty("width", Required = Required.Always)] public double Width { get; set; }
        [JsonProperty("height", Required = Required.Always)] public double Height { get; set; }
    }
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class TextGeometry : BoxGeometry
    {
        [JsonProperty("text", Required = Required.Always)] public string Text { get; set; }
        [JsonProperty("fontFamily", Required = Required.Always)] public string FontFamily { get; set; }
        [JsonProperty("fontSize", Required = Required.Always)] public double FontSize { get; set; }
        [JsonProperty("alignment", Required = Required.Always)] public string Alignment { get; set; }
    }
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class PathGeometry : SceneGeometry
    { [JsonProperty("points", Required = Required.Always)] public List<ScenePoint> Points { get; set; } }
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ScenePoint
    {
        [JsonProperty("x", Required = Required.Always)] public double X { get; set; }
        [JsonProperty("y", Required = Required.Always)] public double Y { get; set; }
    }
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class SectorGeometry : SceneGeometry
    {
        [JsonProperty("cx", Required = Required.Always)] public double Cx { get; set; }
        [JsonProperty("cy", Required = Required.Always)] public double Cy { get; set; }
        [JsonProperty("radius", Required = Required.Always)] public double Radius { get; set; }
        [JsonProperty("startAngle", Required = Required.Always)] public double StartAngle { get; set; }
        [JsonProperty("sweep", Required = Required.Always)] public double Sweep { get; set; }
    }
}
