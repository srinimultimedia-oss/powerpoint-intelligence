using System.Collections.Generic;
using Newtonsoft.Json;
namespace PowerPointIntelligence.Domain
{
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ChartDocumentV3
    {
        [JsonProperty("schemaVersion", Required = Required.Always)] public int SchemaVersion { get; set; }
        [JsonProperty("chartId", Required = Required.Always)] public string ChartId { get; set; }
        [JsonProperty("revision", Required = Required.Always)] public int Revision { get; set; }
        [JsonProperty("chartType", Required = Required.Always)] public string ChartType { get; set; }
        [JsonProperty("dataset", Required = Required.Always)] public EmbeddedDataset Dataset { get; set; }
        [JsonProperty("series", Required = Required.Always)] public List<ChartSeries> Series { get; set; }
        [JsonProperty("axes", Required = Required.Always)] public List<ChartAxis> Axes { get; set; }
        [JsonProperty("layout", Required = Required.Always)] public ChartLayout Layout { get; set; }
        [JsonProperty("style", Required = Required.Always)] public ChartStyleV3 Style { get; set; }
        [JsonProperty("options", Required = Required.Always)] public ChartOptions Options { get; set; }
    }
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ChartStyleV3
    {
        [JsonProperty("showLegend", Required = Required.Always)] public bool ShowLegend { get; set; }
        [JsonProperty("seriesColors", Required = Required.Always)] public List<string> SeriesColors { get; set; }
        [JsonProperty("categoryColors", Required = Required.Always)] public List<string> CategoryColors { get; set; }
    }
    [JsonObject(MemberSerialization.OptIn)]
    public abstract class ChartOptions
    { [JsonProperty("kind", Required = Required.Always)] public string Kind { get; set; } }
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class BarOptions : ChartOptions
    { [JsonProperty("gapRatio", Required = Required.Always)] public double GapRatio { get; set; } }
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class LineOptions : ChartOptions
    {
        [JsonProperty("gapPolicy", Required = Required.Always)] public string GapPolicy { get; set; }
        [JsonProperty("includeZero", Required = Required.Always)] public bool IncludeZero { get; set; }
    }
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class AreaOptions : ChartOptions
    {
        [JsonProperty("gapPolicy", Required = Required.Always)] public string GapPolicy { get; set; }
        [JsonProperty("baseline", Required = Required.Always)] public double Baseline { get; set; }
    }
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class PieOptions : ChartOptions
    { [JsonProperty("startAngle", Required = Required.Always)] public double StartAngle { get; set; } }
}
