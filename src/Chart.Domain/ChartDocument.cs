using System.Collections.Generic;
using Newtonsoft.Json;

namespace PowerPointIntelligence.Domain
{
    // Version 2 is distinct from the S0 XML DTO's version 1. No host is migrated implicitly.
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ChartDocument
    {
        [JsonProperty("schemaVersion", Required = Required.Always, Order = 0)] public int SchemaVersion { get; set; }
        [JsonProperty("chartId", Required = Required.Always, Order = 1)] public string ChartId { get; set; }
        [JsonProperty("revision", Required = Required.Always, Order = 2)] public int Revision { get; set; }
        [JsonProperty("chartType", Required = Required.Always, Order = 3)] public string ChartType { get; set; }
        [JsonProperty("dataset", Required = Required.Always, Order = 4)] public EmbeddedDataset Dataset { get; set; }
        [JsonProperty("series", Required = Required.Always, Order = 5)] public List<ChartSeries> Series { get; set; }
        [JsonProperty("axes", Required = Required.Always, Order = 6)] public List<ChartAxis> Axes { get; set; }
        [JsonProperty("layout", Required = Required.Always, Order = 7)] public ChartLayout Layout { get; set; }
        [JsonProperty("style", Required = Required.Always, Order = 8)] public ChartStyle Style { get; set; }
    }
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class EmbeddedDataset
    {
        [JsonProperty("id", Required = Required.Always, Order = 0)] public string Id { get; set; }
        [JsonProperty("categories", Required = Required.Always, Order = 1)] public List<ChartCategory> Categories { get; set; }
    }
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ChartCategory
    {
        [JsonProperty("id", Required = Required.Always, Order = 0)] public string Id { get; set; }
        [JsonProperty("label", Required = Required.Always, Order = 1)] public string Label { get; set; }
    }
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ChartSeries
    {
        [JsonProperty("id", Required = Required.Always, Order = 0)] public string Id { get; set; }
        [JsonProperty("name", Required = Required.Always, Order = 1)] public string Name { get; set; }
        [JsonProperty("values", Required = Required.Always, Order = 2)] public List<double?> Values { get; set; }
    }
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ChartAxis
    {
        [JsonProperty("id", Required = Required.Always, Order = 0)] public string Id { get; set; }
        [JsonProperty("role", Required = Required.Always, Order = 1)] public string Role { get; set; }
        [JsonProperty("scale", Required = Required.Always, Order = 2)] public string Scale { get; set; }
        [JsonProperty("numberFormat", Required = Required.Always, Order = 3)] public string NumberFormat { get; set; }
    }
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ChartLayout
    {
        [JsonProperty("unit", Required = Required.Always, Order = 0)] public string Unit { get; set; }
        [JsonProperty("x", Required = Required.Always, Order = 1)] public double X { get; set; }
        [JsonProperty("y", Required = Required.Always, Order = 2)] public double Y { get; set; }
        [JsonProperty("width", Required = Required.Always, Order = 3)] public double Width { get; set; }
        [JsonProperty("height", Required = Required.Always, Order = 4)] public double Height { get; set; }
    }
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ChartStyle
    {
        [JsonProperty("showLegend", Required = Required.Always, Order = 0)] public bool ShowLegend { get; set; }
        [JsonProperty("seriesColors", Required = Required.Always, Order = 1)] public List<string> SeriesColors { get; set; }
    }
}
