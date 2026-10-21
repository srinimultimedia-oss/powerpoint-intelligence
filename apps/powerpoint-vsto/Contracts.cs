using System;
using System.Collections.Generic;

namespace PowerPointIntelligence.Host
{
    // Compile this contract into a VSTO project on Windows with .NET Framework 4.8.
    public interface IPowerPointHost
    {
        string GetSelectedSlideId();
        IReadOnlyList<string> GetSelectedObjectIds();
        IReadOnlyList<string> InsertEditableChart(ChartDocument chart);
        void UpdateEditableChart(string chartId, ChartDocument chart);
        ChartDocument ReadChartMetadata(string chartId);
        void SaveChartMetadata(ChartDocument chart);
    }

    public sealed class ChartDocument
    {
        public int SchemaVersion { get; set; } = 1;
        public string ChartId { get; set; } = Guid.NewGuid().ToString("D");
        public string ChartType { get; set; } = "clustered-column";
        public int Revision { get; set; }
        public List<string> Categories { get; set; } = new List<string>();
        public List<ChartSeries> Series { get; set; } = new List<ChartSeries>();
    }

    public sealed class ChartSeries
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public List<double?> Values { get; set; } = new List<double?>();
    }
}
