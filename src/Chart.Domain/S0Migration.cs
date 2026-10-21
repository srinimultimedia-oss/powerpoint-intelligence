using System;
using System.Collections.Generic;
using System.Linq;
using PowerPointIntelligence.Host;

namespace PowerPointIntelligence.Domain
{
    public static class S0Migration
    {
        // Defaults are explicit inputs: S0 XML never stored native layout/colors. Callers
        // must capture them from Office or obtain an explicit user-selected fallback.
        public static ChartDocument FromXml(string xml, ChartLayout layout, ChartStyle style)
        {
            var old = MetadataCodec.Decode(xml);
            var model = new ChartDocument {
                SchemaVersion = ChartCodec.CurrentVersion, ChartId = old.ChartId, Revision = old.Revision,
                ChartType = old.ChartType,
                Dataset = new EmbeddedDataset { Id = old.ChartId + ":dataset", Categories = old.Categories.Select((label, i) =>
                    new ChartCategory { Id = old.ChartId + ":category:" + i, Label = label }).ToList() },
                Series = old.Series.Select((s, i) => new ChartSeries { Id = old.ChartId + ":series:" + i, Name = s.Name,
                    Values = new List<double?>(s.Values) }).ToList(),
                Axes = new List<ChartAxis> {
                    new ChartAxis { Id = old.ChartId + ":axis:category", Role = "category", Scale = "category", NumberFormat = "General" },
                    new ChartAxis { Id = old.ChartId + ":axis:value", Role = "value", Scale = "linear", NumberFormat = "General" }
                },
                Layout = layout, Style = style
            };
            // Clone via the validated codec: caller-owned defaults and legacy lists never alias.
            return ChartCodec.Decode(ChartCodec.Encode(model));
        }
    }
}
