using System;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Serialization;

namespace PowerPointIntelligence.Host
{
    // S0 wire format only. The full S1 model/migration system is deliberately deferred.
    public static class MetadataCodec
    {
        public const string Namespace = "urn:powerpoint-intelligence:s0:chart:v1";
        public const int MaxCharacters = 1048576;
        private static readonly XmlSerializer Serializer = new XmlSerializer(typeof(ChartDocument), Namespace);

        public static void Validate(ChartDocument chart)
        {
            Guid id;
            if (chart == null || chart.SchemaVersion != 1 || !Guid.TryParse(chart.ChartId, out id) || id == Guid.Empty ||
                chart.ChartType != "clustered-column" || chart.Revision < 0)
                throw new InvalidDataException("Unsupported or invalid chart metadata.");
            if (chart.Categories == null || chart.Categories.Count < 1 || chart.Categories.Count > 1000 ||
                chart.Categories.Any(x => String.IsNullOrWhiteSpace(x) || x.Length > 1024) ||
                chart.Series == null || chart.Series.Count < 1 || chart.Series.Count > 100)
                throw new InvalidDataException("Invalid chart dimensions or category labels.");
            if (chart.Series.Any(s => s == null || String.IsNullOrWhiteSpace(s.Id) || s.Id.Length > 128 ||
                String.IsNullOrWhiteSpace(s.Name) || s.Name.Length > 1024 || s.Values == null ||
                s.Values.Count != chart.Categories.Count ||
                s.Values.Any(v => v.HasValue && (Double.IsNaN(v.Value) || Double.IsInfinity(v.Value)))) ||
                chart.Series.Select(s => s.Id).Distinct(StringComparer.Ordinal).Count() != chart.Series.Count)
                throw new InvalidDataException("Invalid chart series.");
        }

        public static string Encode(ChartDocument chart)
        {
            Validate(chart);
            using (var writer = new StringWriter(System.Globalization.CultureInfo.InvariantCulture))
            {
                Serializer.Serialize(writer, chart);
                var xml = writer.ToString();
                if (xml.Length > MaxCharacters) throw new InvalidDataException("Metadata exceeds S0 limit.");
                return xml;
            }
        }

        public static ChartDocument Decode(string xml)
        {
            if (String.IsNullOrWhiteSpace(xml) || xml.Length > MaxCharacters)
                throw new InvalidDataException("Missing or oversized chart metadata.");
            var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null, MaxCharactersInDocument = MaxCharacters };
            using (var reader = XmlReader.Create(new StringReader(xml), settings))
            {
                var chart = (ChartDocument)Serializer.Deserialize(reader);
                Validate(chart);
                return chart;
            }
        }

        public static ChartDocument Sample()
        {
            return new ChartDocument {
                Categories = new System.Collections.Generic.List<string> { "Q1", "Q2", "Q3" },
                Series = new System.Collections.Generic.List<ChartSeries> {
                    new ChartSeries { Id = "actual", Name = "Actual", Values = new System.Collections.Generic.List<double?> { 10, 20, 30 } },
                    new ChartSeries { Id = "plan", Name = "Plan", Values = new System.Collections.Generic.List<double?> { 15, 18, 35 } }
                }
            };
        }
    }
}
