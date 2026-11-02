using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PowerPointIntelligence.Contracts;

namespace PowerPointIntelligence.Domain
{
    public static class ChartV3Codec
    {
        public const int Version = 3;
        public const int MaxCharacters = ChartCodec.MaxCharacters;
        public static readonly string[] BarTypes = { "clustered-column", "clustered-bar", "stacked-column", "stacked-bar", "percent-column", "percent-bar" };
        private static readonly Regex Id = new Regex(@"\A[A-Za-z0-9][A-Za-z0-9:._-]{0,127}\z", RegexOptions.CultureInvariant);
        private static readonly Regex Color = new Regex(@"\A#[0-9A-Fa-f]{6}\z", RegexOptions.CultureInvariant);
        private static void Require(bool condition, string code) { if (!condition) throw new InvalidDataException(code); }
        private static JsonSerializer Serializer()
        {
            var serializer = JsonSerializer.Create(new JsonSerializerSettings { TypeNameHandling = TypeNameHandling.None, MissingMemberHandling = MissingMemberHandling.Error, Culture = CultureInfo.InvariantCulture, MaxDepth = 32, DateParseHandling = DateParseHandling.None, Formatting = Formatting.None });
            serializer.Converters.Add(new OptionsReader()); return serializer;
        }
        private static Type Resolve(JToken value, Type type)
        {
            if (type != typeof(ChartOptions)) return type;
            Require(value.Type == JTokenType.Object && value["kind"] != null && value["kind"].Type == JTokenType.String, "options.kind");
            switch ((string)value["kind"])
            {
                case "bar": return typeof(BarOptions); case "line": return typeof(LineOptions); case "area": return typeof(AreaOptions); case "pie": return typeof(PieOptions);
                default: throw new InvalidDataException("options.kind");
            }
        }
        public static void Validate(ChartDocumentV3 model)
        {
            Require(model != null && model.SchemaVersion == Version, "v3.version");
            bool bar = BarTypes.Contains(model.ChartType, StringComparer.Ordinal), pie = model.ChartType == "pie";
            Require(bar || model.ChartType == "line" || model.ChartType == "area" || pie, "v3.chart-type");
            var common = new ChartDocument { ChartId = model.ChartId, Revision = model.Revision, Dataset = model.Dataset, Series = model.Series, Axes = model.Axes, Layout = model.Layout, Style = model.Style == null ? null : new ChartStyle { ShowLegend = model.Style.ShowLegend, SeriesColors = model.Style.SeriesColors } };
            ChartCodec.ValidateContent(common, !pie, text => ContractText.Valid(text, 1024) && ContractText.NonWhitespace(text));
            var ids = new[] { model.ChartId, model.Dataset.Id }.Concat(model.Dataset.Categories.Select(c => c.Id)).Concat(model.Series.Select(s => s.Id)).Concat(model.Axes.Select(a => a.Id));
            Require(ids.All(id => Id.IsMatch(id)) && model.Style.SeriesColors.All(c => Color.IsMatch(c)), "v3.id-or-color");
            Require(model.Style.CategoryColors != null && model.Style.CategoryColors.Count == (pie ? model.Dataset.Categories.Count : 0) && model.Style.CategoryColors.All(c => c != null && Color.IsMatch(c)), "v3.category-colors");
            Require(model.Options != null, "v3.options");
            if (bar)
            {
                var option = model.Options as BarOptions;
                Require(option != null && option.Kind == "bar" && !Double.IsNaN(option.GapRatio) && option.GapRatio >= 0 && option.GapRatio < 1, "v3.bar-options");
                if (model.ChartType.StartsWith("percent-", StringComparison.Ordinal)) Require(model.Series.All(s => s.Values.All(v => !v.HasValue || v.Value >= 0)), "v3.signed-100-policy-required");
            }
            else if (model.ChartType == "line")
            { var option = model.Options as LineOptions; Require(option != null && option.Kind == "line" && Gap(option.GapPolicy), "v3.line-options"); }
            else if (model.ChartType == "area")
            { var option = model.Options as AreaOptions; Require(option != null && option.Kind == "area" && Gap(option.GapPolicy) && option.Baseline == 0, "v3.area-options"); }
            else
            {
                var option = model.Options as PieOptions;
                Require(option != null && option.Kind == "pie" && !Double.IsNaN(option.StartAngle) && option.StartAngle >= -360 && option.StartAngle <= 360, "v3.pie-options");
                Require(model.Series.Count == 1 && model.Series[0].Values.All(v => !v.HasValue || v.Value >= 0), "v3.pie-data");
            }
        }
        private static bool Gap(string gap) { return gap == "break" || gap == "connect" || gap == "zero"; }
        public static ChartDocumentV3 Decode(string json)
        {
            var token = BoundedJson.Parse(json, MaxCharacters, "v3"); ChartCodec.CheckShape(token, typeof(ChartDocumentV3), Resolve);
            try { var model = token.ToObject<ChartDocumentV3>(Serializer()); Validate(model); return model; }
            catch (JsonException) { throw new InvalidDataException("v3.dto"); }
        }
        public static string Encode(ChartDocumentV3 model)
        {
            Validate(model);
            try {
                var token = JObject.FromObject(model, Serializer()); ChartCodec.CheckShape(token, typeof(ChartDocumentV3), Resolve);
                string json = token.ToString(Formatting.None); Require(json.Length <= MaxCharacters, "v3.json-size"); return json;
            } catch (JsonException) { throw new InvalidDataException("v3.dto"); }
        }
        public static ChartDocumentV3 FromV2(string json, double gapRatio)
        {
            BoundedJson.Parse(json, ChartCodec.MaxCharacters, "v2"); var old = ChartCodec.Decode(json);
            var migrated = new ChartDocumentV3 { SchemaVersion = Version, ChartId = old.ChartId, Revision = old.Revision, ChartType = old.ChartType, Dataset = old.Dataset, Series = old.Series, Axes = old.Axes, Layout = old.Layout,
                Style = new ChartStyleV3 { ShowLegend = old.Style.ShowLegend, SeriesColors = old.Style.SeriesColors, CategoryColors = new List<string>() }, Options = new BarOptions { Kind = "bar", GapRatio = gapRatio } };
            return Decode(Encode(migrated)); // Independent snapshot; no input or defaults mutated.
        }
        private sealed class OptionsReader : JsonConverter
        {
            public override bool CanWrite { get { return false; } }
            public override bool CanConvert(Type type) { return type == typeof(ChartOptions); }
            public override object ReadJson(JsonReader reader, Type type, object existing, JsonSerializer serializer)
            { var value = JObject.Load(reader); return value.ToObject(Resolve(value, typeof(ChartOptions)), serializer); }
            public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer) { throw new NotSupportedException(); }
        }
    }
}
