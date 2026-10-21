using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace PowerPointIntelligence.Domain
{
    public static class ChartCodec
    {
        public const int CurrentVersion = 2;
        public const int MaxCharacters = 1048576;
        private static readonly Regex IdPattern = new Regex(@"^[A-Za-z0-9][A-Za-z0-9:._-]{0,127}$", RegexOptions.CultureInvariant);
        private static readonly Regex ColorPattern = new Regex(@"^#[0-9A-Fa-f]{6}$", RegexOptions.CultureInvariant);
        private static readonly JsonSerializerSettings Settings = new JsonSerializerSettings {
            TypeNameHandling = TypeNameHandling.None, MissingMemberHandling = MissingMemberHandling.Error,
            Culture = CultureInfo.InvariantCulture, MaxDepth = 32, DateParseHandling = DateParseHandling.None,
            Formatting = Formatting.None, FloatParseHandling = FloatParseHandling.Double
        };

        private static void Require(bool valid, string code)
        { if (!valid) throw new InvalidDataException(code); }
        private static bool Finite(double value) { return !Double.IsNaN(value) && !Double.IsInfinity(value); }
        private static bool Text(string value)
        {
            if (String.IsNullOrWhiteSpace(value)) return false;
            // JSON Schema maxLength counts Unicode code points, not UTF-16 code units.
            int count = 0;
            for (int i = 0; i < value.Length; i++)
            {
                if (Char.IsHighSurrogate(value[i]) && i + 1 < value.Length && Char.IsLowSurrogate(value[i + 1])) i++;
                count++;
            }
            return count <= 1024;
        }

        public static void Validate(ChartDocument model)
        {
            Guid chartId;
            Require(model != null, "document.required");
            Require(model.SchemaVersion == CurrentVersion, "version.unsupported");
            Require(Guid.TryParseExact(model.ChartId, "D", out chartId) && chartId != Guid.Empty, "chart.id");
            Require(model.Revision >= 0, "revision.negative");
            Require(model.ChartType == "clustered-column", "chart.type.unsupported");
            var ids = new HashSet<string>(StringComparer.Ordinal);
            Action<string> unique = id => Require(id != null && IdPattern.IsMatch(id) && ids.Add(id), "semantic.id.invalid-or-duplicate");
            unique(model.ChartId);
            Require(model.Dataset != null && model.Dataset.Categories != null && model.Dataset.Categories.Count >= 1 && model.Dataset.Categories.Count <= 1000, "dataset.categories");
            unique(model.Dataset.Id);
            foreach (var category in model.Dataset.Categories)
            { Require(category != null && Text(category.Label), "category.label"); unique(category.Id); }
            Require(model.Series != null && model.Series.Count >= 1 && model.Series.Count <= 100, "series.count");
            foreach (var series in model.Series)
            {
                Require(series != null && Text(series.Name), "series.name"); unique(series.Id);
                Require(series.Values != null && series.Values.Count == model.Dataset.Categories.Count, "series.dimensions");
                Require(series.Values.All(v => !v.HasValue || Finite(v.Value)), "series.nonfinite");
            }
            Require(model.Axes != null && model.Axes.Count == 2 && model.Axes.All(a => a != null), "axes.count");
            Require(model.Axes.Select(a => a.Role).OrderBy(x => x).SequenceEqual(new[] { "category", "value" }), "axes.roles");
            foreach (var axis in model.Axes)
            {
                unique(axis.Id);
                Require(axis.Scale == (axis.Role == "category" ? "category" : "linear"), "axis.scale");
                Require(Text(axis.NumberFormat), "axis.format");
            }
            var layout = model.Layout;
            Require(layout != null && layout.Unit == "pt" && Finite(layout.X) && Finite(layout.Y) &&
                Finite(layout.Width) && Finite(layout.Height) && layout.Width > 0 && layout.Height > 0 &&
                Math.Abs(layout.X) <= 1000000 && Math.Abs(layout.Y) <= 1000000 && layout.Width <= 1000000 && layout.Height <= 1000000, "layout.invalid");
            Require(model.Style != null && model.Style.SeriesColors != null &&
                model.Style.SeriesColors.Count == model.Series.Count && model.Style.SeriesColors.All(c => c != null && ColorPattern.IsMatch(c)), "style.colors");
        }

        public static string Encode(ChartDocument model)
        {
            Validate(model);
            string json = JsonConvert.SerializeObject(model, Settings);
            Require(json.Length <= MaxCharacters, "document.oversized");
            return json;
        }
        public static ChartDocument Decode(string json)
        {
            Require(!String.IsNullOrWhiteSpace(json) && json.Length <= MaxCharacters, "document.missing-or-oversized");
            try
            {
                using (var reader = new StrictReader(new StringReader(json)))
                {
                    var token = JToken.Load(reader, new JsonLoadSettings { DuplicatePropertyNameHandling = DuplicatePropertyNameHandling.Error });
                    Require(!reader.Read(), "document.trailing-content");
                    CheckShape(token, typeof(ChartDocument));
                    var model = token.ToObject<ChartDocument>(JsonSerializer.Create(Settings));
                    Validate(model);
                    return model;
                }
            }
            catch (JsonException) { throw new InvalidDataException("document.invalid-json"); }
        }
        // Json.NET otherwise coerces numeric strings and ignores unknown fields. Check types
        // before deserialization so the C# boundary follows the published JSON contract.
        private static void CheckShape(JToken token, Type type)
        {
            var nullable = Nullable.GetUnderlyingType(type);
            if (nullable != null) { if (token.Type == JTokenType.Null) return; CheckShape(token, nullable); return; }
            if (type == typeof(string)) { Require(token.Type == JTokenType.String, "json.type.string"); return; }
            if (type == typeof(bool)) { Require(token.Type == JTokenType.Boolean, "json.type.boolean"); return; }
            if (type == typeof(int)) {
                Require(token.Type == JTokenType.Integer || token.Type == JTokenType.Float, "json.type.integer");
                double integerValue = (double)token;
                Require(Finite(integerValue) && integerValue == Math.Truncate(integerValue) && integerValue >= Int32.MinValue && integerValue <= Int32.MaxValue, "json.type.integer");
                return;
            }
            if (type == typeof(double)) { Require(token.Type == JTokenType.Integer || token.Type == JTokenType.Float, "json.type.number"); return; }
            if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(List<>))
            {
                Require(token.Type == JTokenType.Array, "json.type.array");
                foreach (var item in token) CheckShape(item, type.GetGenericArguments()[0]);
                return;
            }
            Require(token.Type == JTokenType.Object, "json.type.object");
            var properties = type.GetProperties().Select(p => new { Property = p, Json = p.GetCustomAttribute<JsonPropertyAttribute>() }).ToArray();
            var value = (JObject)token;
            Require(value.Count == properties.Length, "json.members");
            foreach (var property in properties)
            {
                JToken child;
                Require(value.TryGetValue(property.Json.PropertyName, StringComparison.Ordinal, out child), "json.member.required");
                CheckShape(child, property.Property.PropertyType);
            }
        }
        private sealed class StrictReader : JsonTextReader
        {
            public StrictReader(TextReader input) : base(input) { MaxDepth = 32; DateParseHandling = DateParseHandling.None; }
            public override bool Read()
            {
                var result = base.Read();
                if (TokenType == JsonToken.Comment || TokenType == JsonToken.Undefined || TokenType == JsonToken.StartConstructor)
                    throw new JsonReaderException("Non-JSON token.");
                return result;
            }
        }
    }
}
