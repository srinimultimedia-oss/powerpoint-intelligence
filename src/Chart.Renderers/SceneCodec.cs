using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace PowerPointIntelligence.Rendering
{
    public static class SceneCodec
    {
        public const int CurrentVersion = 1;
        public const int MaxCharacters = 8388608;
        private static readonly Regex GuidPattern = new Regex(@"\A(?!00000000-0000-0000-0000-000000000000\z)[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\z", RegexOptions.CultureInvariant);
        private static readonly Regex LocalIdPattern = new Regex(@"\A[A-Za-z0-9][A-Za-z0-9:._-]{0,159}(?:/[A-Za-z0-9][A-Za-z0-9:._-]{0,159}){0,3}\z", RegexOptions.CultureInvariant);
        private static readonly Regex ColorPattern = new Regex(@"\A#[0-9a-fA-F]{6}\z", RegexOptions.CultureInvariant);
        private static readonly Regex CodePattern = new Regex(@"\A[A-Z][A-Z0-9_]{0,63}\z", RegexOptions.CultureInvariant);
        private static readonly Regex HashPattern = new Regex(@"\A[0-9a-f]{64}\z", RegexOptions.CultureInvariant);
        private static JsonSerializer Serializer()
        {
            var serializer = JsonSerializer.Create(new JsonSerializerSettings {
                TypeNameHandling = TypeNameHandling.None, MissingMemberHandling = MissingMemberHandling.Error,
                Culture = CultureInfo.InvariantCulture, MaxDepth = 32, DateParseHandling = DateParseHandling.None,
                FloatParseHandling = FloatParseHandling.Double, Formatting = Formatting.None
            });
            serializer.Converters.Add(new GeometryReader()); return serializer;
        }
        private static void Require(bool valid, string code) { if (!valid) throw new InvalidDataException(code); }
        private static JObject Fields(JToken value, params string[] keys)
        {
            var obj = value as JObject;
            Require(obj != null && obj.Count == keys.Length && keys.All(k => obj.Property(k, StringComparison.Ordinal) != null), "scene.fields");
            return obj;
        }
        private static string Text(JToken value)
        { Require(value != null && value.Type == JTokenType.String, "scene.string"); return (string)value; }
        private static double Number(JToken value, double minimum = -10000000, double maximum = 10000000)
        {
            Require(value != null && (value.Type == JTokenType.Integer || value.Type == JTokenType.Float), "scene.number");
            double number = (double)value;
            Require(!Double.IsNaN(number) && !Double.IsInfinity(number) && number >= minimum && number <= maximum, "scene.number");
            return number;
        }
        private static JArray Items(JToken value, int maximum, int minimum = 0)
        { var array = value as JArray; Require(array != null && array.Count >= minimum && array.Count <= maximum, "scene.array"); return array; }
        private static void Validate(JObject root, bool checkHash)
        {
            Fields(root, "schemaVersion", "chartId", "modelRevision", "unit", "plot", "semanticScene", "diagnostics", "renderHash");
            Require(Number(root["schemaVersion"]) == CurrentVersion && Text(root["unit"]) == "pt", "scene.version-or-unit");
            string chartId = Text(root["chartId"]); Require(GuidPattern.IsMatch(chartId), "scene.chart-id");
            double revision = Number(root["modelRevision"], 0, Int32.MaxValue); Require(revision == Math.Truncate(revision), "scene.revision");
            var plot = Fields(root["plot"], "x", "y", "width", "height", "unit");
            Number(plot["x"], -1000000, 1000000); Number(plot["y"], -1000000, 1000000);
            Require(Number(plot["width"], 0, 1000000) > 0 && Number(plot["height"], 0, 1000000) > 0 && Text(plot["unit"]) == "pt", "scene.plot");
            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (var value in Items(root["semanticScene"], 200000))
            {
                var node = Fields(value, "id", "geometry", "paint"); string id = Text(node["id"]);
                Require(id.StartsWith(chartId + "/", StringComparison.Ordinal) && LocalIdPattern.IsMatch(id.Substring(37)) && ids.Add(id), "scene.node-id");
                var geometry = node["geometry"] as JObject; Require(geometry != null, "scene.geometry");
                string type = Text(geometry["type"]);
                switch (type)
                {
                    case "rectangle": case "ellipse":
                        Fields(geometry, "type", "x", "y", "width", "height"); Box(geometry); break;
                    case "text":
                        Fields(geometry, "type", "x", "y", "width", "height", "text", "fontFamily", "fontSize", "alignment"); Box(geometry);
                        Require(SceneText.Valid(Text(geometry["text"]), 4096), "scene.text");
                        string font = Text(geometry["fontFamily"]);
                        Require(SceneText.Valid(font, 128) && SceneText.NonWhitespace(font), "scene.font");
                        Number(geometry["fontSize"], 1, 1000);
                        Require(new[] { "left", "center", "right" }.Contains(Text(geometry["alignment"]), StringComparer.Ordinal), "scene.alignment"); break;
                    case "polyline": case "polygon":
                        Fields(geometry, "type", "points");
                        foreach (var point in Items(geometry["points"], 10000, type == "polygon" ? 3 : 2))
                        { var p = Fields(point, "x", "y"); Number(p["x"]); Number(p["y"]); } break;
                    case "sector":
                        Fields(geometry, "type", "cx", "cy", "radius", "startAngle", "sweep");
                        Number(geometry["cx"]); Number(geometry["cy"]); Number(geometry["radius"], 0);
                        Number(geometry["startAngle"], -360, 360); Require(Number(geometry["sweep"], 0, 360) > 0, "scene.sweep"); break;
                    default: throw new InvalidDataException("scene.geometry-type");
                }
                var paint = Fields(node["paint"], "fill", "stroke", "strokeWidth", "opacity");
                foreach (string name in new[] { "fill", "stroke" })
                    Require(paint[name].Type == JTokenType.Null || ColorPattern.IsMatch(Text(paint[name])), "scene.color");
                Number(paint["strokeWidth"], 0, 1000); Number(paint["opacity"], 0, 1);
            }
            foreach (var value in Items(root["diagnostics"], 10000))
            {
                var d = Fields(value, "code", "nodeId"); Require(CodePattern.IsMatch(Text(d["code"])), "scene.diagnostic-code");
                Require(d["nodeId"].Type == JTokenType.Null || ids.Contains(Text(d["nodeId"])), "scene.diagnostic-node");
            }
            if (checkHash)
            {
                string hash = Text(root["renderHash"]); Require(HashPattern.IsMatch(hash), "scene.hash-format");
                Require(String.Equals(hash, Hash(root), StringComparison.Ordinal), "scene.hash-mismatch");
            }
        }
        private static void Box(JObject value)
        { Number(value["x"]); Number(value["y"]); Number(value["width"], 0); Number(value["height"], 0); }

        public static RenderResult Decode(string json)
        {
            var root = StrictSceneJson.Parse(json) as JObject; Require(root != null, "scene.fields");
            Validate(root, true);
            try { return root.ToObject<RenderResult>(Serializer()); }
            catch (JsonException) { throw new InvalidDataException("scene.dto"); }
        }
        public static void ValidateJsonSyntax(string json) { StrictSceneJson.Parse(json); }
        private static JObject Token(RenderResult scene)
        {
            Require(scene != null, "scene.required");
            try { return JObject.FromObject(scene, Serializer()); }
            catch (JsonException) { throw new InvalidDataException("scene.dto"); }
        }
        public static string Encode(RenderResult scene)
        {
            var root = Token(scene); Validate(root, true); string json = root.ToString(Formatting.None);
            Require(json.Length <= MaxCharacters, "scene.json-size"); return json;
        }
        public static string ComputeHash(RenderResult scene)
        {
            var root = Token(scene); Validate(root, false);
            Require(root.ToString(Formatting.None).Length <= MaxCharacters, "scene.json-size"); return Hash(root);
        }
        private static string Hash(JObject root)
        {
            using (var stream = new MemoryStream())
            {
                WriteString(stream, "PPI-SCENE-1");
                // Do not mutate the source object or include its renderHash field.
                WriteObject(stream, root.Properties().Where(p => p.Name != "renderHash"));
                using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(stream.ToArray())).Replace("-", "").ToLowerInvariant();
            }
        }
        private static void Length(Stream stream, int length)
        { stream.WriteByte((byte)(length >> 24)); stream.WriteByte((byte)(length >> 16)); stream.WriteByte((byte)(length >> 8)); stream.WriteByte((byte)length); }
        private static void WriteString(Stream stream, string value)
        { byte[] bytes = new UTF8Encoding(false, true).GetBytes(value); Length(stream, bytes.Length); stream.Write(bytes, 0, bytes.Length); }
        private static void WriteObject(Stream stream, IEnumerable<JProperty> properties)
        {
            var sorted = properties.OrderBy(p => p.Name, StringComparer.Ordinal).ToArray(); stream.WriteByte(6); Length(stream, sorted.Length);
            foreach (var p in sorted) { stream.WriteByte(4); WriteString(stream, p.Name); Write(stream, p.Value); }
        }
        private static void Write(Stream stream, JToken token)
        {
            switch (token.Type)
            {
                case JTokenType.Null: stream.WriteByte(0); break;
                case JTokenType.Boolean: stream.WriteByte((bool)token ? (byte)2 : (byte)1); break;
                case JTokenType.Integer: case JTokenType.Float:
                    stream.WriteByte(3); double number = (double)token; byte[] bytes = BitConverter.GetBytes(number == 0 ? 0d : number);
                    if (BitConverter.IsLittleEndian) Array.Reverse(bytes); stream.Write(bytes, 0, bytes.Length); break;
                case JTokenType.String: stream.WriteByte(4); WriteString(stream, (string)token); break;
                case JTokenType.Array: stream.WriteByte(5); Length(stream, ((JArray)token).Count); foreach (var item in token) Write(stream, item); break;
                case JTokenType.Object: WriteObject(stream, ((JObject)token).Properties()); break;
                default: throw new InvalidDataException("scene.hash-value");
            }
        }
        private sealed class GeometryReader : JsonConverter
        {
            public override bool CanWrite { get { return false; } }
            public override bool CanConvert(Type type) { return type == typeof(SceneGeometry); }
            public override object ReadJson(JsonReader reader, Type type, object existingValue, JsonSerializer serializer)
            {
                var value = JObject.Load(reader); Type concrete;
                switch ((string)value["type"])
                {
                    case "rectangle": case "ellipse": concrete = typeof(BoxGeometry); break;
                    case "polyline": case "polygon": concrete = typeof(PathGeometry); break;
                    case "sector": concrete = typeof(SectorGeometry); break;
                    case "text": concrete = typeof(TextGeometry); break;
                    default: throw new InvalidDataException("scene.geometry-type");
                }
                return value.ToObject(concrete, serializer);
            }
            public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer) { throw new NotSupportedException(); }
        }
    }
}
