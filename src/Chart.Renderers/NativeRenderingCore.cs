using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using PowerPointIntelligence.Domain;

namespace PowerPointIntelligence.Rendering
{
    public sealed class ValidatedProjection
    {
        public ChartDocument Model { get; internal set; }
        public RenderResult Scene { get; internal set; }
    }
    public static class ClusteredProjection
    {
        public const int MaximumNativeNodes = 2000; // Feasibility guard, not a production sizing decision.
        private static void Require(bool condition, string code) { if (!condition) throw new InvalidDataException(code); }
        private static bool Near(double a, double b) { return Math.Abs(a - b) <= 1e-9 * Math.Max(1, Math.Max(Math.Abs(a), Math.Abs(b))); }
        public static ValidatedProjection Validate(string canonicalJson, string sceneJson)
        {
            Require(!String.IsNullOrWhiteSpace(canonicalJson) && canonicalJson.Length <= ChartCodec.MaxCharacters, "projection.canonical-size");
            SceneCodec.ValidateJsonSyntax(canonicalJson);
            var model = ChartCodec.Decode(canonicalJson); var scene = SceneCodec.Decode(sceneJson);
            Require(scene.ChartId == Guid.Parse(model.ChartId).ToString("D") && scene.ModelRevision == model.Revision, "projection.identity-or-revision");
            Require(scene.Nodes.Count <= MaximumNativeNodes, "projection.native-size-limit");
            var a = model.Layout; var b = scene.Plot;
            Require(a.X == b.X && a.Y == b.Y && a.Width == b.Width && a.Height == b.Height, "projection.layout");
            Require(a.Width >= 2 && a.Height >= 2, "projection.native-minimum-size");
            double minimum = 0, maximum = 0;
            foreach (var series in model.Series) foreach (var value in series.Values)
                if (value.HasValue) { minimum = Math.Min(minimum, value.Value); maximum = Math.Max(maximum, value.Value); }
            if (minimum == maximum) { minimum = -1; maximum = 1; }
            double range = maximum - minimum;
            Require(!Double.IsInfinity(range), "projection.range-overflow");
            double span = a.Width / model.Dataset.Categories.Count, slot = span * .8 / model.Series.Count;
            Func<double, double> y = value => a.Y + (1 - (value - minimum) / range) * a.Height;
            int index = 0;
            for (int ci = 0; ci < model.Dataset.Categories.Count; ci++) for (int si = 0; si < model.Series.Count; si++)
            {
                var series = model.Series[si]; var value = series.Values[ci]; if (!value.HasValue) continue;
                Require(index < scene.Nodes.Count, "projection.missing-node"); var node = scene.Nodes[index++];
                var geometry = node.Geometry as BoxGeometry;
                Require(node.Id == scene.ChartId + "/bar:" + series.Id + "/" + model.Dataset.Categories[ci].Id && geometry != null && geometry.Type == "rectangle", "projection.node-identity");
                Require(Near(geometry.X, a.X + ci * span + span * .1 + si * slot + slot * .05) && Near(geometry.Width, slot * .9) && Near(geometry.Y, Math.Min(y(0), y(value.Value))) && Near(geometry.Height, Math.Abs(y(value.Value) - y(0))), "projection.numeric-mismatch");
                Require(String.Equals(node.Paint.Fill, model.Style.SeriesColors[si], StringComparison.OrdinalIgnoreCase) && node.Paint.Stroke == null && node.Paint.StrokeWidth == 0 && node.Paint.Opacity == 1, "projection.style-mismatch");
            }
            Require(index == scene.Nodes.Count, "projection.extra-node");
            var codes = model.Style.ShowLegend ? new[] { "AXES_NOT_RENDERED", "LEGEND_NOT_RENDERED" } : new[] { "AXES_NOT_RENDERED" };
            Require(scene.Diagnostics.Select(d => d.Code).SequenceEqual(codes) && scene.Diagnostics.All(d => d.NodeId == null), "projection.diagnostics");
            return new ValidatedProjection { Model = model, Scene = scene };
        }
        public static bool HasArea(SceneNode node)
        { var box = node.Geometry as BoxGeometry; return box == null || box.Width > 0 && box.Height > 0; }
    }
    public sealed class NativeBinding
    {
        public string NodeId { get; set; }
        public string Token { get; set; }
        public int? ShapeId { get; set; }
    }
    public sealed class RenderedChartRecord
    {
        public string CanonicalJson { get; set; }
        public string SceneJson { get; set; }
        public List<NativeBinding> Bindings { get; set; }
    }
    public static class RenderedChartMetadata
    {
        public const string Namespace = "urn:powerpoint-intelligence:rendered-chart:v1";
        public const int MaximumCharacters = 12582912;
        public static string NodeToken(string id)
        { using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(id))).Replace("-", ""); }
        private static void Require(bool valid) { if (!valid) throw new InvalidDataException("rendered-metadata.invalid"); }
        public static ValidatedProjection Validate(RenderedChartRecord record)
        {
            Require(record != null); var projection = ClusteredProjection.Validate(record.CanonicalJson, record.SceneJson);
            Require(record.Bindings != null && record.Bindings.Count == projection.Scene.Nodes.Count);
            var shapeIds = new HashSet<int>();
            for (int i = 0; i < record.Bindings.Count; i++)
            {
                var binding = record.Bindings[i]; var node = projection.Scene.Nodes[i];
                Require(binding != null && binding.NodeId == node.Id && binding.Token == NodeToken(node.Id));
                Require(ClusteredProjection.HasArea(node) ? binding.ShapeId.HasValue && binding.ShapeId.Value > 0 && shapeIds.Add(binding.ShapeId.Value) : !binding.ShapeId.HasValue);
            }
            return projection;
        }
        public static string Encode(RenderedChartRecord record)
        {
            var projection = Validate(record); var text = new StringBuilder();
            using (var xml = XmlWriter.Create(text, new XmlWriterSettings { OmitXmlDeclaration = true }))
            {
                xml.WriteStartElement("RenderedChart", Namespace); xml.WriteAttributeString("version", "1"); xml.WriteAttributeString("chartId", projection.Scene.ChartId);
                xml.WriteElementString("Canonical", Namespace, ChartCodec.Encode(projection.Model));
                xml.WriteElementString("Scene", Namespace, SceneCodec.Encode(projection.Scene));
                xml.WriteStartElement("Bindings", Namespace);
                foreach (var binding in record.Bindings)
                {
                    xml.WriteStartElement("Binding", Namespace); xml.WriteAttributeString("nodeId", binding.NodeId); xml.WriteAttributeString("token", binding.Token);
                    xml.WriteAttributeString("shapeId", binding.ShapeId.HasValue ? binding.ShapeId.Value.ToString(CultureInfo.InvariantCulture) : ""); xml.WriteEndElement();
                }
                xml.WriteEndElement(); xml.WriteEndElement();
            }
            Require(text.Length <= MaximumCharacters); return text.ToString();
        }
        public static RenderedChartRecord Decode(string text)
        {
            Require(!String.IsNullOrWhiteSpace(text) && text.Length <= MaximumCharacters);
            try
            {
                var doc = new XmlDocument { XmlResolver = null };
                using (var reader = XmlReader.Create(new StringReader(text), new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = MaximumCharacters, IgnoreWhitespace = true })) doc.Load(reader);
                var root = doc.DocumentElement; Require(root != null && root.LocalName == "RenderedChart" && root.NamespaceURI == Namespace && root.GetAttribute("version") == "1");
                Require(root.Attributes.Cast<XmlAttribute>().All(a => a.Name == "version" || a.Name == "chartId" || a.Name == "xmlns"));
                var children = root.ChildNodes.Cast<XmlNode>().ToArray();
                Require(children.Length == 3 && children[0].LocalName == "Canonical" && children[1].LocalName == "Scene" && children[2].LocalName == "Bindings" && children.All(c => c.NodeType == XmlNodeType.Element && c.NamespaceURI == Namespace && c.Attributes.Count == 0));
                Require(children.Take(2).All(c => c.ChildNodes.Cast<XmlNode>().All(t => t.NodeType == XmlNodeType.Text || t.NodeType == XmlNodeType.CDATA)));
                var record = new RenderedChartRecord { CanonicalJson = children[0].InnerText, SceneJson = children[1].InnerText, Bindings = new List<NativeBinding>() };
                foreach (XmlNode item in children[2].ChildNodes)
                {
                    Require(item.NodeType == XmlNodeType.Element && item.LocalName == "Binding" && item.NamespaceURI == Namespace && item.Attributes.Count == 3 && item.ChildNodes.Count == 0);
                    Require(item.Attributes.Cast<XmlAttribute>().Select(a => a.Name).OrderBy(n => n).SequenceEqual(new[] { "nodeId", "shapeId", "token" }));
                    string shape = item.Attributes["shapeId"].Value; int shapeId;
                    Require(shape.Length == 0 || Int32.TryParse(shape, NumberStyles.None, CultureInfo.InvariantCulture, out shapeId) && shapeId > 0);
                    record.Bindings.Add(new NativeBinding { NodeId = item.Attributes["nodeId"].Value, Token = item.Attributes["token"].Value, ShapeId = shape.Length == 0 ? (int?)null : Int32.Parse(shape, CultureInfo.InvariantCulture) });
                }
                var projection = Validate(record); Require(root.GetAttribute("chartId") == projection.Scene.ChartId); return record;
            }
            catch (XmlException) { throw new InvalidDataException("rendered-metadata.invalid-xml"); }
        }
    }
    // The surface owns COM RCWs. The coordinator passes only validated DTOs and IDs.
    public interface IEditableSceneSurface
    {
        void VerifyAccess();
        void Begin(string operationId, string chartId);
        int CreateAnchor(ScenePlot plot, bool trailing);
        int CreateNode(SceneNode node, string token);
        int Group(IReadOnlyList<int> children);
        void Persist(int rootId, string chartId, string xml);
        bool Rollback();
        void Complete();
    }
    public sealed class NativeRenderOutcome
    {
        public int RootShapeId { get; internal set; }
        public RenderedChartRecord Record { get; internal set; }
    }
    public sealed class NativeRenderException : Exception
    {
        public bool RollbackComplete { get; private set; }
        public string Operation { get; private set; }
        internal NativeRenderException(string operation, bool rollbackComplete, Exception inner) : base("Native rendering failed at " + operation + "; rollback=" + rollbackComplete, inner) { Operation = operation; RollbackComplete = rollbackComplete; }
    }
    public sealed class NativeSceneRenderer
    {
        private readonly IEditableSceneSurface surface;
        private bool busy;
        public NativeSceneRenderer(IEditableSceneSurface surface) { this.surface = surface ?? throw new ArgumentNullException("surface"); }
        public NativeRenderOutcome Insert(string canonicalJson, string sceneJson)
        {
            surface.VerifyAccess(); if (busy) throw new InvalidOperationException("render.reentrant");
            var projection = ClusteredProjection.Validate(canonicalJson, sceneJson); // All rejection before the first mutation.
            busy = true; string operation = "begin"; bool begun = false;
            try
            {
                surface.Begin(Guid.NewGuid().ToString("D"), projection.Scene.ChartId); begun = true;
                operation = "anchors"; var children = new List<int> { surface.CreateAnchor(projection.Scene.Plot, false), surface.CreateAnchor(projection.Scene.Plot, true) };
                var bindings = new List<NativeBinding>(); operation = "nodes";
                foreach (var node in projection.Scene.Nodes)
                {
                    string token = RenderedChartMetadata.NodeToken(node.Id);
                    int? id = ClusteredProjection.HasArea(node) ? (int?)surface.CreateNode(node, token) : null;
                    if (id.HasValue) children.Add(id.Value);
                    bindings.Add(new NativeBinding { NodeId = node.Id, Token = token, ShapeId = id });
                }
                operation = "group"; int rootId = surface.Group(children);
                var record = new RenderedChartRecord { CanonicalJson = ChartCodec.Encode(projection.Model), SceneJson = SceneCodec.Encode(projection.Scene), Bindings = bindings };
                operation = "metadata"; surface.Persist(rootId, projection.Scene.ChartId, RenderedChartMetadata.Encode(record));
                operation = "complete"; surface.Complete(); return new NativeRenderOutcome { RootShapeId = rootId, Record = record };
            }
            catch (Exception ex)
            {
                bool rolledBack = !begun; if (begun) try { rolledBack = surface.Rollback(); } catch { }
                throw new NativeRenderException(operation, rolledBack, ex);
            }
            finally { busy = false; }
        }
    }
}
