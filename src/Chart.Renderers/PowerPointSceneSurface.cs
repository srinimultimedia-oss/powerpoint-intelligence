using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointIntelligence.Rendering
{
    // Bounded local feasibility adapter; the ribbon exposes only a fixed sample.
    // General production callers still require the command/revision/approval boundary.
    public sealed class PowerPointSceneSurface : IEditableSceneSurface
    {
        public const string ChartTag = "PPI_RENDER_CHART";
        public const string NodeTag = "PPI_RENDER_NODE";
        public const string OperationTag = "PPI_RENDER_OP";
        private readonly PowerPoint.Presentation presentation;
        private readonly PowerPoint.Slide slide;
        private readonly int threadId;
        private readonly Dictionary<int, PowerPoint.Shape> created = new Dictionary<int, PowerPoint.Shape>();
        private string operationId, chartId;
        private bool active, uncertain;
        private Office.CustomXMLPart part;
        public PowerPointSceneSurface(PowerPoint.Presentation presentation, PowerPoint.Slide slide)
        { this.presentation = presentation ?? throw new ArgumentNullException("presentation"); this.slide = slide ?? throw new ArgumentNullException("slide"); threadId = Thread.CurrentThread.ManagedThreadId; VerifyAccess(); }
        public void VerifyAccess()
        {
            if (Thread.CurrentThread.ManagedThreadId != threadId || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA) throw new InvalidOperationException("render.requires-owner-sta");
        }
        public void Begin(string operation, string chart)
        {
            VerifyAccess(); if (active) throw new InvalidOperationException("render.surface-busy");
            if (FindChartObjects(chart).Any()) throw new InvalidDataException("render.duplicate-chart-id");
            // Read-only preflight. A failing Begin must never mutate the document.
            if (((PowerPoint.Presentation)slide.Parent).FullName != presentation.FullName) throw new InvalidOperationException("render.slide-scope");
            operationId = operation; chartId = chart; created.Clear(); part = null; uncertain = false; active = true;
        }
        private void MutationAccess() { VerifyAccess(); if (!active) throw new InvalidOperationException("render.no-operation"); }
        private int Create(Func<PowerPoint.Shape> factory, string token)
        {
            MutationAccess(); PowerPoint.Shape shape = null; bool registered = false;
            try
            {
                shape = factory(); created.Add(shape.Id, shape); registered = true;
                shape.Name = "ppi_" + operationId + "_" + created.Count;
                shape.Tags.Add(OperationTag, operationId); shape.Tags.Add(ChartTag, chartId);
                if (token != null) { shape.Tags.Add(NodeTag, token); shape.Tags.Add("PPI_RENDER_ROLE", "NODE"); }
                return shape.Id;
            }
            catch { if (!registered) uncertain = true; throw; }
        }
        public int CreateAnchor(ScenePlot plot, bool trailing)
        {
            int id = Create(() => slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, (float)(plot.X + (trailing ? plot.Width - 1 : 0)), (float)(plot.Y + (trailing ? plot.Height - 1 : 0)), 1, 1), null);
            var shape = created[id]; shape.Fill.Visible = Office.MsoTriState.msoFalse; shape.Line.Visible = Office.MsoTriState.msoFalse;
            shape.Tags.Add("PPI_RENDER_ROLE", "ANCHOR"); shape.AlternativeText = "Chart boundary anchor"; return id;
        }
        public int CreateNode(SceneNode node, string token)
        {
            MutationAccess();
            // The first canonical insertion profile uses rectangles only. Keep the
            // lower-level geometry adapter closed and reject unsupported types.
            var box = node.Geometry as BoxGeometry;
            if (box == null || box.Type != "rectangle") throw new NotSupportedException("render.profile-rectangle-only");
            int id = Create(() => slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, (float)box.X, (float)box.Y, (float)box.Width, (float)box.Height), token);
            Paint(created[id], node.Paint); created[id].AlternativeText = "Editable chart bar"; return id;
        }
        private static int Rgb(string color)
        { int rgb = Int32.Parse(color.Substring(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture); return ((rgb >> 16) & 255) | (rgb & 0xff00) | ((rgb & 255) << 16); }
        private static void Paint(PowerPoint.Shape shape, ScenePaint paint)
        {
            shape.Fill.Visible = paint.Fill == null ? Office.MsoTriState.msoFalse : Office.MsoTriState.msoTrue;
            if (paint.Fill != null) { shape.Fill.Solid(); shape.Fill.ForeColor.RGB = Rgb(paint.Fill); shape.Fill.Transparency = (float)(1 - paint.Opacity); }
            shape.Line.Visible = paint.Stroke == null || paint.StrokeWidth == 0 ? Office.MsoTriState.msoFalse : Office.MsoTriState.msoTrue;
            if (paint.Stroke != null && paint.StrokeWidth > 0) { shape.Line.ForeColor.RGB = Rgb(paint.Stroke); shape.Line.Weight = (float)paint.StrokeWidth; shape.Line.Transparency = (float)(1 - paint.Opacity); }
        }
        public int Group(IReadOnlyList<int> children)
        {
            MutationAccess();
            if (children.Count != created.Count || children.Distinct().Count() != children.Count || children.Any(id => !created.ContainsKey(id))) throw new InvalidDataException("render.group-ownership");
            var names = children.Select(id => created[id].Name).ToArray();
            // Shapes.Range integer arguments are collection indices, not shape IDs.
            // Unique names identify only the objects made by this operation.
            var root = slide.Shapes.Range(names).Group();
            root.Name = "ppi_chart_" + chartId; root.Tags.Add(OperationTag, operationId); root.Tags.Add(ChartTag, chartId);
            root.Tags.Add("PPI_RENDER_ROLE", "ROOT");
            root.AlternativeText = "Editable clustered-column chart. Canonical data is embedded; axes and legend are not rendered in this feasibility profile.";
            return root.Id;
        }
        private PowerPoint.Shape TopLevel(int id)
        { for (int i = 1; i <= slide.Shapes.Count; i++) if (slide.Shapes[i].Id == id) return slide.Shapes[i]; throw new InvalidDataException("render.root-missing"); }
        public void Persist(int rootId, string chart, string xml)
        {
            MutationAccess(); var root = TopLevel(rootId);
            if (!OnlyOwned(root)) throw new InvalidDataException("render.root-ownership");
            RenderedChartMetadata.Decode(xml);
            part = root.CustomerData.Add();
            if (!part.LoadXML(xml)) throw new InvalidDataException("render.metadata-rejected");
            var read = ReadVerified(root);
            if (Guid.Parse(chart) != Guid.Parse(ClusteredProjection.Validate(read.CanonicalJson, read.SceneJson).Scene.ChartId)) throw new InvalidDataException("render.metadata-identity");
        }
        public void Complete() { MutationAccess(); active = false; created.Clear(); part = null; }
        private bool OnlyOwned(PowerPoint.Shape shape)
        {
            if (shape.Type != Office.MsoShapeType.msoGroup) return created.ContainsKey(shape.Id);
            for (int i = 1; i <= shape.GroupItems.Count; i++) if (!OnlyOwned(shape.GroupItems[i])) return false;
            return shape.GroupItems.Count > 0;
        }
        private bool ContainsOwned(PowerPoint.Shape shape)
        {
            if (shape.Type != Office.MsoShapeType.msoGroup) return created.ContainsKey(shape.Id);
            for (int i = 1; i <= shape.GroupItems.Count; i++) if (ContainsOwned(shape.GroupItems[i])) return true; return false;
        }
        private void RemoveOwned(PowerPoint.Shape shape)
        {
            if (OnlyOwned(shape)) { shape.Delete(); return; }
            if (shape.Type == Office.MsoShapeType.msoGroup)
                for (int i = shape.GroupItems.Count; i >= 1; i--) if (ContainsOwned(shape.GroupItems[i])) RemoveOwned(shape.GroupItems[i]);
        }
        public bool Rollback()
        {
            VerifyAccess(); bool complete = !uncertain;
            try { if (part != null) part.Delete(); } catch { complete = false; }
            try { for (int i = slide.Shapes.Count; i >= 1; i--) if (ContainsOwned(slide.Shapes[i])) RemoveOwned(slide.Shapes[i]); }
            catch { complete = false; }
            try { for (int i = 1; i <= slide.Shapes.Count; i++) if (ContainsOwned(slide.Shapes[i])) complete = false; }
            catch { complete = false; }
            active = false; created.Clear(); part = null; return complete;
        }
        private static IEnumerable<PowerPoint.Shape> Descendants(PowerPoint.Shape shape)
        {
            yield return shape;
            if (shape.Type == Office.MsoShapeType.msoGroup) for (int i = 1; i <= shape.GroupItems.Count; i++) foreach (var child in Descendants(shape.GroupItems[i])) yield return child;
        }
        private IEnumerable<PowerPoint.Shape> FindChartObjects(string chart)
        {
            for (int s = 1; s <= presentation.Slides.Count; s++) for (int i = 1; i <= presentation.Slides[s].Shapes.Count; i++)
                foreach (var shape in Descendants(presentation.Slides[s].Shapes[i])) if (String.Equals(shape.Tags[ChartTag], chart, StringComparison.OrdinalIgnoreCase)) yield return shape;
        }
        private static IEnumerable<PowerPoint.Shape> Leaves(PowerPoint.Shape root)
        {
            if (root.Type != Office.MsoShapeType.msoGroup) { yield return root; yield break; }
            for (int i = 1; i <= root.GroupItems.Count; i++) foreach (var child in Leaves(root.GroupItems[i])) yield return child;
        }
        public RenderedChartRecord ReadVerified(PowerPoint.Shape root)
        {
            VerifyAccess(); if (root.Type != Office.MsoShapeType.msoGroup) throw new InvalidDataException("render.root-not-group");
            Office.CustomXMLPart found = null;
            foreach (Office.CustomXMLPart candidate in root.CustomerData)
                if (candidate.NamespaceURI == RenderedChartMetadata.Namespace) { if (found != null) throw new InvalidDataException("render.metadata-ambiguous"); found = candidate; }
            if (found == null) throw new InvalidDataException("render.metadata-missing");
            var record = RenderedChartMetadata.Decode(found.XML); var projection = RenderedChartMetadata.Validate(record);
            if (!String.Equals(root.Tags[ChartTag], projection.Scene.ChartId, StringComparison.OrdinalIgnoreCase) || FindChartObjects(projection.Scene.ChartId).Count(s => s.Tags["PPI_RENDER_ROLE"] == "ROOT") != 1) throw new InvalidDataException("render.duplicate-or-wrong-identity");
            var byToken = new Dictionary<string, PowerPoint.Shape>(StringComparer.Ordinal);
            foreach (var shape in Leaves(root))
            {
                string token = shape.Tags[NodeTag]; if (String.IsNullOrEmpty(token)) continue;
                if (byToken.ContainsKey(token)) throw new InvalidDataException("render.duplicate-node-token"); byToken.Add(token, shape);
            }
            foreach (var binding in record.Bindings)
            {
                var node = projection.Scene.Nodes.Single(n => n.Id == binding.NodeId); PowerPoint.Shape shape;
                if (!ClusteredProjection.HasArea(node)) { if (byToken.ContainsKey(binding.Token)) throw new InvalidDataException("render.unexpected-empty-node"); continue; }
                if (!byToken.TryGetValue(binding.Token, out shape)) throw new InvalidDataException("render.node-missing");
                var geometry = (BoxGeometry)node.Geometry;
                // Proposed spike tolerance for PowerPoint's Single-precision position API.
                if (shape.AutoShapeType != Office.MsoAutoShapeType.msoShapeRectangle || Math.Abs(shape.Left - geometry.X) > .05 || Math.Abs(shape.Top - geometry.Y) > .05 || Math.Abs(shape.Width - geometry.Width) > .05 || Math.Abs(shape.Height - geometry.Height) > .05) throw new InvalidDataException("render.native-geometry-drift");
                if (shape.Fill.Visible != Office.MsoTriState.msoTrue || shape.Fill.ForeColor.RGB != Rgb(node.Paint.Fill) || Math.Abs(shape.Fill.Transparency - (1 - node.Paint.Opacity)) > .001 || shape.Line.Visible != Office.MsoTriState.msoFalse) throw new InvalidDataException("render.native-style-drift");
                binding.ShapeId = shape.Id; byToken.Remove(binding.Token);
            }
            if (byToken.Count != 0) throw new InvalidDataException("render.unrecognized-node-token"); return record;
        }
    }
}
