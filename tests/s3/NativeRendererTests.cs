using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml;
using PowerPointIntelligence.Domain;
using PowerPointIntelligence.Rendering;
using PowerPointIntelligence.AddIn;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

internal static class NativeRendererTests
{
    private static int passed;
    private static string canonical, scene;
    private static void Require(bool valid, string name) { if (!valid) throw new Exception(name); }
    private static void Test(string name, Action action) { action(); Console.WriteLine("PASS " + name); passed++; }
    private static void Reject(Action action) { try { action(); } catch (InvalidDataException) { return; } throw new Exception("Expected InvalidDataException."); }
    private static string ChangedScene(Action<RenderResult> change)
    { var value = SceneCodec.Decode(scene); change(value); value.RenderHash = SceneCodec.ComputeHash(value); return SceneCodec.Encode(value); }
    private sealed class FakeSurface : IEditableSceneSurface
    {
        private readonly int threadId = Thread.CurrentThread.ManagedThreadId;
        public readonly Dictionary<int, string> Objects = new Dictionary<int, string> { { 99, "user" } };
        public int FailAt, Steps, Begins, Next = 100;
        public bool After, FailRollback, ThrowRollback;
        public string Xml;
        public Action OnBegin;
        public void VerifyAccess() { if (Thread.CurrentThread.ManagedThreadId != threadId || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA) throw new InvalidOperationException("wrong thread"); }
        public void Begin(string operationId, string chartId) { Begins++; if (OnBegin != null) OnBegin(); }
        private void Step(Action mutation)
        { Steps++; if (Steps == FailAt && !After) throw new InvalidOperationException("injected"); mutation(); if (Steps == FailAt && After) throw new InvalidOperationException("injected"); }
        public int CreateAnchor(ScenePlot plot, bool trailing) { int id = Next++; Step(() => Objects.Add(id, "anchor")); return id; }
        public int CreateNode(SceneNode node, string token) { int id = Next++; Step(() => Objects.Add(id, token)); return id; }
        public int Group(IReadOnlyList<int> children)
        { int id = Next++; Step(() => { foreach (int child in children) Objects.Remove(child); Objects.Add(id, "root"); }); return id; }
        public void Persist(int rootId, string chartId, string xml) { Step(() => { RenderedChartMetadata.Decode(xml); Xml = xml; }); }
        public void Complete() { Step(() => { }); }
        public bool Rollback()
        {
            if (ThrowRollback) throw new InvalidOperationException("rollback failure");
            if (FailRollback) return false;
            foreach (int id in Objects.Keys.Where(id => id != 99).ToArray()) Objects.Remove(id); Xml = null; return true;
        }
    }
    [STAThread]
    public static int Main(string[] args)
    {
        try
        {
            Console.WriteLine("Native renderer runner architecture=" + (Environment.Is64BitProcess ? "x64" : "x86"));
            canonical = File.ReadAllText(args[0]); scene = File.ReadAllText(args[1]);
            if (args.Length == 4 && args[3] == "--interactive-office") { OfficeRoundtrip(args[2]); return 0; }
            Test("canonical projection snapshots match identity, raw values, geometry and style", () => {
                var p = ClusteredProjection.Validate(canonical, scene); Require(p.Scene.Nodes.Count == 5 && p.Model.Revision == 4 && p.Model.Series[0].Values[0] == -1.25, "projection");
            });
            foreach (string variant in new[] {
                ChangedScene(s => s.ModelRevision++), ChangedScene(s => ((BoxGeometry)s.Nodes[0].Geometry).X += 1),
                ChangedScene(s => ((BoxGeometry)s.Nodes[0].Geometry).Height += 1), ChangedScene(s => s.Nodes[0].Paint.Fill = "#FF0000"),
                ChangedScene(s => s.Nodes.RemoveAt(0)), ChangedScene(s => s.Nodes.Reverse()), ChangedScene(s => s.Diagnostics.Clear()) })
            {
                string input = variant; Test("mismatched but correctly hashed projection rejects before mutation", () => { var surface = new FakeSurface(); Reject(() => new NativeSceneRenderer(surface).Insert(canonical, input)); Require(surface.Begins == 0 && surface.Objects.Count == 1, "preflight mutated"); });
            }
            Test("canonical data tampering and non-JSON extensions reject before mutation", () => {
                var model = ChartCodec.Decode(canonical); model.Series[0].Values[0] += 1; Reject(() => ClusteredProjection.Validate(ChartCodec.Encode(model), scene));
                Reject(() => ClusteredProjection.Validate(canonical.Replace("\"revision\": 4", "\"revision\": 4,"), scene));
            });
            NativeRenderOutcome outcome = null; int steps = 0;
            Test("new chart owns only new objects and records zero bars without inventing geometry", () => {
                var surface = new FakeSurface(); outcome = new NativeSceneRenderer(surface).Insert(canonical, scene); steps = surface.Steps;
                Require(surface.Objects.Count == 2 && surface.Objects[99] == "user", "user object changed");
                Require(outcome.Record.Bindings.Count == 5 && outcome.Record.Bindings.Count(b => b.ShapeId.HasValue) == 4, "zero/null policy");
                Require(RenderedChartMetadata.Decode(surface.Xml).Bindings.Count == 5, "persistence");
            });
            for (int step = 1; step <= steps; step++) foreach (bool after in new[] { false, true })
            {
                int target = step; bool mode = after;
                Test("rollback preserves user object at operation " + step + (after ? " after" : " before"), () => {
                    var surface = new FakeSurface { FailAt = target, After = mode };
                    try { new NativeSceneRenderer(surface).Insert(canonical, scene); throw new Exception("Missing injected failure"); }
                    catch (NativeRenderException ex) { Require(ex.RollbackComplete && surface.Objects.Count == 1 && surface.Objects[99] == "user" && surface.Xml == null, "rollback failed"); }
                });
            }
            foreach (bool throws in new[] { false, true }) Test("incomplete rollback remains an explicit failed outcome", () => {
                var surface = new FakeSurface { FailAt = 3, FailRollback = !throws, ThrowRollback = throws };
                try { new NativeSceneRenderer(surface).Insert(canonical, scene); throw new Exception("Missing failure"); }
                catch (NativeRenderException ex) { Require(!ex.RollbackComplete, "partial rollback hidden"); }
            });
            Test("reentrant mutation is rejected", () => {
                var surface = new FakeSurface(); var renderer = new NativeSceneRenderer(surface); bool rejected = false;
                surface.OnBegin = () => { try { renderer.Insert(canonical, scene); } catch (InvalidOperationException) { rejected = true; } };
                renderer.Insert(canonical, scene); Require(rejected, "reentrant accepted");
            });
            Test("non-owner thread rejects before mutation", () => {
                var surface = new FakeSurface(); Exception error = null;
                var thread = new Thread(() => { try { new NativeSceneRenderer(surface).Insert(canonical, scene); } catch (Exception ex) { error = ex; } });
                thread.Start(); thread.Join(); Require(error is InvalidOperationException && surface.Begins == 0, "thread guard");
            });
            string xml = RenderedChartMetadata.Encode(outcome.Record);
            Test("persistence XML conforms to the published v1 structural schema", () => {
                var schemas = new System.Xml.Schema.XmlSchemaSet { XmlResolver = null };
                schemas.Add(RenderedChartMetadata.Namespace, "packages/chart-engine/schema/rendered-chart.v1.xsd");
                var document = new XmlDocument { XmlResolver = null, Schemas = schemas }; document.LoadXml(xml);
                document.Validate((sender, error) => { throw new Exception("XSD validation failed"); });
            });
            Test("versioned XML retains canonical numbers, Unicode, scene and stable tokens", () => {
                var decoded = RenderedChartMetadata.Decode(xml); Require(RenderedChartMetadata.Encode(decoded) == xml, "roundtrip");
                Require(ChartCodec.Decode(decoded.CanonicalJson).Dataset.Categories[2].Label == "日本語", "Unicode");
                Require(RenderedChartMetadata.NodeToken("a") != RenderedChartMetadata.NodeToken("A") && RenderedChartMetadata.NodeToken("a").All(c => "0123456789ABCDEF".Contains(c)), "uppercase-safe tokens");
                File.WriteAllText(Path.Combine(args[2], "rendered-chart.v1.xml"), xml);
            });
            foreach (Action<XmlDocument> change in new Action<XmlDocument>[] {
                d => d.DocumentElement.SetAttribute("version", "99"), d => d.DocumentElement.SetAttribute("chartId", Guid.NewGuid().ToString("D")),
                d => d.DocumentElement.SetAttribute("unknown", "1"), d => d.DocumentElement.LastChild.RemoveChild(d.DocumentElement.LastChild.FirstChild),
                d => ((XmlElement)d.DocumentElement.LastChild.FirstChild).SetAttribute("token", "BAD"),
                d => ((XmlElement)d.DocumentElement.LastChild.LastChild).SetAttribute("shapeId", ((XmlElement)d.DocumentElement.LastChild.FirstChild).GetAttribute("shapeId")) })
            {
                var update = change; Test("malformed or ambiguous mapping metadata rejects", () => { var doc = new XmlDocument(); doc.LoadXml(xml); update(doc); Reject(() => RenderedChartMetadata.Decode(doc.OuterXml)); });
            }
            Test("DTD and oversized metadata reject", () => {
                Reject(() => RenderedChartMetadata.Decode("<!DOCTYPE RenderedChart [<!ENTITY x SYSTEM 'file:///not-read'>]>" + xml));
                Reject(() => RenderedChartMetadata.Decode(new string('x', RenderedChartMetadata.MaximumCharacters + 1)));
            });
            Console.WriteLine("Native renderer core tests passed: " + passed + ". Office behavior not tested without --interactive-office.");
            RendererRibbonTests.Run(canonical, scene); return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex.GetType().Name + "; HRESULT=0x" + ex.HResult.ToString("X8") + "; native renderer validation failed."); return 1; }
    }
    private sealed class FailPersist : IEditableSceneSurface
    {
        private readonly IEditableSceneSurface inner;
        public FailPersist(IEditableSceneSurface inner) { this.inner = inner; }
        public void VerifyAccess() { inner.VerifyAccess(); } public void Begin(string op, string chart) { inner.Begin(op, chart); }
        public int CreateAnchor(ScenePlot plot, bool trailing) { return inner.CreateAnchor(plot, trailing); }
        public int CreateNode(SceneNode node, string token) { return inner.CreateNode(node, token); }
        public int Group(IReadOnlyList<int> ids) { return inner.Group(ids); }
        public void Persist(int id, string chart, string xml) { inner.Persist(id, chart, xml); throw new InvalidOperationException("Injected persistence failure"); }
        public bool Rollback() { return inner.Rollback(); } public void Complete() { inner.Complete(); }
    }
    private static void OfficeRoundtrip(string output)
    {
        PowerPoint.Presentation deck = null;
        try
        {
            var app = new PowerPoint.Application(); app.Visible = Office.MsoTriState.msoTrue;
            if (app.Presentations.Count != 0) throw new InvalidOperationException("Close user presentations before this test; none was changed.");
            Console.WriteLine("Office version=" + app.Version + "; process bitness=" + (Environment.Is64BitProcess ? "x64" : "x86"));
            deck = app.Presentations.Add(Office.MsoTriState.msoTrue); var slide = deck.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            var unrelated = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, 10, 10, 20, 20); unrelated.Name = "user-fixture";
            slide.Select();
            var surface = new PowerPointSceneSurface(deck, slide); var result = RendererRibbonCommand.Insert(app);
            var root = Enumerable.Range(1, slide.Shapes.Count).Select(i => slide.Shapes[i]).Single(s => s.Id == result.RootShapeId);
            Require(root.Type == Office.MsoShapeType.msoGroup && root.GroupItems.Count == 6 && slide.Shapes.Count == 2, "native group/zero bar");
            surface.ReadVerified(root); Console.WriteLine("PASS ribbon insertion command creates native editable bar group, unrelated object and metadata");
            string path = Path.Combine(output, "native-shape-roundtrip.pptx");
            deck.SaveAs(path, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation); deck.Close(); deck = null;
            deck = app.Presentations.Open(path, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue); slide = deck.Slides[1];
            root = Enumerable.Range(1, slide.Shapes.Count).Select(i => slide.Shapes[i]).Single(s => !String.IsNullOrEmpty(s.Tags[PowerPointSceneSurface.ChartTag]));
            surface = new PowerPointSceneSurface(deck, slide); var restored = surface.ReadVerified(root);
            Require(ChartCodec.Decode(restored.CanonicalJson).Series[0].Values[0] == -1.25 && restored.Bindings.Count == 5, "reopen data");
            var bar = Enumerable.Range(1, root.GroupItems.Count).Select(i => root.GroupItems[i]).First(s => !String.IsNullOrEmpty(s.Tags[PowerPointSceneSurface.NodeTag]));
            float original = bar.Left; string editedToken = bar.Tags[PowerPointSceneSurface.NodeTag];
            bar.Left += 5; float edited = bar.Left; Reject(() => surface.ReadVerified(root));
            deck.Save(); deck.Close(); deck = null;
            deck = app.Presentations.Open(path, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue); slide = deck.Slides[1];
            root = Enumerable.Range(1, slide.Shapes.Count).Select(i => slide.Shapes[i]).Single(s => !String.IsNullOrEmpty(s.Tags[PowerPointSceneSurface.ChartTag]));
            surface = new PowerPointSceneSurface(deck, slide);
            bar = Enumerable.Range(1, root.GroupItems.Count).Select(i => root.GroupItems[i]).Single(s => s.Tags[PowerPointSceneSurface.NodeTag] == editedToken);
            Require(Math.Abs(bar.Left - edited) <= .05, "native manual edit lost after save/reopen");
            Office.CustomXMLPart retainedPart = null;
            foreach (Office.CustomXMLPart part in root.CustomerData) if (part.NamespaceURI == RenderedChartMetadata.Namespace) { Require(retainedPart == null, "ambiguous persisted metadata"); retainedPart = part; }
            Require(retainedPart != null, "metadata missing after native edit");
            var retained = RenderedChartMetadata.Decode(retainedPart.XML);
            Require(retained.CanonicalJson == restored.CanonicalJson && retained.SceneJson == restored.SceneJson, "manual edit silently rewrote canonical data");
            Reject(() => surface.ReadVerified(root));
            bar.Left = original; surface.ReadVerified(root);
            Console.WriteLine("PASS native bar edit survives second save/reopen; canonical data retained and divergence detected");
            var duplicate = root.Duplicate()[1]; Reject(() => surface.ReadVerified(root)); duplicate.Delete(); surface.ReadVerified(root);
            Console.WriteLine("PASS duplicate identity fails closed without deleting either copy");
            string newId = Guid.NewGuid().ToString("D"); var model = ChartCodec.Decode(canonical); string oldId = model.ChartId; model.ChartId = newId;
            var next = SceneCodec.Decode(scene); next.ChartId = newId;
            foreach(var node in next.Nodes) node.Id = node.Id.Replace(oldId, newId);
            foreach(var diagnostic in next.Diagnostics) if(diagnostic.NodeId != null) diagnostic.NodeId = diagnostic.NodeId.Replace(oldId, newId);
            next.RenderHash = SceneCodec.ComputeHash(next);
            int before = slide.Shapes.Count;
            try { new NativeSceneRenderer(new FailPersist(surface)).Insert(ChartCodec.Encode(model), SceneCodec.Encode(next)); throw new Exception("Missing persistence fault"); }
            catch (NativeRenderException ex) { Require(ex.RollbackComplete && slide.Shapes.Count == before, "native rollback"); }
            surface.ReadVerified(root); Console.WriteLine("PASS native persistence failure rolls back new group and preserves original");
        }
        finally { if (deck != null) { deck.Saved = Office.MsoTriState.msoTrue; deck.Close(); } }
    }
}
