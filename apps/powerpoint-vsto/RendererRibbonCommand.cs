using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using PowerPointIntelligence.Domain;
using PowerPointIntelligence.Rendering;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
namespace PowerPointIntelligence.AddIn
{
    // Explicit local sample command. No external input or AI-to-COM path.
    internal static class RendererRibbonCommand
    {
        private static string Resource(string name)
        {
            using(var stream=typeof(RendererRibbonCommand).Assembly.GetManifestResourceStream("PowerPointIntelligence.RendererSamples."+name))
            {
                if(stream==null) throw new InvalidDataException("Renderer sample resource missing.");
                using(var reader=new StreamReader(stream)) return reader.ReadToEnd();
            }
        }
        internal static ValidatedProjection Prepare()
        {
            // Validate the original pair before replacing identity: remapping must
            // not hide an invalid embedded resource or mismatched chart identity.
            var original=ClusteredProjection.Validate(Resource("canonical.json"),Resource("scene.json"));
            var model=original.Model;
            var scene=original.Scene;
            string oldId=scene.ChartId,newId=Guid.NewGuid().ToString("D");
            model.ChartId=newId;scene.ChartId=newId;
            foreach(var node in scene.Nodes) node.Id=newId+node.Id.Substring(oldId.Length);
            foreach(var diagnostic in scene.Diagnostics) if(diagnostic.NodeId!=null) diagnostic.NodeId=newId+diagnostic.NodeId.Substring(oldId.Length);
            scene.RenderHash=SceneCodec.ComputeHash(scene);
            return ClusteredProjection.Validate(ChartCodec.Encode(model),SceneCodec.Encode(scene));
        }
        internal static NativeRenderOutcome Insert(PowerPoint.Application application)
        {
            if(application==null) throw new ArgumentNullException("application");
            if(application.Presentations.Count==0) throw new InvalidOperationException("Open a presentation and select a slide first.");
            var slide=application.ActiveWindow.View.Slide as PowerPoint.Slide;
            if(slide==null) throw new InvalidOperationException("Select a slide in Normal view first.");
            var projection=Prepare();
            return new NativeSceneRenderer(new PowerPointSceneSurface(application.ActivePresentation,slide)).Insert(ChartCodec.Encode(projection.Model),SceneCodec.Encode(projection.Scene));
        }
        internal static ValidatedProjection Prepare(ChartDocument model)
        {
            ChartCodec.Validate(model);
            var a=model.Layout;
            double min=0,max=0;
            foreach(var series in model.Series)foreach(var value in series.Values)if(value.HasValue){min=Math.Min(min,value.Value);max=Math.Max(max,value.Value);}
            if(min==max){min=-1;max=1;}
            double range=max-min,span=a.Width/model.Dataset.Categories.Count,slot=span*.8/model.Series.Count;
            if(Double.IsInfinity(range))throw new InvalidDataException("Source numeric range is too large.");
            Func<double,double> y=value=>a.Y+(1-(value-min)/range)*a.Height;
            var scene=new RenderResult {SchemaVersion=1,ChartId=model.ChartId,ModelRevision=model.Revision,Unit="pt",Plot=new ScenePlot {Unit="pt",X=a.X,Y=a.Y,Width=a.Width,Height=a.Height},Nodes=new List<SceneNode>(),Diagnostics=new List<SceneDiagnostic>{new SceneDiagnostic{Code="AXES_NOT_RENDERED"}}};
            if(model.Style.ShowLegend)scene.Diagnostics.Add(new SceneDiagnostic{Code="LEGEND_NOT_RENDERED"});
            for(int ci=0;ci<model.Dataset.Categories.Count;ci++)for(int si=0;si<model.Series.Count;si++)
            {var value=model.Series[si].Values[ci];if(!value.HasValue)continue;scene.Nodes.Add(new SceneNode {Id=model.ChartId+"/bar:"+model.Series[si].Id+"/"+model.Dataset.Categories[ci].Id,Geometry=new BoxGeometry {Type="rectangle",X=a.X+ci*span+span*.1+si*slot+slot*.05,Y=Math.Min(y(0),y(value.Value)),Width=slot*.9,Height=Math.Abs(y(value.Value)-y(0))},Paint=new ScenePaint {Fill=model.Style.SeriesColors[si],Opacity=1}});}
            scene.RenderHash=SceneCodec.ComputeHash(scene);return ClusteredProjection.Validate(ChartCodec.Encode(model),SceneCodec.Encode(scene));
        }
        internal static NativeRenderOutcome Insert(PowerPoint.Application application,ChartDocument model)
        {
            var deck=application.ActivePresentation;var slide=application.ActiveWindow.View.Slide as PowerPoint.Slide;
            if(slide==null)throw new InvalidOperationException("Select a slide in the current presentation.");
            var projection=Prepare(model);
            return new NativeSceneRenderer(new PowerPointSceneSurface(deck,slide)).Insert(ChartCodec.Encode(projection.Model),SceneCodec.Encode(projection.Scene));
        }
    }
}
