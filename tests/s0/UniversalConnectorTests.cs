using System;
using System.Drawing;
using System.IO;
using PowerPointIntelligence.AddIn;
internal static class UniversalConnectorTests
{
    internal static void Run()
    {
        int count=0;Action<bool> check=v=>{if(!v)throw new Exception("Connector assertion failed.");count++;};Action<Action> reject=a=>{try{a();}catch(InvalidDataException){count++;return;}throw new Exception("Invalid connector accepted.");};
        Func<ConnectorSpec> valid=()=>new ConnectorSpec {Begin=new ConnectorEndpoint {ShapeId=10,Site=1},End=new ConnectorEndpoint {Kind="point",ShapeId=11,Series=2,Point=3,Anchor="top"}};
        var s=valid();string xml=ConnectorCodec.Encode(s);check(ConnectorCodec.Encode(ConnectorCodec.Decode(xml))==xml);check(ConnectorCodec.Decode(xml).End.Series==2&&ConnectorCodec.Decode(xml).End.Point==3);
        foreach(string route in new[]{"straight","elbow","curve"}){s=valid();s.Route=route;check(ConnectorCodec.Decode(ConnectorCodec.Encode(s)).Route==route);}
        foreach(string dash in new[]{"solid","dash","dot","dashdot"}){s=valid();s.Dash=dash;check(ConnectorCodec.Decode(ConnectorCodec.Encode(s)).Dash==dash);}
        foreach(var anchor in new[]{"center","top","right","bottom","left"}){var p=ConnectorSpec.AnchorAt(new RectangleF(10,20,40,60),anchor);check(p.X>=10&&p.X<=50&&p.Y>=20&&p.Y<=80);}
        check(ConnectorSpec.AnchorAt(new RectangleF(10,20,40,60),"top")==new PointF(30,20));check(ConnectorSpec.AnchorAt(new RectangleF(10,20,40,60),"right")==new PointF(50,50));
        reject(()=>ConnectorSpec.AnchorAt(new RectangleF(Single.NaN,0,1,1),"center"));reject(()=>ConnectorSpec.AnchorAt(new RectangleF(0,0,-1,1),"center"));reject(()=>ConnectorSpec.AnchorAt(new RectangleF(0,0,1,1),"unknown"));
        foreach(string bad in new[]{"unknown","",null}){s=valid();s.Route=bad;reject(()=>s.Validate());s=valid();s.Dash=bad;reject(()=>s.Validate());}
        foreach(float weight in new[]{0f,-1f,13f,Single.NaN,Single.PositiveInfinity}){s=valid();s.Weight=weight;reject(()=>s.Validate());}
        s=valid();s.End=s.Begin;reject(()=>s.Validate());s=valid();s.End.Series=0;reject(()=>s.Validate());s=valid();s.Begin.Site=-1;reject(()=>s.Validate());s=valid();s.End.Point=10001;reject(()=>s.Validate());s=valid();s.Color=0x1000000;reject(()=>s.Validate());s=valid();s.Begin.ShapeId=0;reject(()=>s.Validate());
        foreach(string mutation in new[]{xml.Replace("version=\"2\"","version=\"3\""),xml.Replace("route=\"elbow\"","route=\"weird\""),xml.Replace("<begin ","<begin extra=\"x\" "),xml.Replace("auto=\"1\"","auto=\"true\""),xml.Replace("urn:powerpoint-intelligence:connector:v1","urn:wrong"),"<!DOCTYPE x [<!ENTITY ext SYSTEM 'file:///never'>]>"+xml,new string('x',16385)})reject(()=>ConnectorCodec.Decode(mutation));
        // Native attachment state, not stale metadata, controls live site bindings.
        var prior=new ConnectorEndpoint {Kind="site",ShapeId=10,Site=1,Rendered=true,X=40,Y=50};
        var reconciled=ConnectorBindingPlan.Reconcile(prior,true,20,3,new PointF(100,120));check(reconciled.Kind=="site"&&reconciled.ShapeId==20&&reconciled.Site==3&&prior.ShapeId==10);
        reconciled=ConnectorBindingPlan.Reconcile(prior,false,0,0,new PointF(40,50));check(reconciled.Kind=="free"&&reconciled.ShapeId==0&&reconciled.X==40);
        var tracked=new ConnectorEndpoint {Kind="point",ShapeId=11,Series=2,Point=3,X=40,Y=50,Rendered=true};reconciled=ConnectorBindingPlan.Reconcile(tracked,false,0,0,new PointF(40,50));check(reconciled.Kind=="point"&&reconciled.ShapeId==11);
        reconciled=ConnectorBindingPlan.Reconcile(tracked,false,0,0,new PointF(80,90));check(reconciled.Kind=="free"&&reconciled.ShapeId==0&&reconciled.X==80);
        reconciled=ConnectorBindingPlan.Reconcile(tracked,true,12,2,new PointF(80,90));check(reconciled.Kind=="site"&&reconciled.ShapeId==12&&reconciled.Site==2);
        var settled=new ConnectorEndpoint {Kind="site",ShapeId=12,Site=2,X=80,Y=90,Rendered=true};
        var stable=ConnectorBindingPlan.Reconcile(settled,true,12,2,new PointF(80.01f,89.99f));check(stable.X==80&&stable.Y==90); // Native float noise must not dirty a saved deck.
        var moved=ConnectorBindingPlan.Reconcile(settled,true,12,2,new PointF(81,92));check(moved.X==81&&moved.Y==92);
        var rebound=ConnectorBindingPlan.Reconcile(settled,true,13,2,new PointF(80.01f,89.99f));check(rebound.ShapeId==13&&rebound.X==80.01f);
        tracked.Rendered=false;reconciled=ConnectorBindingPlan.Reconcile(tracked,false,0,0,new PointF(80,90));check(reconciled.Kind=="free"&&reconciled.Rendered); // Unknown legacy drag history must not be overwritten.
        reject(()=>ConnectorBindingPlan.Reconcile(prior,true,0,1,new PointF(10,10)));reject(()=>ConnectorBindingPlan.Reconcile(prior,true,10,0,new PointF(10,10)));reject(()=>ConnectorBindingPlan.Reconcile(prior,false,0,0,new PointF(Single.NaN,10)));reject(()=>ConnectorBindingPlan.Reconcile(prior,false,0,0,new PointF(-1,10)));
        s=valid();s.Begin.Rendered=true;s.Begin.X=40;s.Begin.Y=50;var trackedCopy=ConnectorCodec.Decode(ConnectorCodec.Encode(s));check(trackedCopy.Begin.Rendered&&trackedCopy.Begin.X==40);
        string legacy=xml.Replace("version=\"2\"","version=\"1\"").Replace(" rendered=\"0\"","");check(!ConnectorCodec.Decode(legacy).Begin.Rendered);reject(()=>ConnectorCodec.Decode(xml.Replace(" rendered=\"0\"","")));reject(()=>ConnectorCodec.Decode(legacy.Replace("version=\"1\"","version=\"2\"")));
        reject(()=>ConnectorCodec.Decode(xml.Replace("rendered=\"0\"","rendered=\"yes\"")));reject(()=>ConnectorCodec.Decode(xml.Replace("weight=\"1.5\"","weight=\"NaN\"")));reject(()=>ConnectorCodec.Decode(xml.Replace("x=\"0\"","x=\"Infinity\"")));reject(()=>ConnectorCodec.Decode("<!--unexpected-->"+xml));reject(()=>ConnectorCodec.Decode(xml.Replace("</connector>","<other /></connector>")));
        reject(()=>ConnectorSpec.AnchorAt(new RectangleF(Single.MaxValue,0,Single.MaxValue,1),"right"));
        foreach(string beginArrow in new[]{"none","triangle","open","stealth"})foreach(string endArrow in new[]{"none","triangle","open","stealth"}){s=valid();s.BeginArrow=beginArrow;s.EndArrow=endArrow;var copy=ConnectorCodec.Decode(ConnectorCodec.Encode(s));check(copy.BeginArrow==beginArrow&&copy.EndArrow==endArrow);}
        s=valid();s.Begin.Kind="free";s.Begin.ShapeId=0;s.Begin.Site=0;s.Begin.X=100;s.Begin.Y=150;check(ConnectorCodec.Decode(ConnectorCodec.Encode(s)).Begin.Kind=="free");s.Begin.ShapeId=3;reject(()=>s.Validate());s=valid();s.Begin.X=-1;reject(()=>s.Validate());s=valid();s.Begin.Y=10001;reject(()=>s.Validate());s=valid();s.BeginArrow="circle";reject(()=>s.Validate());s=valid();s.EndArrow="invalid";reject(()=>s.Validate());
        s=new ConnectorSpec {Begin=new ConnectorEndpoint {Kind="free",X=100,Y=150},End=new ConnectorEndpoint {Kind="free",X=100,Y=150}};reject(()=>s.Validate());
        s=valid();s.End=s.Begin.Copy();s.End.Series=2;s.End.Anchor="left";reject(()=>s.Validate());s=valid();s.Begin=s.End.Copy();s.Begin.Site=7;reject(()=>s.Validate());
        reject(()=>ConnectorCodec.Encode(null));reject(()=>ConnectorCodec.Decode(xml.Replace("series=\"1\"","series=\"2147483648\"")));reject(()=>ConnectorCodec.Decode(xml.Replace("<end ","<begin ")));
        var nativeBounds=new RectangleF(100,100,80,40);
        var expected=new[]{new[]{new PointF(100,100),new PointF(160,80),new PointF(180,140),new PointF(120,160)},new[]{new PointF(180,100),new PointF(160,160),new PointF(100,140),new PointF(120,80)},new[]{new PointF(100,140),new PointF(120,80),new PointF(180,100),new PointF(160,160)},new[]{new PointF(180,140),new PointF(120,160),new PointF(100,100),new PointF(160,80)}};
        for(int flip=0;flip<4;flip++)for(int quarter=0;quarter<4;quarter++){var begin=ConnectorGeometryPlan.Position(nativeBounds,quarter*90,(flip&1)!=0,(flip&2)!=0,true);var end=ConnectorGeometryPlan.Position(nativeBounds,quarter*90,(flip&1)!=0,(flip&2)!=0,false);check(begin==expected[flip][quarter]);check(end==new PointF(280-begin.X,240-begin.Y));}
        foreach(float angle in new[]{15f,45f,137f,359f})for(int flip=0;flip<4;flip++){var begin=ConnectorGeometryPlan.Position(nativeBounds,angle,(flip&1)!=0,(flip&2)!=0,true);var end=ConnectorGeometryPlan.Position(nativeBounds,angle,(flip&1)!=0,(flip&2)!=0,false);check(Math.Abs(begin.X+end.X-280)<.001&&Math.Abs(begin.Y+end.Y-240)<.001);check(Math.Abs(Math.Pow(begin.X-end.X,2)+Math.Pow(begin.Y-end.Y,2)-8000)<.01);}
        check(ConnectorGeometryPlan.Position(nativeBounds,-90,false,false,true)==new PointF(120,160));check(ConnectorGeometryPlan.Position(nativeBounds,450,false,false,true)==new PointF(160,80));check(ConnectorGeometryPlan.Position(new RectangleF(100,100,0,0),90,false,false,true)==new PointF(100,100));check(ConnectorGeometryPlan.Position(new RectangleF(100,100,0,40),90,false,false,true)==new PointF(120,120));
        reject(()=>ConnectorGeometryPlan.Position(nativeBounds,Single.NaN,false,false,true));reject(()=>ConnectorGeometryPlan.Position(new RectangleF(0,0,-1,10),0,false,false,true));reject(()=>ConnectorGeometryPlan.Position(new RectangleF(Single.PositiveInfinity,0,10,10),0,false,false,true));
        Console.WriteLine("PASS "+count+" connector endpoint/geometry/style/strict-metadata assertions. Native attachment and chart-point geometry not executed.");
    }
}



