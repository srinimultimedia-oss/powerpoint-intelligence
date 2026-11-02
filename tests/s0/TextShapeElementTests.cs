using System;
using System.Drawing;
using System.IO;
using PowerPointIntelligence.AddIn;
internal static class TextShapeElementTests
{
    private static int count;
    private static void Check(bool ok,string name){if(!ok)throw new Exception("FAIL text element "+name);count++;}
    private static void Near(float a,float b,string name){Check(Math.Abs(a-b)<.01f,name);}
    private static void Reject(Action action,string name){try{action();}catch(Exception ex) when(ex is InvalidDataException||ex is InvalidOperationException){count++;return;}throw new Exception("FAIL expected rejection "+name);}
    internal static void Run()
    {
        count=0;Near(TextShapePlan.ParseLength("24 pt",960,540),24,"point");Near(TextShapePlan.ParseLength("25.4mm",960,540),72,"mm");Near(TextShapePlan.ParseLength("2.54 cm",960,540),72,"cm");Near(TextShapePlan.ParseLength("1 IN",960,540),72,"inches");Near(TextShapePlan.ParseLength("20%w",960,540),192,"relative width");Near(TextShapePlan.ParseLength("50 %h",960,540),270,"relative height");Near(TextShapePlan.ParseLength("0.5 slidewidth",960,540),480,"slide width unit");Near(TextShapePlan.ParseLength("0.25 sh",960,540),135,"slide height unit");Near(TextShapePlan.ParseLength("12",960,540),12,"default points");
        foreach(string bad in new[]{"","-1pt","0pt","1e3pt","NaN","Infinity","1,5cm","20px","10001pt","1pt;write",".5in","20 % width",new string('1',65)})Reject(()=>TextShapePlan.ParseLength(bad,960,540),"invalid unit "+bad);
        Near(TextShapePlan.ParseLength("1/2 sh",960,540),270,"fractional slide height");Near(TextShapePlan.ParseLength("2 / 3 slidewidth",960,540),640,"fractional slide width");Reject(()=>TextShapePlan.ParseLength("1/0 sh",960,540),"zero denominator");Reject(()=>TextShapePlan.ParseLength("1/2 pt",960,540),"fraction requires slide unit");
        Reject(()=>TextShapePlan.ParseLength("12pt",Single.NaN,540),"invalid slide");
        var box=new RectangleF(200,200,100,60);var required=new SizeF(160,90);
        var fit=TextShapePlan.Fit(box,required,TextEdgeLocks.Left|TextEdgeLocks.Top,960,540);Near(fit.X,200,"left fixed");Near(fit.Y,200,"top fixed");Near(fit.Width,160,"width grows");Near(fit.Height,90,"height grows");
        fit=TextShapePlan.Fit(box,required,TextEdgeLocks.Right|TextEdgeLocks.Bottom,960,540);Near(fit.Right,300,"right fixed");Near(fit.Bottom,260,"bottom fixed");
        fit=TextShapePlan.Fit(box,required,TextEdgeLocks.None,960,540);Near(fit.Left+fit.Width/2,250,"horizontal center fixed");Near(fit.Top+fit.Height/2,230,"vertical center fixed");
        fit=TextShapePlan.Fit(box,new SizeF(40,20),TextEdgeLocks.Left|TextEdgeLocks.Top,960,540);Near(fit.Width,40,"width shrinks");Near(fit.Height,20,"height shrinks");Near(fit.Left,200,"left remains fixed while shrinking");
        fit=TextShapePlan.Fit(box,new SizeF(40,20),TextEdgeLocks.Right|TextEdgeLocks.Bottom,960,540);Near(fit.Right,300,"right remains fixed shrinking");Near(fit.Bottom,260,"bottom remains fixed shrinking");
        fit=TextShapePlan.Fit(box,new SizeF(40,20),(TextEdgeLocks)15,960,540);Check(fit==box,"all locked retains manual bounds");
        Reject(()=>TextShapePlan.Fit(box,required,TextEdgeLocks.Left|TextEdgeLocks.Right,960,540),"width locks reject overflow");Reject(()=>TextShapePlan.Fit(box,required,TextEdgeLocks.Top|TextEdgeLocks.Bottom,960,540),"height locks reject overflow");Reject(()=>TextShapePlan.Fit(new RectangleF(0,0,100,60),required,TextEdgeLocks.Right,960,540),"opposite growth offslide");Reject(()=>TextShapePlan.Fit(box,new SizeF(1000,1000),TextEdgeLocks.Left,960,540),"large growth offslide");Reject(()=>TextShapePlan.Fit(box,new SizeF(Single.NaN,20),TextEdgeLocks.None,960,540),"invalid measured geometry");Reject(()=>TextShapePlan.Fit(box,required,(TextEdgeLocks)16,960,540),"unknown edge lock");
        foreach(string kind in new[]{"textbox","rightPentagon","chevron","roundedRectangle"})for(int locks=0;locks<=15;locks++)
        {
            var spec=new TextShapeSpec {Kind=kind,Locks=(TextEdgeLocks)locks,Adjustment=.35f,AutoFit=locks%2==0};string xml=TextShapeCodec.Encode(spec);var copy=TextShapeCodec.Decode(xml);Check(copy.Kind==kind&&copy.Locks==spec.Locks&&copy.AutoFit==spec.AutoFit&&Math.Abs(copy.Adjustment-.35)<.001,"settings roundtrip");Check(TextShapeCodec.Encode(copy)==xml,"deterministic settings");
        }
        string valid=TextShapeCodec.Encode(new TextShapeSpec());Reject(()=>TextShapeCodec.Decode(valid.Replace("version=\"1\"","version=\"2\"")),"version");Reject(()=>TextShapeCodec.Decode(valid.Replace("element:v1","element:v2")),"namespace");Reject(()=>TextShapeCodec.Decode(valid.Replace("kind=\"textbox\"","kind=\"picture\"")),"unknown kind");Reject(()=>TextShapeCodec.Decode(valid.Replace("locks=\"7\"","locks=\"16\"")),"bad lock metadata");Reject(()=>TextShapeCodec.Decode(valid.Replace("autoFit=\"true\"","autoFit=\"yes\"")),"strict boolean");Reject(()=>TextShapeCodec.Decode(valid.Replace("adjustment=\"0.25\"","adjustment=\"NaN\"")),"NaN adjustment");Reject(()=>TextShapeCodec.Decode(valid.Replace("version=","extra=\"x\" version=")),"unknown XML field");Reject(()=>TextShapeCodec.Decode("<!DOCTYPE textElement [<!ENTITY test 'x'>]>"+valid),"DTD");Reject(()=>TextShapeCodec.Decode("<!DOCTYPE textElement SYSTEM 'file:///never-read'>"+valid),"external DTD");Reject(()=>TextShapeCodec.Decode("<!--unexpected-->"+valid),"comment");Reject(()=>TextShapeCodec.Decode(valid.Replace(" />","><extra /></textElement>")),"child element");Reject(()=>TextShapeCodec.Decode(valid.Substring(1)),"malformed XML");Reject(()=>TextShapeCodec.Decode(new string('x',4097)),"bounded XML");
        Reject(()=>TextShapeCodec.Encode(new TextShapeSpec {Adjustment=.01f}),"minimum adjustment");Reject(()=>TextShapeCodec.Encode(new TextShapeSpec {Adjustment=.51f}),"maximum adjustment");Check(TextShapePlan.Kind("Right Pentagon")=="rightPentagon"&&TextShapePlan.Kind("rounded-rectangle")=="roundedRectangle","kind aliases");Reject(()=>TextShapePlan.Kind("octagon"),"unsupported element");
        Console.WriteLine("Text shape element tests passed: "+count);
    }
}
