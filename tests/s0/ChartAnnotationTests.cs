using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using PowerPointIntelligence.AddIn;
internal static class ChartAnnotationTests
{
    private static int passed;
    private static void Check(bool value,string label){if(!value)throw new Exception("Chart annotations: "+label);passed++;}
    private static void Reject(Action action,string label){try{action();}catch(InvalidDataException){Check(true,label);return;}throw new Exception("Chart annotations accepted "+label);}
    private static bool Near(float a,float b){return Math.Abs(a-b)<.001f;}
    internal static void Run()
    {
        var plot=new RectangleF(100,80,400,200);
        var line=ChartAnnotationPlan.Reference(plot,-20,80,30,false,false);Check(Near(line.Start.X,100)&&Near(line.End.X,500)&&Near(line.Start.Y,180)&&Near(line.End.Y,180),"reference follows native plot/axis span");
        Check(Near(ChartAnnotationPlan.Reference(plot,-20,80,80,false,false).Start.Y,80),"maximum at top");Check(Near(ChartAnnotationPlan.Reference(plot,-20,80,-20,false,false).Start.Y,280),"negative minimum at bottom");
        Check(Near(ChartAnnotationPlan.Reference(plot,0,100,20,false,true).Start.Y,120),"reversed vertical axis");
        var bar=ChartAnnotationPlan.Reference(plot,0,100,25,true,false);Check(Near(bar.Start.X,200)&&Near(bar.Start.Y,80)&&Near(bar.End.Y,280),"bar reference has vertical geometry");Check(Near(ChartAnnotationPlan.Reference(plot,0,100,25,true,true).Start.X,400),"reversed horizontal values");
        var values=new[]{10d,30d,-5d};var delta=ChartAnnotationPlan.Difference(plot,-10,50,values,1,2,false,false,false);Check(delta.Arrow&&delta.Text=="Δ +20"&&Near(delta.Start.X,300)&&Near(delta.Start.Y,213.33333f)&&Near(delta.End.Y,146.66667f),"difference uses selected numeric points and destination category");
        var decline=ChartAnnotationPlan.Difference(plot,-10,50,values,2,3,false,false,false);Check(decline.Text=="Δ -35"&&decline.Start.Y<decline.End.Y,"signed decline geometry");
        var horizontal=ChartAnnotationPlan.Difference(plot,-10,50,values,1,2,true,false,true);Check(Near(horizontal.Start.X,233.33333f)&&Near(horizontal.End.X,366.66667f)&&Near(horizontal.Start.Y,180),"bar difference and reversed categories");
        var onTick=ChartAnnotationPlan.Difference(plot,-10,50,values,1,3,false,false,false,false);Check(Near(onTick.Start.X,500),"on-tick line category geometry");
        Check(ChartAnnotationPlan.Reference(plot,0,.001,.000001,false,false).Text.Contains("E"),"small nonzero label never shown as zero");
        foreach(double bad in new[]{Double.NaN,Double.PositiveInfinity,-21,81})Reject(()=>ChartAnnotationPlan.Reference(plot,-20,80,bad,false,false),"invalid/out-of-range reference");
        Reject(()=>ChartAnnotationPlan.Reference(new RectangleF(0,0,-1,10),0,100,50,false,false),"negative plot width");Reject(()=>ChartAnnotationPlan.Reference(plot,5,5,5,false,false),"zero axis span");Reject(()=>ChartAnnotationPlan.Reference(plot,-Double.MaxValue,Double.MaxValue,0,false,false),"overflowing span");
        Reject(()=>ChartAnnotationPlan.Difference(plot,-10,50,values,0,2,false,false,false),"zero category index");Reject(()=>ChartAnnotationPlan.Difference(plot,-10,50,values,1,4,false,false,false),"absent category");Reject(()=>ChartAnnotationPlan.Difference(plot,-10,50,values,2,2,false,false,false),"same category");Reject(()=>ChartAnnotationPlan.Difference(plot,-10,50,new[]{10d,Double.NaN},1,2,false,false,false),"missing selected value");Reject(()=>ChartAnnotationPlan.Difference(plot,-10,50,new[]{10d,10d},1,2,false,false,false),"zero-length difference");
        var spec=new ChartAnnotationSpec {Kind="difference",ChartId=Guid.NewGuid().ToString("D"),Series=2,From=3,To=1};var tags=spec.Tags();var restored=ChartAnnotationSpec.Read(k=>tags[k]);Check(restored.Series==2&&restored.From==3&&restored.To==1&&restored.ChartId==spec.ChartId,"versioned association round trip");
        var reference=new ChartAnnotationSpec {Kind="reference",ChartId=spec.ChartId,Value=3.125};var prior=CultureInfo.CurrentCulture;try{CultureInfo.CurrentCulture=CultureInfo.GetCultureInfo("de-DE");var rt=reference.Tags();Check(rt["PPI_CA_VALUE"]=="3.125"&&ChartAnnotationSpec.Read(k=>rt[k]).Value==3.125,"culture-invariant persistence");}finally{CultureInfo.CurrentCulture=prior;}
        Reject(()=>ChartAnnotationPlan.Reference(new RectangleF(Single.MaxValue,0,Single.MaxValue,10),0,100,50,false,false),"overflowing plot extent");
        Reject(()=>new ChartAnnotationSpec {Kind="difference",ChartId=spec.ChartId,Series=1,From=1,To=2,Value=Double.NaN}.Tags(),"nonfinite unused metadata value");
        var unknown=new Dictionary<string,string>(tags);unknown["PPI_CA_SCHEMA"]="2";Reject(()=>ChartAnnotationSpec.Read(k=>unknown[k]),"future schema");unknown=new Dictionary<string,string>(tags);unknown["PPI_CA_CHART"]=Guid.Empty.ToString();Reject(()=>ChartAnnotationSpec.Read(k=>unknown[k]),"empty target identity");unknown=new Dictionary<string,string>(tags);unknown["PPI_CA_SERIES"]="bad";Reject(()=>ChartAnnotationSpec.Read(k=>unknown[k]),"malformed series");
        foreach(string kind in new[]{"chart-script","","REFERENCE"})Reject(()=>new ChartAnnotationSpec {Kind=kind,ChartId=spec.ChartId}.Tags(),"unknown typed kind");
        bool noHost=false;try{ChartAnnotations.Reference(null,2);}catch(InvalidOperationException){noHost=true;}Check(noHost,"absent host rejected before COM");
        Console.WriteLine("PASS "+passed+" chart annotation geometry/metadata assertions. Native PowerPoint behavior not tested.");
    }
}
