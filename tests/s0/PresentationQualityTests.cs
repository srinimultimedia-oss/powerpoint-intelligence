using System;
using System.Linq;
using PowerPointIntelligence.AddIn;
internal static class PresentationQualityTests
{
    internal static int Run()
    {
        int n=0;Action<bool,string> check=(ok,name)=>{if(!ok)throw new Exception("FAIL quality "+name);Console.WriteLine("PASS quality "+name);n++;};
        Func<QualitySlideSnapshot> slide=()=>new QualitySlideSnapshot{Id=42,Index=1,Width=960,Height=540};
        Func<QualityShapeSnapshot> shape=()=>new QualityShapeSnapshot{Id=7,LocateShapeId=7,Left=20,Top=20,Width=100,Height=100};
        Func<QualitySlideSnapshot,QualityInspectionReport> inspect=sample=>PresentationQuality.Evaluate(new[]{sample});
        var s=slide();s.Shapes.Add(shape());check(inspect(s).Issues.Count==0,"clean bounds");
        s.Hidden=true;check(inspect(s).Issues.Single().Rule=="hidden-slide","hidden slide review");s.Hidden=false;
        var a=s.Shapes[0];a.Left=-1;check(inspect(s).Issues.Single().Rule=="off-slide","left outside");a.Left=900;check(inspect(s).Issues.Count==1,"right outside");a.Left=20;a.Top=500;check(inspect(s).Issues.Count==1,"bottom outside");a.Top=-1;check(inspect(s).Issues.Count==1,"top outside");a.Top=-.4f;check(inspect(s).Issues.Count==0,"geometry tolerance");
        a.Left=-10;a.Rotation=45;check(inspect(s).Issues.Count==0,"rotated geometry skipped");a.Rotation=0;a.TopLevel=false;check(inspect(s).Issues.Count==0,"nested geometry skipped");a.TopLevel=true;a.Left=20;a.Top=20;
        a.EditableText=true;check(inspect(s).Issues.Single().Rule=="empty-text","empty editable text");a.EditableText=false;check(inspect(s).Issues.Count==0,"decorative empty shape not flagged");a.HasText=true;a.FontSizes.Add(9);check(inspect(s).Issues.Single().Rule=="small-text","tiny text");a.FontSizes[0]=10;check(inspect(s).Issues.Count==0,"threshold inclusive");a.FontSizes.Add(float.NaN);a.FontSizes.Add(-2);check(inspect(s).Issues.Count==0,"unknown font sizes ignored");a.FontSizes.Add(7);check(inspect(s).Issues.Count==1,"mixed tiny font found");
        a.FontSizes.Clear();a.Picture=true;check(inspect(s).Issues.Single().Rule=="image-alt-text","missing alt text");a.HasAlternativeText=true;check(inspect(s).Issues.Count==0,"present alt text");a.FontScanTruncated=true;check(inspect(s).Truncated&&inspect(s).Issues.Single().Rule=="font-scan-limit","font inspection limit explicit");a.FontScanTruncated=false;
        a.Picture=false;a.EditableText=true;a.HasText=false;a.Left=-20;var limited=PresentationQuality.Evaluate(new[]{s},new QualityInspectionOptions{MaxIssues=1});check(limited.Truncated&&limited.Issues.Count==1,"issue cap");
        var s2=slide();check(PresentationQuality.Evaluate(new[]{s,s2},new QualityInspectionOptions{MaxSlides=1}).Truncated,"slide cap");s.Shapes.Add(shape());check(PresentationQuality.Evaluate(new[]{s},new QualityInspectionOptions{MaxShapes=1}).ShapesInspected==1,"shape cap");
        var issue=inspect(s).Issues.First();check(issue.SlideId==42&&issue.LocateShapeId==7&&issue.ShapeId==7,"locate identities");check(inspect(s).Limitations.Any(v=>v.Contains("rotated")),"rotation limitation visible");
        Action<Action,string> rejects=(action,name)=>{bool failed=false;try{action();}catch(ArgumentException){failed=true;}check(failed,name);};
        rejects(()=>PresentationQuality.Evaluate(null),"null input");rejects(()=>PresentationQuality.Evaluate(new QualitySlideSnapshot[]{null}),"null slide");
        var invalid=slide();invalid.Width=0;rejects(()=>inspect(invalid),"invalid slide dimensions");
        rejects(()=>PresentationQuality.Evaluate(new[]{s},new QualityInspectionOptions{MinimumFontPoints=float.NaN}),"invalid threshold");rejects(()=>PresentationQuality.Evaluate(new[]{s},new QualityInspectionOptions{MaxSlides=0}),"invalid slide limit");rejects(()=>PresentationQuality.Evaluate(new[]{s},new QualityInspectionOptions{MaxGroupDepth=17}),"invalid group limit");rejects(()=>PresentationQuality.Evaluate(new[]{s},new QualityInspectionOptions{MaxFontCharacters=0}),"invalid character limit");
        check(typeof(QualityShapeSnapshot).GetFields(System.Reflection.BindingFlags.NonPublic|System.Reflection.BindingFlags.Instance).All(f=>f.FieldType!=typeof(string)),"snapshot retains no text or alt-text contents");
        Console.WriteLine(n+" presentation quality assertions passed; native Office UI not exercised by pure rules tests.");return n;
    }
}

