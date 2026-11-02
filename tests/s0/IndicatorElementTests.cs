using System;
using System.IO;
using PowerPointIntelligence.AddIn;
internal static class IndicatorElementTests {
 internal static void Run(){int checks=0;Action<bool,string> ok=(v,n)=>{if(!v)throw new Exception("FAIL "+n);checks++;};Action<Action,string> rejects=(a,n)=>{try{a();}catch(InvalidDataException){checks++;return;}throw new Exception("FAIL accepted "+n);};
 var s=new IndicatorSpec();var r=IndicatorCodec.Decode(IndicatorCodec.Encode(s));ok(r.Kind=="harvey"&&r.Filled==1&&r.Segments==4,"roundtrip");ok(s.Step(Int32.MaxValue).Filled==4,"upper clamp");ok(s.Step(Int32.MinValue).Filled==0,"lower clamp");ok(s.Filled==1,"immutable step");s.Segments=100;ok(s.Shortcut(1).Filled==13,"eighth rounded away");ok(s.Shortcut(8).Filled==100,"full shortcut");ok(s.Shortcut(0).Filled==0,"empty shortcut");s.Segments=7;ok(s.Shortcut(8).Filled==7,"shortcuts clamp small counts");s.Kind="checkbox";ok(s.Step(1).CheckState==1&&s.Step(-1).CheckState==2,"checkbox cycle");ok(s.Shortcut(2).CheckState==2,"checkbox shortcut");rejects(()=>s.Shortcut(3),"checkbox shortcut bound");s.CustomSymbol="&<>";ok(IndicatorCodec.Decode(IndicatorCodec.Encode(s)).CustomSymbol==s.CustomSymbol,"symbol XML escaped");
 rejects(()=>IndicatorCodec.Decode("<!DOCTYPE x [<!ENTITY a SYSTEM 'file:///x'>]><x>&a;</x>"),"DTD");rejects(()=>IndicatorCodec.Decode(IndicatorCodec.Encode(s).Replace("version=\"1\"","version=\"2\"")),"version");rejects(()=>IndicatorCodec.Decode(IndicatorCodec.Encode(s).Replace("indicator ","indicator extra=\"x\" ")),"unknown attribute");rejects(()=>IndicatorCodec.Decode(new string('x',4097)),"size bound");rejects(()=>s.Shortcut(9),"shortcut bound");
 foreach(int count in new[]{0,101}){var t=s.Copy();t.Segments=count;rejects(()=>t.Validate(),"segments");}foreach(float value in new[]{Single.NaN,Single.PositiveInfinity,5,721}){var t=s.Copy();t.SizePt=value;rejects(()=>t.Validate(),"size");}var bad=s.Copy();bad.FilledColor=-1;rejects(()=>bad.Validate(),"color");bad=s.Copy();bad.CustomSymbol="123456789";rejects(()=>bad.Validate(),"symbol bound");Console.WriteLine(checks+" indicator contract assertions passed; Office rendering unverified.");
 }
}

