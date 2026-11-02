using System;
using System.Drawing;
using System.IO;
using System.Collections.Generic;
using PowerPointIntelligence.AddIn;
internal static class NativeFormattingTests
{
    internal static void Run()
    {
        int count=0;Action<bool> check=v=>{if(!v)throw new Exception("Native formatting assertion failed.");count++;};Action<Action> reject=a=>{try{a();}catch(InvalidDataException){count++;return;}throw new Exception("Invalid formatting accepted.");};
        check(NativeFormattingPlan.Rgb(Color.FromArgb(1,2,3))==0x030201);check(NativeFormattingPlan.ColorKind("color-text")=="text");check(NativeFormattingPlan.TextMode("text-vertical90")=="downward");check(NativeFormattingPlan.TextMode("270")=="upward");
        foreach(string kind in new[]{"fill","outline","text"})check(NativeFormattingPlan.ColorKind(kind)==kind);foreach(string mode in new[]{"left","center","right","justify","top","middle","bottom","horizontal","downward","upward"})check(NativeFormattingPlan.TextMode("text-"+mode)==mode);
        reject(()=>NativeFormattingPlan.ColorKind(null));reject(()=>NativeFormattingPlan.ColorKind("gradient"));reject(()=>NativeFormattingPlan.TextMode("unknown"));reject(()=>NativeFormattingPlan.ColorValue(-1));reject(()=>NativeFormattingPlan.ColorValue(0x1000000));reject(()=>NativeFormattingPlan.Bounds(0,0,0));reject(()=>NativeFormattingPlan.Bounds(501,0,0));reject(()=>NativeFormattingPlan.Bounds(1,2001,0));reject(()=>NativeFormattingPlan.Bounds(1,0,17));reject(()=>NativeFormattingPlan.Paragraph("top"));reject(()=>NativeFormattingPlan.Anchor("left"));reject(()=>NativeFormattingPlan.Orientation("center"));
        bool touched=false;check(!NativeFormatting.Color(null,"fill",()=>null)&&!touched);reject(()=>NativeFormatting.ApplyColor(null,"fill",-1));
        var state=new[]{0,0,0};var order=new List<int>();var changes=new[]{new NativeFormattingChange(()=>state[0]=1,()=>{state[0]=0;order.Add(0);}),new NativeFormattingChange(()=>{state[1]=2;throw new IOException("fixture");},()=>{state[1]=0;order.Add(1);}),new NativeFormattingChange(()=>state[2]=3,()=>state[2]=0)};bool failed=false;try{NativeFormattingTransaction.Run(changes);}catch(IOException){failed=true;}check(failed&&state[0]==0&&state[1]==0&&state[2]==0&&String.Join(",",order)=="1,0");
        bool incomplete=false;try{NativeFormattingTransaction.Run(new[]{new NativeFormattingChange(()=>{throw new IOException("fixture");},()=>{throw new IOException("restore");})});}catch(InvalidOperationException ex){incomplete=ex.Message.Contains("rollback was incomplete");}check(incomplete);reject(()=>NativeFormattingTransaction.Run(new NativeFormattingChange[0]));
        Console.WriteLine("PASS "+count+" native formatting validation/cancel/rollback assertions. PowerPoint and color dialog not executed.");
    }
}
