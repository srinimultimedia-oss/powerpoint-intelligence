using System;
using PowerPointIntelligence.AddIn;
internal static class ContextChartFormattingTests {
 internal static void Run(){int count=0;Action<bool> check=v=>{if(!v)throw new Exception("Chart font contract failed.");count++;};check(ContextChartFontPlan.Family("Arial")=="Arial");check(ContextChartFontPlan.Family(DBNull.Value)==null);check(ContextChartFontPlan.Family(3)==null);check(ContextChartFontPlan.Size(12.5)==12.5f);foreach(object n in new object[]{null,DBNull.Value,true,0,-2,Single.NaN,Double.PositiveInfinity,"bad",4001})check(!ContextChartFontPlan.Size(n).HasValue);check(ContextChartFontPlan.Flag(true)==true);check(ContextChartFontPlan.Flag(-1)==true);check(ContextChartFontPlan.Flag(0)==false);check(ContextChartFontPlan.Flag(-2)==null);check(ContextChartFontPlan.Flag(.1)==null);check(ContextChartFontPlan.Flag(DBNull.Value)==null);check(ContextChartFontPlan.Color(0xFFFFFF)==0xFFFFFF);foreach(object n in new object[]{null,DBNull.Value,true,-2,0x1000000,2.5,"bad"})check(!ContextChartFontPlan.Color(n).HasValue);Console.WriteLine(count+" chart font contract assertions passed; chart-wide native formatting remains unverified.");}
}

