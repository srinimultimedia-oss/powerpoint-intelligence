using System;
using System.IO;
using PowerPointIntelligence.AddIn;
internal static class NativeChartControlTests
{
    internal static void Run()
    {
        int count=0;Action<NativeAxisSettings> valid=s=>{s.Validate();count++;};Action<NativeAxisSettings> bad=s=>{try{s.Validate();}catch(InvalidDataException){count++;return;}throw new Exception("Invalid axis accepted.");};
        valid(new NativeAxisSettings());valid(new NativeAxisSettings {Automatic=false,Minimum=-100,Maximum=50,Major=10});valid(new NativeAxisSettings {Automatic=false,Logarithmic=true,Minimum=0.1,Maximum=1000,Major=10});
        bad(new NativeAxisSettings {Automatic=false,Minimum=100,Maximum=10});bad(new NativeAxisSettings {Minimum=Double.NaN});bad(new NativeAxisSettings {Maximum=Double.PositiveInfinity});bad(new NativeAxisSettings {Automatic=false,Major=0});bad(new NativeAxisSettings {Automatic=false,Logarithmic=true,Minimum=0});bad(new NativeAxisSettings {NumberFormat=""});bad(new NativeAxisSettings {NumberFormat="0\n"});bad(new NativeAxisSettings {NumberFormat=new String('x',101)});bad(new NativeAxisSettings {Automatic=false,Minimum=-Double.MaxValue,Maximum=Double.MaxValue});
        NativeAxisSettings.ValidateLogValues(new object[]{1,2.5,null,DBNull.Value,"",100});count++;
        foreach(var values in new[]{new object[]{0},new object[]{-1},new object[]{Double.NaN},new object[]{Double.PositiveInfinity},new object[]{"invalid"},new object[0],new object[]{null,DBNull.Value,""}}){bool rejected=false;try{NativeAxisSettings.ValidateLogValues(values);}catch(InvalidDataException){rejected=true;}if(!rejected)throw new Exception("Invalid logarithmic data accepted.");count++;}
        foreach(var type in NativeSeriesControls.Allowed){NativeSeriesControls.Validate(1,1,type);count++;}
        foreach(int index in new[]{0,3}){bool badSeries=false;try{NativeSeriesControls.Validate(index,2,Microsoft.Office.Core.XlChartType.xlLine);}catch(InvalidDataException){badSeries=true;}if(!badSeries)throw new Exception("Invalid series index accepted.");count++;}
        bool unsupported=false;try{NativeSeriesControls.Validate(1,1,Microsoft.Office.Core.XlChartType.xlPie);}catch(InvalidDataException){unsupported=true;}if(!unsupported)throw new Exception("Unsupported series type accepted.");count++;
        Console.WriteLine("PASS "+count+" native axis validation assertions. Office mutation and rollback not executed.");
    }
}
