using PowerPointIntelligence.Security;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using PowerPointIntelligence.AddIn;
using PowerPointIntelligence.Host;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
using Office=Microsoft.Office.Core;
internal static class LinkedChartDiagnostic
{
    internal static int Run(PowerPoint.Application app,string output,bool pauseOwnAddIn=false)
    {output=PathPolicy.TestPath(output);
        Directory.CreateDirectory(output);string log=Path.Combine(output,"linked-chart-stages.tsv");
        File.WriteAllText(log,"Stage\tField\tValue\r\n");
        Action<string,string,string> record=(stage,key,value)=>File.AppendAllText(log,stage+"\t"+key+"\t"+value.Replace('\r',' ').Replace('\n',' ')+"\r\n");
        PowerPoint.Presentation owned=null;PowerPoint.Shape shape=null;Office.COMAddIn paused=null;bool reconnect=false;
        Action<string> snapshot=stage=>{
            record(stage,"utc",DateTime.UtcNow.ToString("o",CultureInfo.InvariantCulture));Console.WriteLine("STEP linked-chart diagnostic: "+stage);
            if(stage=="after-workbook-close"){record(stage,"snapshot-policy","no COM reads during immediate OLE handoff");return;}
            try{
                PowerPoint.Chart chart=shape.Chart;PowerPoint.SeriesCollection collection=(PowerPoint.SeriesCollection)chart.SeriesCollection();record(stage,"series-count",collection.Count.ToString(CultureInfo.InvariantCulture));record(stage,"chart-data-is-linked",chart.ChartData.IsLinked.ToString());
                for(int i=1;i<=collection.Count;i++){
                    PowerPoint.Series item=collection.Item(i);object raw=item.Values;Array native=raw as Array;
                    record(stage,"series-"+i+"-native-type",raw==null?"null":raw.GetType().Name);
                    record(stage,"series-"+i+"-normalized-length",PowerPointHost.SeriesVector(raw).Length.ToString(CultureInfo.InvariantCulture));
                    if(native!=null){record(stage,"series-"+i+"-rank",native.Rank.ToString());record(stage,"series-"+i+"-lower-bound",native.GetLowerBound(0).ToString());}
                    if(stage=="source-bound")try{string formula=item.Formula??"";record(stage,"series-"+i+"-formula-ref-error",formula.IndexOf("#REF!",StringComparison.OrdinalIgnoreCase)>=0?"true":"false");record(stage,"series-"+i+"-formula-has-workbook",formula.IndexOf('[')>=0?"true":"false");}catch(Exception e){record(stage,"series-"+i+"-formula-read-error",e.GetType().Name+" 0x"+e.HResult.ToString("X8"));}
                }
            }catch(Exception e){record(stage,"snapshot-error",e.GetType().Name+" 0x"+e.HResult.ToString("X8"));}
        };
        try{
            record("environment","office",app.Version);record("environment","os",Environment.OSVersion.ToString());var officeFile=Path.Combine(app.Path,"POWERPNT.EXE");if(File.Exists(officeFile)){var build=System.Diagnostics.FileVersionInfo.GetVersionInfo(officeFile);record("environment","office-file-version",build.FileVersion);record("environment","office-product-version",build.ProductVersion);}record("environment","runtime",Environment.Version.ToString());record("environment","process",Environment.Is64BitProcess?"x64":"x86");
            var matches=app.COMAddIns.Cast<Office.COMAddIn>().Where(a=>(a.ProgId+" "+a.Description).Contains("PowerPointIntelligence")||(a.ProgId+" "+a.Description).Contains("PowerPoint Intelligence")).ToArray();
            record("environment","own-addin-matches",matches.Length.ToString());
            if(pauseOwnAddIn){if(matches.Length!=1)throw new InvalidOperationException("Own add-in identity is ambiguous; no add-in was paused.");paused=matches[0];reconnect=paused.Connect;if(reconnect)paused.Connect=false;record("environment","own-addin-paused",(!paused.Connect).ToString());}
            else record("environment","own-addin-paused","False");
            string source=Path.Combine(PathPolicy.TestPath(output),"diagnostic-fixture.csv");File.WriteAllText(source,"Category,Actual,Plan\r\nA,137.5,150\r\nB,20,25\r\nC,30,35\r\n");
            owned=app.Presentations.Add(Office.MsoTriState.msoTrue);var slide=owned.Slides.Add(1,PowerPoint.PpSlideLayout.ppLayoutBlank);slide.Select();
            var link=DataSources.Save(owned,DataSources.Create("csv",source));var table=LinkedChartData.Load(link);
            shape=slide.Shapes.AddChart2(-1,Office.XlChartType.xlColumnClustered,60,80,600,340,false);snapshot("chart-created");
            LinkedChartData.WriteDiagnosticObserver=snapshot;LinkedChartData.Write(shape,table,link);snapshot("write-returned");
            // Remove instrumentation before refresh: observations must not be required for correctness.
            LinkedChartData.WriteDiagnosticObserver=null;
            File.WriteAllText(source,"Category,Actual,Plan\r\nA,42.5,150\r\nB,20,25\r\nC,30,35\r\n");
            table=LinkedChartData.Load(link);LinkedChartData.Write(shape,table,link);snapshot("uninstrumented-refresh-returned");
            shape.Chart.HasTitle=true;shape.Chart.ChartTitle.Text="Editable diagnostic chart";
            string saved=Path.Combine(PathPolicy.TestPath(output),"linked-chart-success.pptx");
            owned.SaveAs(saved,PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);owned.Close();owned=null;shape=null;
            owned=app.Presentations.Open(saved,Office.MsoTriState.msoFalse,Office.MsoTriState.msoFalse,Office.MsoTriState.msoTrue);
            shape=owned.Slides[1].Shapes[1];snapshot("save-reopen");
            PowerPoint.SeriesCollection restored=(PowerPoint.SeriesCollection)shape.Chart.SeriesCollection();
            if(restored.Count!=2)throw new InvalidDataException("Reopened diagnostic series count differs.");
            double[][] expected={new[]{42.5,20.0,30.0},new[]{150.0,25.0,35.0}};
            for(int series=1;series<=2;series++){
                var values=PowerPointHost.NativeSeriesValues(restored,series);
                if(values.Length!=3)throw new InvalidDataException("Reopened diagnostic point count differs.");
                for(int point=0;point<3;point++)if(Convert.ToDouble(values[point],CultureInfo.InvariantCulture)!=expected[series-1][point])throw new InvalidDataException("Reopened diagnostic value differs; source values omitted.");
            }
            if(!shape.Chart.HasTitle||shape.Chart.ChartTitle.Text!="Editable diagnostic chart")throw new InvalidDataException("Native chart title edit did not survive reopening.");
            record("result","status","PASS initial write, uninstrumented refresh, native title edit and save/reopen data");
            Console.WriteLine("PASS linked-chart diagnostic: uninstrumented refresh, native edit and save/reopen");return 0;
        }catch(Exception error){
            record("result","status","FAIL");record("result","exception",error.GetType().Name+" 0x"+error.HResult.ToString("X8"));Console.Error.WriteLine(error);
            if(shape!=null)snapshot("write-failed");
            if(owned!=null)try{owned.SaveAs(Path.Combine(PathPolicy.TestPath(output),"linked-chart-failure.pptx"),PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);record("result","fixture-preserved","true");}catch(Exception e){record("result","fixture-save-error",e.GetType().Name+" 0x"+e.HResult.ToString("X8"));}
            return 1;
        }finally{
            LinkedChartData.WriteDiagnosticObserver=null;
            if(owned!=null)try{owned.Saved=Office.MsoTriState.msoTrue;owned.Close();}catch(Exception e){record("cleanup","exception",e.GetType().Name+" 0x"+e.HResult.ToString("X8"));}
            if(reconnect)try{paused.Connect=true;record("cleanup","own-addin-restored",paused.Connect.ToString());}catch(Exception e){record("cleanup","own-addin-restore-failed",e.GetType().Name+" 0x"+e.HResult.ToString("X8"));Console.Error.WriteLine("PowerPoint became unavailable before add-in restoration. Re-enable PowerPoint Intelligence through PowerPoint COM Add-ins after restarting.");}
        }
    }
}