using PowerPointIntelligence.Security;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using PowerPointIntelligence.AddIn;
using Office=Microsoft.Office.Core;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
internal static class DemoOfficeTests
{
    private static void Close(PowerPoint.Presentation owned)
    {
        if(owned==null)return;
        owned.Saved=Office.MsoTriState.msoTrue;owned.Close();
    }
    private static bool Disconnected(Exception error,int depth=0)
    {
        if(error==null||depth>16)return false;
        if(error.HResult==unchecked((int)0x800706BA)||error.HResult==unchecked((int)0x800706BE)||error.HResult==unchecked((int)0x80010108))return true;
        var aggregate=error as AggregateException;if(aggregate!=null){foreach(var inner in aggregate.InnerExceptions)if(Disconnected(inner,depth+1))return true;}
        return Disconnected(error.InnerException,depth+1);
    }
    [STAThread]
    public static int Main(string[] args)
    {
        try { foreach(int p in new int[]{1})if(p<args.Length)args[p]=PathPolicy.TestPath(args[p]); } catch(System.IO.InvalidDataException error) { Console.Error.WriteLine(error.Message); return 2; }

        if(args.Length!=2||(args[0]!="--interactive-office"&&args[0]!="--interactive-agenda"&&args[0]!="--interactive-elements"&&args[0]!="--interactive-properties"&&args[0]!="--diagnose-linked-chart"&&args[0]!="--diagnose-linked-chart-paused"&&args[0]!="--interactive-mapped-excel"))return 2;
        bool agendaOnly=args[0]=="--interactive-agenda";
        Directory.CreateDirectory(args[1]);var rows=new List<string>{"ToolId,Status,Evidence"};
        PowerPoint.Application app;
        try{app=new PowerPoint.Application();app.Visible=Office.MsoTriState.msoTrue;if(app.Presentations.Count>0)throw new InvalidOperationException("Save and close user presentations before runtime validation. No fixture was created.");}
        catch(Exception ex){if(args[0]=="--interactive-properties")rows.Add("context-properties-native,BLOCKED,"+ex.GetType().Name);else if(agendaOnly)rows.Add("chapter-agenda-native,BLOCKED,"+ex.GetType().Name);else if(args[0]=="--interactive-elements")foreach(var id in new[]{"productivity-native","text-elements-native","process-grid-native","indicators-native","universal-connectors-native"})rows.Add(id+",BLOCKED,"+ex.GetType().Name);else foreach(var tool in DemoTools.Catalog)rows.Add(tool.Id+",BLOCKED,"+ex.GetType().Name);File.WriteAllLines(Path.Combine(args[1],"features.csv"),rows);Console.Error.WriteLine(ex);return 1;}
        using(var mutationLease=PowerPointIntelligence.Host.OfficeMutationLease.ForAutomation(app)){
        Console.WriteLine("Explicit native tests own a cross-process mutation lease; installed automatic writes are deferred, context-property observation stays active.");
        if(args[0]=="--interactive-mapped-excel"){
            try{MappedExcelOfficeTests.Run(app,args[1]);File.WriteAllText(Path.Combine(args[1],"mapped-excel-result.txt"),"PASS");return 0;}
            catch(Exception error){File.WriteAllText(Path.Combine(args[1],"mapped-excel-result.txt"),"FAIL "+error.GetType().Name+" 0x"+error.HResult.ToString("X8"));Console.Error.WriteLine(error);return 1;}
        }
        if(args[0]=="--diagnose-linked-chart"||args[0]=="--diagnose-linked-chart-paused")return LinkedChartDiagnostic.Run(app,args[1],args[0]=="--diagnose-linked-chart-paused");
        if(args[0]=="--interactive-properties"){try{ContextPropertiesOfficeTests.Run(app,args[1]);ContextFormattingOfficeTests.Run(app,args[1]);rows.Add("context-properties-native,PASS,Installed automatic property box scope native edits and save reopen");File.WriteAllLines(Path.Combine(args[1],"features.csv"),rows);return 0;}catch(Exception error){rows.Add("context-properties-native,"+(error is ContextTestForegroundUnavailableException?"BLOCKED":"FAIL")+","+error.GetType().Name);File.WriteAllLines(Path.Combine(args[1],"features.csv"),rows);Console.Error.WriteLine(error);return 1;}}
        if(agendaOnly){try{ChapterAgendaOfficeTests.Run(app,args[1]);rows.Add("chapter-agenda-native,PASS,Native chapter hierarchy edits moves copy split and persistence");File.WriteAllLines(Path.Combine(args[1],"features.csv"),rows);return 0;}catch(Exception error){rows.Add("chapter-agenda-native,FAIL,"+error.GetType().Name);File.WriteAllLines(Path.Combine(args[1],"features.csv"),rows);Console.Error.WriteLine(error);return 1;}}
        int failed=0,executedCommands=0,blockedCommands=0;bool disconnected=false;
        foreach(var suite in new[]{new {Id="context-properties-native",Run=(Action<PowerPoint.Application,string>)ContextPropertiesOfficeTests.Run},new {Id="context-font-native",Run=(Action<PowerPoint.Application,string>)ContextFormattingOfficeTests.Run},new {Id="productivity-native",Run=(Action<PowerPoint.Application,string>)((a,o)=>{ProductivityToolOfficeTests.Run(a,o);})},new {Id="text-elements-native",Run=(Action<PowerPoint.Application,string>)TextShapeOfficeTests.Run},new {Id="process-grid-native",Run=(Action<PowerPoint.Application,string>)ProcessFlowOfficeTests.Run},new {Id="indicators-native",Run=(Action<PowerPoint.Application,string>)IndicatorOfficeTests.Run},new {Id="universal-connectors-native",Run=(Action<PowerPoint.Application,string>)UniversalConnectorOfficeTests.Run}})
        {
            try{if(disconnected)throw new InvalidOperationException("Prior Office RPC disconnection.");suite.Run(app,args[1]);rows.Add(suite.Id+",PASS,Native edits metadata and save reopen");}
            catch(Exception error){failed++;rows.Add(suite.Id+","+(error is ContextTestForegroundUnavailableException?"BLOCKED":"FAIL")+","+error.GetType().Name);Console.Error.WriteLine(suite.Id+": "+error);if(Disconnected(error))disconnected=true;}
            File.WriteAllLines(Path.Combine(args[1],"features.csv"),rows);
        }
        if(args[0]=="--interactive-elements")return failed==0?0:1;
        try { DataSourceTests.OfficeRoundtrip(app,args[1]); }
        catch(Exception ex) { failed++;rows.Add("data-sources,FAIL,"+ex.GetType().Name);Console.Error.WriteLine("Data source reference persistence failed: "+ex.GetType().Name); }
        try { if(disconnected)throw new InvalidOperationException("Office disconnected before chapter agenda regression.");ChapterAgendaOfficeTests.Run(app,args[1]);rows.Add("chapter-agenda-native,PASS,Hierarchy captions movement splitters and save reopen"); }
        catch(Exception ex) { failed++;rows.Add("chapter-agenda-native,FAIL,"+ex.GetType().Name);Console.Error.WriteLine("Chapter agenda regression: "+ex);if(Disconnected(ex))disconnected=true; }
        try { MappedExcelOfficeTests.Run(app,args[1]);rows.Add("mapped-excel-native,PASS,Explicit range refresh refusal and persistence"); }
        catch(Exception ex) { failed++;rows.Add("mapped-excel-native,FAIL,"+ex.GetType().Name);Console.Error.WriteLine("Mapped Excel regression: "+ex);if(Disconnected(ex))disconnected=true; }
        try { if(disconnected)throw new InvalidOperationException("Office disconnected before workflow regression.");WorkflowOfficeTests.Run(app,args[1]);rows.Add("workflow-native,PASS,Agenda and annotations edit save reopen"); }
        catch(Exception ex) { failed++;rows.Add("workflow-native,FAIL,"+ex.GetType().Name);Console.Error.WriteLine("Workflow native regression: "+ex);if(Disconnected(ex))disconnected=true; }
        foreach(var tool in DemoTools.Catalog)
        {
            if(disconnected){blockedCommands++;rows.Add(tool.Id+",BLOCKED,Prior PowerPoint RPC disconnection; command not executed");File.WriteAllLines(Path.Combine(args[1],"features.csv"),rows);continue;}
            var scenario=new DemoScenario(app,tool.Id,args[1]);PowerPoint.Presentation reopened=null;string outcome="PASS",evidence="Command-specific assertion and native save/reopen; fixture dialog inputs";
            try
            {
                executedCommands++;scenario.Prepare();scenario.ExecuteAndAssert(tool);
                var target=scenario.CreatedTour??scenario.Deck;int count=target.Slides[1].Shapes.Count;
                var file=Path.Combine(args[1],tool.Id+".pptx");target.SaveAs(file,PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                if(scenario.CreatedTour!=null){Close(scenario.CreatedTour);scenario.CreatedTour=null;}else{Close(scenario.Deck);scenario.Deck=null;}
                reopened=app.Presentations.Open(file);
                if(tool.Id=="demo-tour")scenario.AssertReport(reopened);
                if(reopened.Slides[1].Shapes.Count!=count)throw new Exception("Native object count changed after save/reopen");
            }
            catch(Exception ex){if(Disconnected(ex))disconnected=true;outcome="FAIL";evidence=ex.GetType().Name+" 0x"+ex.HResult.ToString("X8",CultureInfo.InvariantCulture);Console.Error.WriteLine(tool.Id+": "+ex);}
            finally
            {
                foreach(var owned in new[]{reopened,scenario.CreatedTour,scenario.Deck})
                {
                    try{Close(owned);}catch(System.Runtime.InteropServices.COMException ex){outcome="FAIL";evidence="Cleanup COM failure 0x"+ex.HResult.ToString("X8",CultureInfo.InvariantCulture);}
                }
                if(outcome!="PASS")failed++;
                rows.Add(tool.Id+","+outcome+","+evidence);File.WriteAllLines(Path.Combine(args[1],"features.csv"),rows);
                Console.WriteLine(outcome+" demo command "+tool.Id);Application.DoEvents();
            }
        }
        Console.WriteLine("Feature commands executed="+executedCommands+"; blocked="+blockedCommands+"; failures="+failed+". Native picker visuals and installed ribbon layout are separate checks.");
        return failed>0||blockedCommands>0?1:0;
        }
    }
}
