using System;
using System.IO;
using System.Linq;
using System.Xml;
using PowerPointIntelligence.AddIn;
using Office=Microsoft.Office.Core;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;

internal static class DataSourceTests
{
    private static int passed;
    private static void Check(bool value,string name)
    { if(!value)throw new Exception("Data Source: "+name);passed++;Console.WriteLine("PASS Data Source "+name); }
    private static void Reject(Action action,string name)
    { try { action(); } catch(Exception ex) when(ex is InvalidDataException || ex is IOException) { Check(true,name);return; } throw new Exception("Data Source accepted "+name); }
    internal static void Run(string output)
    {
        Directory.CreateDirectory(output);var registryPath=Path.Combine(Path.GetFullPath(output),"registry-fixture.csv");File.WriteAllText(registryPath,"Category,Value\r\nFixture,1\r\n");
        var registryFixture=DataSources.Create("csv",registryPath);
        var registryXml=DataSources.Encode(registryFixture);
        Check(DataSources.ReadRegistry(new[]{Tuple.Create(registryXml,false),Tuple.Create(registryXml,true)}).Count==1,"shape-attached reference is not a duplicate deck registry entry");
        Reject(()=>DataSources.ReadRegistry(new[]{Tuple.Create(registryXml,false),Tuple.Create(registryXml,false)}),"distinct deck registry entries sharing a source ID remain rejected");
        Check(DataSources.ReadRegistry(new[]{Tuple.Create(registryXml,true)}).Count==0,"orphan shape reference never becomes an implicit registry entry");

        var textual=LinkedChartData.Csv("Name,Role\nAlice,Owner\n,Reviewer",false);var cells=LinkedTableProjection.Cells(textual);Check(cells.GetLength(0)==3&&cells[1,1]=="Owner"&&cells[2,0]=="","textual linked-table headers, values and blanks preserved");cells[1,1]="changed";Check(textual.Rows[0][1]=="Owner","table projection does not mutate linked source");Reject(()=>textual.Validate(),"text-only table data cannot bypass chart numeric validation");
        var batch=LinkedChartData.Csv("Category,Actual,Target\nNorth,-2.5,0\nSouth,,8");var matrix=LinkedChartData.WorkbookValues(batch,false);
        Check(matrix.GetLength(0)==3&&matrix.GetLength(1)==3&&Equals(matrix[0,1],"Actual")&&Equals(matrix[1,0],"North"),"batch workbook projection preserves headers category labels and dimensions");
        Check(Equals(matrix[1,1],-2.5)&&Equals(matrix[1,2],0d)&&matrix[2,1]==null,"batch workbook projection preserves signed zero and missing numeric values");
        matrix[1,0]="edited";Check(batch.Rows[0][0]=="North","batch workbook projection does not mutate source table");
        var xy=LinkedChartData.Csv("X,Y\n1.5,2\n3,4");Check(Equals(LinkedChartData.WorkbookValues(xy,true)[1,0],1.5),"XY batch projection keeps numeric X values");
        Check(LinkedChartData.EmbeddedSourceAddress("Sales data","$A$1:$C$4")=="='Sales data'!$A$1:$C$4","chart source stays local to embedded workbook rather than temporary external workbook name");
        Check(LinkedChartData.EmbeddedSourceAddress("Owner's data","$A$1:$B$3")=="='Owner''s data'!$A$1:$B$3","embedded binding escapes worksheet apostrophes");
        Reject(()=>LinkedChartData.EmbeddedSourceAddress("[external]Sheet1","$A$1:$B$3"),"embedded binding rejects external workbook syntax");
        Reject(()=>LinkedChartData.EmbeddedSourceAddress("Sheet1","$A$1:$B$3,Sheet2!A1"),"embedded binding rejects range/formula injection");
        int workbookReads=0,workbookPauses=0;var embeddedObject=new object();Check(Object.ReferenceEquals(LinkedChartData.AwaitEmbeddedWorkbook(()=>{if(++workbookReads==1)throw new System.Runtime.InteropServices.COMException("busy",unchecked((int)0x80004005));return embeddedObject;},()=>workbookPauses++,2),embeddedObject)&&workbookReads==2&&workbookPauses==1,"read-only workbook access retries transient activation failure within budget");
        bool fatalWorkbook=false;try{LinkedChartData.AwaitEmbeddedWorkbook(()=>{throw new System.Runtime.InteropServices.COMException("disconnected",unchecked((int)0x80010108));},()=>{throw new Exception("must not retry");},2);}catch(System.Runtime.InteropServices.COMException){fatalWorkbook=true;}Check(fatalWorkbook,"workbook access never retries a disconnection");
        int exhaustedReads=0,exhaustedPauses=0;bool exhausted=false;try{LinkedChartData.AwaitEmbeddedWorkbook(()=>{exhaustedReads++;throw new System.Runtime.InteropServices.COMException("not ready",unchecked((int)0x80004005));},()=>exhaustedPauses++,3);}catch(System.Runtime.InteropServices.COMException){exhausted=true;}Check(exhausted&&exhaustedReads==3&&exhaustedPauses==2,"workbook-read retry budget ends without replaying activation");
        var readBack=Array.CreateInstance(typeof(object),new[]{matrix.GetLength(0),matrix.GetLength(1)},new[]{1,1});for(int r=0;r<matrix.GetLength(0);r++)for(int c=0;c<matrix.GetLength(1);c++)readBack.SetValue(matrix[r,c],r+1,c+1);
        LinkedChartData.ValidateWorkbookValues(readBack,matrix);Check(true,"one-based native workbook matrix is verified without array casts");
        Reject(()=>LinkedChartData.ValidateWorkbookValues(new object[1,1],matrix),"workbook write dimension mismatch is detected before chart binding");
        readBack.SetValue("-2.5",2,2);Reject(()=>LinkedChartData.ValidateWorkbookValues(readBack,matrix),"numeric source must remain a native number rather than text");readBack.SetValue("wrong",2,2);Reject(()=>LinkedChartData.ValidateWorkbookValues(readBack,matrix),"workbook write value mismatch is detected without logging source values");
        int closeCalls=0,verifications=0;LinkedChartData.CloseWorkbook(()=>closeCalls++,()=>{verifications++;return true;});Check(closeCalls==1&&verifications==0,"normal workbook close not replayed");
        closeCalls=0;LinkedChartData.CloseWorkbook(()=>{closeCalls++;throw new System.Runtime.InteropServices.COMException("fixture",unchecked((int)0x80010108));},()=>{verifications++;return true;});Check(closeCalls==1&&verifications==1,"disconnected close accepted only after independent commit verification");
        bool closeRejected=false;try{LinkedChartData.CloseWorkbook(()=>{throw new System.Runtime.InteropServices.COMException("fixture",unchecked((int)0x80010108));},()=>false);}catch(System.Runtime.InteropServices.COMException){closeRejected=true;}Check(closeRejected,"uncommitted disconnected close remains failure");
        bool unrelated=false;int unexpectedVerify=0;try{LinkedChartData.CloseWorkbook(()=>{throw new System.Runtime.InteropServices.COMException("fixture",unchecked((int)0x800706BA));},()=>{unexpectedVerify++;return true;});}catch(System.Runtime.InteropServices.COMException){unrelated=true;}Check(unrelated&&unexpectedVerify==0,"server unavailable is never treated as successful close");
        int interfaceClose=0;LinkedChartData.CloseWorkbook(()=>{interfaceClose++;throw new System.Runtime.InteropServices.COMException("fixture",unchecked((int)0x800706B5));},()=>true);Check(interfaceClose==1,"unknown workbook interface is not replayed and requires verified chart state");
        bool unknownRejected=false;try{LinkedChartData.CloseWorkbook(()=>{throw new System.Runtime.InteropServices.COMException("fixture",unchecked((int)0x800706B5));},()=>false);}catch(System.Runtime.InteropServices.COMException){unknownRejected=true;}Check(unknownRejected,"unverified unknown interface remains failure");
        var workbookFixture=new WorkbookCloseFixture();LinkedChartData.CloseNativeWorkbook(workbookFixture);Check(workbookFixture.Saved,"reflection close binds save without dynamic metadata scan");
        string folder=Path.Combine(Path.GetFullPath(output),"data-source-fixtures");Directory.CreateDirectory(folder);
        Check(DataSources.Options.Select(o=>o.Id).SequenceEqual(new[]{"excel","csv","json","xml"}),"four ordered format options");
        foreach(var option in DataSources.Options)
        {
            int saved=0;
            var cancelled=DataSources.ChooseAndLink(option.Id,o=>null,link=>{saved++;return link;});
            Check(cancelled==null && saved==0,option.Id+" cancellation makes no link");
            foreach(string extension in option.Extensions)
            {
                string path=Path.Combine(folder,"日本語 & source"+extension);File.WriteAllText(path,"owned reference-only fixture");
                bool picked=false;
                var result=DataSources.ChooseAndLink(option.Id,o=>{picked=o.Id==option.Id && o.Filter.Split('|')[1].Split(';').Contains("*"+extension);return path;},link=>{saved++;return link;});
                Check(picked && saved>0 && result.Format==option.Id && result.Path==path,option.Id+" filtered selection links "+extension);
                string xml=DataSources.Encode(result);var roundtrip=DataSources.Decode(xml);
                Check(DataSources.Encode(roundtrip)==xml && roundtrip.Id==result.Id && roundtrip.LinkedAtUtc.Kind==DateTimeKind.Utc && xml.Contains("&amp;"),extension+" versioned reference round trip and Unicode/XML escaping");
            }
            int before=saved;
            string wrong=Path.Combine(folder,"wrong.bin");File.WriteAllText(wrong,"fixture");
            Reject(()=>DataSources.ChooseAndLink(option.Id,o=>wrong,link=>{saved++;return link;}),option.Id+" wrong extension rejected");
            Check(saved==before,option.Id+" rejection occurs before persistence");
        }
        bool dialogOpened=false;
        Reject(()=>DataSources.ChooseAndLink("other",o=>{dialogOpened=true;return null;},l=>l),"unknown format rejected");
        Check(!dialogOpened,"unknown format does not open a picker");
        Reject(()=>DataSources.Create("csv",Path.Combine(folder,"absent.csv")),"missing file rejected");
        Reject(()=>DataSources.Create("csv","relative.csv"),"relative input rejected");
        var sample=DataSources.Create("csv",Path.Combine(folder,"日本語 & source.csv"));
        sample.Path=Path.Combine(folder,"moved-source.csv");
        Check(DataSources.Decode(DataSources.Encode(sample)).Path==sample.Path,"missing source reference survives metadata reopen");
        string valid=DataSources.Encode(sample);
        foreach(Action<XmlDocument> corrupt in new Action<XmlDocument>[] {
            x=>x.DocumentElement.SetAttribute("version","99"),x=>x.DocumentElement.SetAttribute("id",Guid.Empty.ToString()),
            x=>x.DocumentElement.SetAttribute("format","other"),x=>x.DocumentElement.SetAttribute("extra","1"),
            x=>x.DocumentElement.SetAttribute("linkedAtUtc","not a timestamp"),
            x=>x.DocumentElement.AppendChild(x.DocumentElement.FirstChild.CloneNode(true)),
            x=>x.DocumentElement.FirstChild.InnerText="relative.csv" })
        {
            var xml=new XmlDocument();xml.LoadXml(valid);corrupt(xml);Reject(()=>DataSources.Decode(xml.OuterXml),"ambiguous or unsupported metadata rejected");
        }
        Reject(()=>DataSources.Decode("<!DOCTYPE DataSource [<!ENTITY x SYSTEM 'file:///not-read'>]>"+valid),"DTD/external entities rejected");
        Reject(()=>DataSources.Decode(new String('x',DataSources.MaximumMetadataCharacters+1)),"oversized metadata rejected");
        var csv=LinkedChartData.Csv("Category,Revenue,Cost\r\n\"East, Japan\",137.5,\r\nWest,-8,20\r\n");
        Check(csv.Rows[0][0]=="East, Japan"&&csv.Number(0,1)==137.5&&csv.Number(0,2)==null&&csv.Number(1,1)==-8,"CSV consumes actual linked numbers, quoted labels, signed and missing values");
        Check(LinkedChartData.Csv("Category,Value\n\"日本語 \"\"quoted\"\"\",5").Rows[0][0]=="日本語 \"quoted\"","CSV Unicode and escaped quotes");
        foreach(string invalid in new[]{"Category,Value\nA,NaN","Category,Value\nA,Infinity","Category,Value\nA,text","Category,Value\nA,1,2","Category,Value\n\"A,1","Category,Value\n\"A\"x,1","Category,Value"})Reject(()=>LinkedChartData.Csv(invalid),"invalid source fails without sample fallback");
        foreach(string format in new[]{"csv","json","xml"})
        {
            string path=Path.Combine(folder,"actual-values."+format);
            string content=format=="csv"?"Category,Revenue,Cost\nEast,137.5,\nWest,-8,20":format=="json"?"{\"headers\":[\"Category\",\"Revenue\",\"Cost\"],\"rows\":[[\"East\",\"137.5\",null],[\"West\",\"-8\",\"20\"]]}":"<table><header><cell>Category</cell><cell>Revenue</cell><cell>Cost</cell></header><row><cell>East</cell><cell>137.5</cell><cell/></row><row><cell>West</cell><cell>-8</cell><cell>20</cell></row></table>";
            File.WriteAllText(path,content);var link=DataSources.Create(format,path);var table=LinkedChartData.Load(link);
            Check(table.Headers[1]=="Revenue"&&table.Number(0,1)==137.5&&table.Number(0,2)==null&&table.Number(1,1)==-8,format+" linked file parsed, no sample values");
            File.Delete(path);Reject(()=>LinkedChartData.Load(link),format+" moved file fails instead of using sample data");
        }
        string numericJson=Path.Combine(folder,"numeric.json");File.WriteAllText(numericJson,"{\"headers\":[\"Category\",\"Value\"],\"rows\":[[\"East\",137.5]]}");
        Check(LinkedChartData.Load(DataSources.Create("json",numericJson)).Number(0,1)==137.5,"JSON numeric tokens are read as actual values");
        var wide=new LinkedTable {Headers=new[]{"Notes","Sales","Region","Costs"},Rows=new[]{new[]{"ignore text","137.5","East","20"},new[]{"more notes","-8","West","5"}}};
        var subset=LinkedChartData.ChooseColumns(wide,t=>new[]{2,1,3});
        Check(subset.Headers.SequenceEqual(new[]{"Region","Sales","Costs"})&&subset.Rows[0][0]=="East"&&subset.Number(0,1)==137.5&&subset.Number(1,2)==5,"Excel selection chooses category and numeric columns, excludes unrelated text");
        Check(wide.Headers.Length==4&&wide.Rows[0][0]=="ignore text","column selection does not mutate source table");
        var oneSeries=LinkedChartData.SelectColumns(wide,2,new[]{3});
        Check(oneSeries.Headers.SequenceEqual(new[]{"Region","Costs"})&&oneSeries.Number(0,1)==20,"only explicitly selected series included");
        foreach(int[] selection in new[]{new int[0],new[]{2},new[]{2,2},new[]{2,1,1},new[]{2,9},new[]{-1,1},new[]{2,0}})
            Reject(()=>LinkedChartData.ChooseColumns(wide,t=>selection),"invalid Excel column selection rejected");
        bool columnSelectionCancelled=false;try {LinkedChartData.ChooseColumns(wide,t=>null);}catch(OperationCanceledException){columnSelectionCancelled=true;}
        Check(columnSelectionCancelled,"Excel column selection cancellation stops chart creation");
        Console.WriteLine("Data Source tests passed: "+passed+". Native file-picker and PowerPoint persistence behavior require Office validation.");
    }
    internal static void OfficeRoundtrip(PowerPoint.Application app,string output)
    {
        PowerPoint.Presentation owned=null;
        try
        {
            string directory=Path.Combine(output,"data-source-reference-fixtures");Directory.CreateDirectory(directory);
            owned=app.Presentations.Add(Office.MsoTriState.msoTrue);
            var slide=owned.Slides.Add(1,PowerPoint.PpSlideLayout.ppLayoutBlank);
            var ids=new System.Collections.Generic.List<string>();
            foreach(var option in DataSources.Options)
            {
                string source=Path.GetFullPath(Path.Combine(directory,"reference"+option.Extensions[0]));File.WriteAllText(source,"reference only");
                var link=DataSources.Save(owned,DataSources.Create(option.Id,source));ids.Add(link.Id);
                if(DataSources.Save(owned,DataSources.Create(option.Id,source)).Id!=link.Id)throw new Exception("Duplicate source created another link.");
            }
            if(DataSources.Read(owned).Count!=4 || slide.Shapes.Count!=0)throw new Exception("Data source linking changed unrelated slide objects.");
            string file=Path.GetFullPath(Path.Combine(output,"data-sources-roundtrip.pptx"));
            owned.SaveAs(file,PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);owned.Close();owned=null;
            owned=app.Presentations.Open(file);
            if(!DataSources.Read(owned).Select(l=>l.Id).OrderBy(x=>x).SequenceEqual(ids.OrderBy(x=>x)))throw new Exception("Source identities lost on reopen.");
            File.WriteAllText(Path.Combine(output,"data-sources-runtime.txt"),"PASS four source references, duplicate suppression, unchanged slide and native save/reopen. Native picker UI not exercised.");
        }
        finally { if(owned!=null){owned.Saved=Office.MsoTriState.msoTrue;owned.Close();} }
    }
    public sealed class WorkbookCloseFixture
    {
        public bool Saved;
        public void Close(object save,[System.Runtime.InteropServices.Optional]object file,[System.Runtime.InteropServices.Optional]object route){Saved=Convert.ToBoolean(save,System.Globalization.CultureInfo.InvariantCulture);}
    }

}

