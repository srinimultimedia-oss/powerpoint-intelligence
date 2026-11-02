using PowerPointIntelligence.Security;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointIntelligence.AddIn
{
    [System.Runtime.Serialization.DataContract]
    internal sealed class LinkedTable
    {
        [System.Runtime.Serialization.DataMember(Name="headers",IsRequired=true)] public string[] Headers;
        [System.Runtime.Serialization.DataMember(Name="rows",IsRequired=true)] public string[][] Rows;
        internal double? Number(int row,int column)
        {
            string value=Rows[row][column]; if(String.IsNullOrWhiteSpace(value))return null;
            double number;if(!Double.TryParse(value,NumberStyles.Float,CultureInfo.InvariantCulture,out number)||Double.IsNaN(number)||Double.IsInfinity(number))throw new InvalidDataException("Data cells must contain finite numbers or be empty.");
            return number;
        }
        internal LinkedTable Validate(bool numeric=true)
        {
            if(Headers==null||Headers.Length<2||Headers.Length>15||Headers.Any(h=>String.IsNullOrWhiteSpace(h)||h.Length>200)||Rows==null||Rows.Length<1||Rows.Length>49)throw new InvalidDataException("Use headers and 1–49 data rows, with 2–15 columns.");
            for(int r=0;r<Rows.Length;r++) {if(Rows[r]==null||Rows[r].Length!=Headers.Length||Rows[r].Any(v=>v!=null&&v.Length>2000)||(numeric&&(String.IsNullOrWhiteSpace(Rows[r][0])||Rows[r][0].Length>200)))throw new InvalidDataException("Each row must match the header dimensions; chart rows require a category. Cells are limited to 2000 characters.");if(numeric)for(int c=1;c<Headers.Length;c++)Number(r,c);}
            return this;
        }
    }
    internal static class LinkedTableProjection
    {
        internal static string[,] Cells(LinkedTable table)
        {
            if(table==null)throw new InvalidDataException("Link a table data source first.");table.Validate(false);if((table.Rows.Length+1)*table.Headers.Length>400)throw new InvalidDataException("Native linked tables are limited to 400 cells.");var cells=new string[table.Rows.Length+1,table.Headers.Length];for(int c=0;c<table.Headers.Length;c++)cells[0,c]=table.Headers[c];for(int r=0;r<table.Rows.Length;r++)for(int c=0;c<table.Headers.Length;c++)cells[r+1,c]=table.Rows[r][c]??"";return cells;
        }
    }
    internal static class LinkedChartData
    {
        internal static LinkedTable SelectColumns(LinkedTable table,int category,int[] series)
        {
            if(table==null||table.Headers==null||table.Headers.Length<2||table.Headers.Length>15||table.Rows==null||table.Rows.Length<1||table.Rows.Length>49||table.Rows.Any(r=>r==null||r.Length!=table.Headers.Length))throw new InvalidDataException("Excel needs headers and a rectangular table of 1–49 data rows and 2–15 columns.");
            if(category<0||category>=table.Headers.Length||series==null||series.Length==0||series.Any(c=>c<0||c>=table.Headers.Length||c==category)||series.Distinct().Count()!=series.Length)throw new InvalidDataException("Choose a category column and at least one different data column.");
            var columns=new[]{category}.Concat(series).ToArray();
            return new LinkedTable {Headers=columns.Select(c=>table.Headers[c]).ToArray(),Rows=table.Rows.Select(r=>columns.Select(c=>r[c]).ToArray()).ToArray()}.Validate();
        }
        internal static LinkedTable ChooseColumns(LinkedTable table,Func<LinkedTable,int[]> choose)
        {
            // Selection precedes numeric validation: unrelated text columns are allowed.
            var selected=choose(table);if(selected==null)throw new OperationCanceledException();
            if(selected.Length<2)throw new InvalidDataException("Choose a category column and at least one data column.");
            return SelectColumns(table,selected[0],selected.Skip(1).ToArray());
        }
        internal static int[] PickMappedColumns(LinkedTable table) { return PickColumns(table); }
        private static int[] PickColumns(LinkedTable table)
        {
            using(var form=new Form {Text="Choose Excel columns",Width=560,Height=460,MinimumSize=new System.Drawing.Size(480,400),StartPosition=FormStartPosition.CenterScreen})
            using(var layout=new TableLayoutPanel {Dock=DockStyle.Fill,Padding=new Padding(16),ColumnCount=1,RowCount=6})
            using(var category=new ComboBox {Dock=DockStyle.Fill,DropDownStyle=ComboBoxStyle.DropDownList,AccessibleName="Category or X column"})
            using(var series=new CheckedListBox {Dock=DockStyle.Fill,CheckOnClick=true,AccessibleName="Data columns"})
            using(var error=new Label {Dock=DockStyle.Fill,AutoSize=true,ForeColor=System.Drawing.Color.Firebrick})
            using(var buttons=new FlowLayoutPanel {Dock=DockStyle.Fill,FlowDirection=FlowDirection.RightToLeft,AutoSize=true})
            using(var ok=new Button {Text="Build chart",AutoSize=true})
            using(var cancel=new Button {Text="Cancel",DialogResult=DialogResult.Cancel,AutoSize=true})
            {
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));layout.RowStyles.Add(new RowStyle(SizeType.Percent,100));layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                for(int c=0;c<table.Headers.Length;c++){string label=(c+1)+" — "+table.Headers[c];category.Items.Add(label);series.Items.Add(label,false);}
                category.SelectedIndex=0;
                category.SelectedIndexChanged+=(sender,args)=>{series.SetItemChecked(category.SelectedIndex,false);error.Text="";};
                series.ItemCheck+=(sender,args)=>{if(args.Index==category.SelectedIndex)args.NewValue=CheckState.Unchecked;};
                int[] result=null;
                ok.Click+=(sender,args)=>{
                    var chosen=new[]{category.SelectedIndex}.Concat(series.CheckedIndices.Cast<int>()).ToArray();
                    try {SelectColumns(table,chosen[0],chosen.Skip(1).ToArray());result=chosen;form.DialogResult=DialogResult.OK;form.Close();}
                    catch(InvalidDataException ex){error.Text=ex.Message;}
                };
                layout.Controls.Add(new Label {Text="Category / X / task column",AutoSize=true});layout.Controls.Add(category);
                layout.Controls.Add(new Label {Text="Select data columns (worksheet order). Only selected columns are used.\nBubble: Y, Size. Gantt: Start, Duration.",AutoSize=true});layout.Controls.Add(series);layout.Controls.Add(error);
                buttons.Controls.Add(cancel);buttons.Controls.Add(ok);layout.Controls.Add(buttons);form.Controls.Add(layout);form.AcceptButton=ok;form.CancelButton=cancel;
                return form.ShowDialog()==DialogResult.OK?result:null;
            }
        }
        internal static DataSourceLink Choose(PowerPoint.Presentation presentation)
        {
            var links=DataSources.Read(presentation).ToArray();
            if(links.Length==0)throw new InvalidOperationException("Link a file using Data Source before creating a chart.");
            if(links.Length==1)return links[0];
            using(var form=new Form {Text="Choose linked data source",Width=520,Height=280,StartPosition=FormStartPosition.CenterScreen})
            using(var list=new ListBox {Dock=DockStyle.Fill})
            using(var ok=new Button {Text="Use selected source",Dock=DockStyle.Bottom,DialogResult=DialogResult.OK})
            {foreach(var link in links)list.Items.Add(link.Path);list.SelectedIndex=0;form.Controls.Add(list);form.Controls.Add(ok);form.AcceptButton=ok;if(form.ShowDialog()!=DialogResult.OK)throw new OperationCanceledException();return links[list.SelectedIndex];}
        }
        internal static LinkedTable Load(DataSourceLink link,bool numeric=true)
        {
            link.Path=PathPolicy.Selected(link.Path);var info=new FileInfo(link.Path);if(!info.Exists||info.Length>20*1024*1024)throw new InvalidDataException("The linked source is missing or exceeds 20 MB. Relink it using Data Source.");
            if(link.Format=="excel"){if(link.Mapping!=null)return link.Mapping.Project(MappedExcelData.Read(link.Path,link.Mapping),numeric);var raw=Excel(link.Path);return numeric?ChooseColumns(raw,PickColumns):raw.Validate(false);}
            string text=File.ReadAllText(link.Path);
            if(link.Format=="csv")return Csv(text,numeric);
            if(link.Format=="json")
            {
                try { using(var stream=new MemoryStream(System.Text.Encoding.UTF8.GetBytes(text)))return ((LinkedTable)new System.Runtime.Serialization.Json.DataContractJsonSerializer(typeof(LinkedTable)).ReadObject(stream)).Validate(numeric); }
                catch(System.Runtime.Serialization.SerializationException) {throw new InvalidDataException("JSON requires headers and rows arrays; each row contains category then numeric text values.");}
            }
            if(link.Format=="xml")
            {
                var rows=new List<string[]>();
                using(var reader=XmlReader.Create(new StringReader(text),new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=20*1024*1024}))
                {var doc=new XmlDocument {XmlResolver=null};doc.Load(reader);if(doc.DocumentElement.Name!="table")throw new InvalidDataException("XML requires table/header/cell and table/row/cell elements.");foreach(XmlNode row in doc.DocumentElement.ChildNodes){if(row.NodeType!=XmlNodeType.Element)continue;if(row.Name!=(rows.Count==0?"header":"row"))throw new InvalidDataException("Invalid XML table row.");var cells=new List<string>();foreach(XmlNode cell in row.ChildNodes) {if(cell.NodeType!=XmlNodeType.Element)continue;if(cell.Name!="cell"||cell.ChildNodes.Cast<XmlNode>().Any(n=>n.NodeType==XmlNodeType.Element))throw new InvalidDataException("Invalid XML table cell.");cells.Add(cell.InnerText);}rows.Add(cells.ToArray());}}
                return FromRows(rows,numeric);
            }
            throw new InvalidDataException("Unsupported linked source format.");
        }
        private static LinkedTable FromRows(List<string[]> rows,bool numeric=true)
        {if(rows.Count<2)throw new InvalidDataException("Source needs a header and data rows.");return new LinkedTable {Headers=rows[0],Rows=rows.Skip(1).ToArray()}.Validate(numeric);}
        internal static LinkedTable Csv(string text,bool numeric=true)
        {
            var rows=new List<string[]>();var cells=new List<string>();var cell=new System.Text.StringBuilder();bool quoted=false,closed=false;
            for(int i=0;i<text.Length;i++)
            {char ch=text[i];if(quoted){if(ch=='"'){if(i+1<text.Length&&text[i+1]=='"'){cell.Append('"');i++;}else{quoted=false;closed=true;}}else cell.Append(ch);continue;}
                if(ch=='"'){if(cell.Length!=0||closed)throw new InvalidDataException("Malformed CSV quoting.");quoted=true;continue;}
                if(ch==','||ch=='\r'||ch=='\n'){cells.Add(cell.ToString());cell.Clear();closed=false;if(ch!=','){rows.Add(cells.ToArray());cells.Clear();if(ch=='\r'&&i+1<text.Length&&text[i+1]=='\n')i++;}continue;}
                if(closed)throw new InvalidDataException("Unexpected text after CSV quote.");cell.Append(ch);
            }
            if(quoted)throw new InvalidDataException("Unclosed CSV quote.");if(cell.Length>0||cells.Count>0||closed){cells.Add(cell.ToString());rows.Add(cells.ToArray());}return FromRows(rows,numeric);
        }
        private static LinkedTable Excel(string path)
        {
            dynamic excel=null,book=null;
            try {excel=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true));excel.Visible=false;excel.DisplayAlerts=false;excel.AutomationSecurity=3;book=excel.Workbooks.Open(path,0,true);dynamic range=book.Worksheets[1].UsedRange;int count=(int)range.Rows.Count,columns=(int)range.Columns.Count;if(count<2||count>50||columns<2||columns>15)throw new InvalidDataException("First Excel sheet must contain 2–50 rows and 2–15 columns.");var values=(Array)(object)range.Value2;var rows=new List<string[]>();for(int r=0;r<count;r++){var row=new string[columns];for(int c=0;c<columns;c++)row[c]=Convert.ToString(values.GetValue(r+values.GetLowerBound(0),c+values.GetLowerBound(1)),CultureInfo.InvariantCulture);rows.Add(row);}return new LinkedTable {Headers=rows[0],Rows=rows.Skip(1).ToArray()};}
            finally {try {if(book!=null)book.Close(false);}finally {if(excel!=null)System.Runtime.InteropServices.Marshal.ReleaseComObject((object)excel);}}
        }
        internal static object[,] WorkbookValues(LinkedTable table,bool numericX)
        {
            if(table==null)throw new ArgumentNullException(nameof(table));table.Validate();
            var values=new object[table.Rows.Length+1,table.Headers.Length];
            for(int c=0;c<table.Headers.Length;c++)values[0,c]=table.Headers[c];
            for(int r=0;r<table.Rows.Length;r++)for(int c=0;c<table.Headers.Length;c++)values[r+1,c]=c==0&&!numericX?(object)table.Rows[r][0]:table.Number(r,c);
            return values;
        }
        internal static void ValidateWorkbookValues(object actual,object[,] expected)
        {
            var values=actual as Array;if(values==null||values.Rank!=2||values.GetLength(0)!=expected.GetLength(0)||values.GetLength(1)!=expected.GetLength(1))throw new InvalidDataException("Embedded workbook write did not retain expected dimensions; source values were not logged.");
            for(int r=0;r<expected.GetLength(0);r++)for(int c=0;c<expected.GetLength(1);c++){
                object value=values.GetValue(r+values.GetLowerBound(0),c+values.GetLowerBound(1)),wanted=expected[r,c];bool same;
                if(wanted==null)same=value==null||value==DBNull.Value;
                else if(wanted is string)same=String.Equals(value as string,(string)wanted,StringComparison.Ordinal);
                else{double number; same=value!=null&&!(value is string)&&Double.TryParse(Convert.ToString(value,CultureInfo.InvariantCulture),NumberStyles.Float,CultureInfo.InvariantCulture,out number)&&!Double.IsNaN(number)&&!Double.IsInfinity(number)&&Math.Abs(number-Convert.ToDouble(wanted,CultureInfo.InvariantCulture))<1e-9;}
                if(!same)throw new InvalidDataException("Embedded workbook write verification failed; row="+(r+1)+"; column="+(c+1)+"; nativeType="+(value==null?"null":value.GetType().Name)+". Source values were not logged.");
            }
        }
        internal static string EmbeddedSourceAddress(string worksheet,string address)
        {
            if(String.IsNullOrWhiteSpace(worksheet)||worksheet.Length>31||worksheet.Any(Char.IsControl)||worksheet.IndexOfAny(new[]{'[',']',':','*','?','/','\\'})>=0||!System.Text.RegularExpressions.Regex.IsMatch(address??"",@"\A\$[A-Z]{1,3}\$[1-9][0-9]*:\$[A-Z]{1,3}\$[1-9][0-9]*\z"))throw new InvalidDataException("Invalid embedded worksheet/range binding.");
            return "='"+worksheet.Replace("'","''")+"'!"+address;
        }
        internal static object AwaitEmbeddedWorkbook(Func<object> read,Action pause,int attempts=31)
        {
            if(attempts<1||attempts>101)throw new ArgumentOutOfRangeException(nameof(attempts));
            for(int i=1;;i++){try{var value=read();if(value==null)throw new InvalidDataException("Embedded workbook is unavailable.");return value;}
                catch(System.Runtime.InteropServices.COMException error) when(i<attempts&&(error.HResult==unchecked((int)0x80004005)||error.HResult==unchecked((int)0x80010001)||error.HResult==unchecked((int)0x8001010A))){pause();}}
        }
        [ThreadStatic] private static Action<string> writeDiagnosticObserver;
        internal static Action<string> WriteDiagnosticObserver { get { return writeDiagnosticObserver; } set { writeDiagnosticObserver=value; } }
        private static void ObserveWrite(string stage){var observer=WriteDiagnosticObserver;if(observer!=null)observer(stage);}
        internal static void Write(PowerPoint.Shape shape,LinkedTable table,DataSourceLink link,bool rows=false,bool numericX=false)
        {
            table.Validate();var references=new List<DataSourceLink>();foreach(Office.CustomXMLPart reference in XmlPartUpdates.AttachedParts(shape,DataSources.Namespace))references.Add(DataSources.Decode(reference.XML));
            if(references.Count>1||references.Count==1&&!DataLinkInventory.Same(references[0],link))throw new InvalidDataException("Existing chart source is ambiguous or different; explicit relinking is required.");
            var chart=shape.Chart;chart.ChartData.Activate();dynamic book=AwaitEmbeddedWorkbook(()=>chart.ChartData.Workbook,()=>{Application.DoEvents();System.Threading.Thread.Sleep(100);});bool writeCompleted=false;ObserveWrite("workbook-activated");
            try {
                dynamic sheet=book.Worksheets[1];sheet.Cells.Clear();
                dynamic range=sheet.Range[sheet.Cells[1,1],sheet.Cells[table.Rows.Length+1,table.Headers.Length]];
                dynamic headers=sheet.Range[sheet.Cells[1,1],sheet.Cells[1,table.Headers.Length]];headers.NumberFormat="@";
                if(!numericX){dynamic labels=sheet.Range[sheet.Cells[2,1],sheet.Cells[table.Rows.Length+1,1]];labels.NumberFormat="@";}
                // One native write avoids queuing a chart refresh after each individual cell mutation.
                var written=WorkbookValues(table,numericX);range.Value2=written;ValidateWorkbookValues((object)range.Value2,written);ObserveWrite("worksheet-verified");sheet.Calculate();
                if(numericX){
                    var series=(PowerPoint.SeriesCollection)chart.SeriesCollection();while(series.Count>0)series.Item(series.Count).Delete();
                    bool bubble=chart.ChartType==Office.XlChartType.xlBubble;
                    for(int c=1;c<(bubble?2:table.Headers.Length);c++){
                        PowerPoint.Series item=series.NewSeries();item.Name=table.Headers[c];
                        item.XValues=Enumerable.Range(0,table.Rows.Length).Select(r=>table.Number(r,0).Value).ToArray();
                        item.Values=Enumerable.Range(0,table.Rows.Length).Select(r=>(object)table.Number(r,c)).ToArray();
                        if(bubble)item.BubbleSizes=Enumerable.Range(0,table.Rows.Length).Select(r=>table.Number(r,2).Value).ToArray();
                    }
                }else {chart.SetSourceData(EmbeddedSourceAddress((string)sheet.Name,(string)range.Address[true,true,1,false]),rows?PowerPoint.XlRowCol.xlRows:PowerPoint.XlRowCol.xlColumns);}
                ObserveWrite("source-bound");
                // Excel accepting the cells does not establish that the chart consumed them.
                // Await the actual values while its embedded workbook is still available.
                if(!WaitForChartValues(chart,table,rows,true))throw new InvalidOperationException("Native chart did not acknowledge the embedded workbook before close. "+lastCommitDiagnostic+". Source values were not logged.");
                writeCompleted=true;ObserveWrite("chart-values-acknowledged");
            }
            finally {ObserveWrite("before-workbook-close");CloseWorkbook(()=>CloseNativeWorkbook((object)book),()=>writeCompleted&&WaitForChartValues(shape.Chart,table,rows,true));ObserveWrite("after-workbook-close");}
            if(references.Count==0){var part=shape.CustomerData.Add();if(!part.LoadXML(DataSources.Encode(link)))throw new InvalidDataException("Chart source reference could not be saved.");}
            shape.Tags.Add("PPI_SOURCE_ROWS",rows?"1":"0");shape.Tags.Add("PPI_SOURCE_NUMERIC_X",numericX?"1":"0");
            if(!WaitForChartValues(shape.Chart,table,rows,true))throw new InvalidOperationException("PowerPoint did not commit the linked chart values within eight seconds. "+lastCommitDiagnostic+". No source values were logged; no sample fallback was used.");
        }
        internal static CultureInfo WorkbookAutomationCulture {get{return CultureInfo.GetCultureInfo("en-US");}}
        internal static void CloseNativeWorkbook(object workbook)
        {
            if(workbook==null)throw new ArgumentNullException(nameof(workbook));
            try{workbook.GetType().InvokeMember("Close",System.Reflection.BindingFlags.InvokeMethod|System.Reflection.BindingFlags.Public|System.Reflection.BindingFlags.Instance|System.Reflection.BindingFlags.OptionalParamBinding,null,workbook,new object[]{true,Type.Missing,Type.Missing},WorkbookAutomationCulture);}
            catch(System.Reflection.TargetInvocationException ex) when(ex.InnerException is System.Runtime.InteropServices.COMException){System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(ex.InnerException).Throw();throw;}
        }
        internal static void CloseWorkbook(Action close,Func<bool> committed)
        {
            try{Host.PowerPointHost.RetryRejectedCall(close,()=>{Application.DoEvents();System.Threading.Thread.Sleep(100);},151);}
            catch(System.Runtime.InteropServices.COMException ex) when(ex.HResult==unchecked((int)0x80010108)||ex.HResult==unchecked((int)0x800706B5)){
                // Close may have completed before the workbook proxy disconnected. Never replay it.
                if(!committed())throw;System.Diagnostics.Trace.WriteLine("Embedded workbook close disconnected; native chart values independently verified.");
            }
        }
        [ThreadStatic] private static string lastCommitDiagnostic;
        private static bool WaitForChartValues(PowerPoint.Chart chart,LinkedTable table,bool rows,bool refresh)
        {
            var watch=System.Diagnostics.Stopwatch.StartNew();bool refreshed=false;
            do {
                try {var series=(PowerPoint.SeriesCollection)chart.SeriesCollection();
                    int expectedSeries=rows?table.Rows.Length:(chart.ChartType==Office.XlChartType.xlBubble||chart.ChartType==Office.XlChartType.xlPie||chart.ChartType==Office.XlChartType.xlDoughnut)?1:table.Headers.Length-1;
                    bool ready=series.Count==expectedSeries;lastCommitDiagnostic="stage=series-count; expected="+expectedSeries+"; actual="+series.Count;
                    for(int si=1;ready&&si<=series.Count;si++){
                        var actual=Host.PowerPointHost.NativeSeriesValues(series,si);int expectedPoints=rows?table.Headers.Length-1:table.Rows.Length;ready=actual.Length==expectedPoints;lastCommitDiagnostic="stage=point-count; series="+si+"; expected="+expectedPoints+"; actual="+actual.Length;
                        for(int pi=0;ready&&pi<actual.Length;pi++) {var expected=rows?table.Number(si-1,pi+1):table.Number(pi,si);ready=expected.HasValue?actual[pi]!=null&&Math.Abs(Convert.ToDouble(actual[pi],CultureInfo.InvariantCulture)-expected.Value)<1e-9:true;if(!ready)lastCommitDiagnostic="stage=numeric-match; series="+si+"; point="+(pi+1)+"; nativeType="+(actual[pi]==null?"null":actual[pi].GetType().Name);}
                    }
                    if(ready)return true;
                    if(refresh&&!refreshed&&watch.ElapsedMilliseconds>=250){refreshed=true;chart.Refresh();}
                } catch(System.Runtime.InteropServices.COMException ex) when(ex.HResult==unchecked((int)0x80010001)||ex.HResult==unchecked((int)0x8001010A)) { }
                Application.DoEvents();System.Threading.Thread.Sleep(100);
            }while(watch.ElapsedMilliseconds<8000);
            return false;
        }
        internal static void Persist(PowerPoint.Shape shape,LinkedTable table,DataSourceLink link)
        {
            var doc=new XmlDocument {XmlResolver=null};var root=doc.CreateElement("LinkedTable","urn:powerpoint-intelligence:linked-table:v1");doc.AppendChild(root);root.SetAttribute("version","1");root.SetAttribute("sourceId",link.Id);
            using(var stream=new MemoryStream()){new System.Runtime.Serialization.Json.DataContractJsonSerializer(typeof(LinkedTable)).WriteObject(stream,table);root.InnerText=System.Text.Encoding.UTF8.GetString(stream.ToArray());}
            if(!shape.CustomerData.Add().LoadXML(DataSources.Encode(link))||!shape.CustomerData.Add().LoadXML(doc.OuterXml))throw new InvalidDataException("Linked diagram data could not be saved.");
        }
    }
    internal sealed class DataSourceOption
    {
        internal string Id { get; private set; }
        internal string Label { get; private set; }
        internal string Filter { get; private set; }
        internal string Icon { get; private set; }
        internal IReadOnlyList<string> Extensions { get; private set; }
        internal DataSourceOption(string id, string label, string icon, params string[] extensions)
        {
            Id=id; Label=label; Icon=icon; Extensions=Array.AsReadOnly(extensions);
            string patterns=String.Join(";",extensions.Select(e=>"*"+e));
            Filter=label+" ("+patterns+")|"+patterns;
        }
    }
    internal sealed class DataSourceLink
    {
        internal string Id { get; set; }
        internal string Format { get; set; }
        internal string Path { get; set; }
        internal DateTime LinkedAtUtc { get; set; }
        internal ExcelRangeMapping Mapping { get; set; }
    }
    internal static class DataSources
    {
        internal const string Namespace="urn:powerpoint-intelligence:data-source:v1";
        internal const int MaximumMetadataCharacters=131072;
        internal static readonly IReadOnlyList<DataSourceOption> Options=Array.AsReadOnly(new[] {
            new DataSourceOption("excel","Excel sheet","source-excel",".xlsx",".xls",".xlsm",".xlsb"),
            new DataSourceOption("csv","CSV file","source-csv",".csv"),
            new DataSourceOption("json","JSON file","source-json",".json"),
            new DataSourceOption("xml","XML file","source-xml",".xml")
        });
        internal static DataSourceOption Option(string id)
        { var value=Options.SingleOrDefault(o=>o.Id==id); if(value==null)throw new InvalidDataException("Unsupported data source format.");return value; }
        private static void Require(bool valid) { if(!valid)throw new InvalidDataException("Invalid data source link metadata."); }
        private static void Validate(DataSourceLink link)
        {
            Guid id;
            Require(link!=null && Guid.TryParseExact(link.Id,"D",out id) && id!=Guid.Empty && link.LinkedAtUtc.Kind==DateTimeKind.Utc);
            var option=Option(link.Format);
            if(link.Mapping!=null){Require(link.Format=="excel");link.Mapping.Validate();}
            Require(!String.IsNullOrWhiteSpace(link.Path) && link.Path.Length<=32767 && Path.IsPathRooted(link.Path));
            Require(!link.Path.StartsWith(@"\\?\",StringComparison.Ordinal) && !link.Path.StartsWith(@"\\.\",StringComparison.Ordinal));
            Require(String.Equals(PathPolicy.Selected(link.Path),link.Path,StringComparison.OrdinalIgnoreCase));
            Require(option.Extensions.Contains(Path.GetExtension(link.Path),StringComparer.OrdinalIgnoreCase));
        }
        internal static DataSourceLink Create(string format,string path)
        {
            var option=Option(format);
            if(String.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path))throw new InvalidDataException("Choose an existing source file.");
            var link=new DataSourceLink{Id=Guid.NewGuid().ToString("D"),Format=option.Id,Path=PathPolicy.Selected(path),LinkedAtUtc=DateTime.UtcNow};
            Validate(link);
            if(!new FileInfo(link.Path).Exists)throw new FileNotFoundException("The selected source file is unavailable.");
            return link;
        }
        // A file reference only: selecting a source never opens its workbook,
        // evaluates macros/formulas or interprets JSON/XML/CSV contents.
        internal static DataSourceLink ChooseAndLink(string format,Func<DataSourceOption,string> pick,Func<DataSourceLink,DataSourceLink> save)
        {
            var option=Option(format);string path=pick(option);
            if(path==null)return null;
            var link=Create(format,path);return save(link);
        }
        internal static string PickFile(DataSourceOption option)
        {
            using(var dialog=new OpenFileDialog{Title="Link "+option.Label,Filter=option.Filter,FilterIndex=1,
                DefaultExt=option.Extensions[0].Substring(1),CheckFileExists=true,CheckPathExists=true,
                Multiselect=false,ValidateNames=true,RestoreDirectory=true})
                return dialog.ShowDialog()==DialogResult.OK?dialog.FileName:null;
        }
        internal static string Encode(DataSourceLink link)
        {
            Validate(link);
            using(var text=new StringWriter(CultureInfo.InvariantCulture))
            {
                using(var writer=XmlWriter.Create(text,new XmlWriterSettings{OmitXmlDeclaration=true}))
                {
                    writer.WriteStartElement("DataSource",Namespace);writer.WriteAttributeString("version",link.Mapping==null?"1":"2");
                    writer.WriteAttributeString("id",link.Id);writer.WriteAttributeString("format",link.Format);
                    writer.WriteAttributeString("linkedAtUtc",link.LinkedAtUtc.ToString("o",CultureInfo.InvariantCulture));
                    writer.WriteElementString("Path",Namespace,link.Path);if(link.Mapping!=null)link.Mapping.Write(writer,Namespace);writer.WriteEndElement();
                }
                var xml=text.ToString();Require(xml.Length<=MaximumMetadataCharacters);return xml;
            }
        }
        internal static DataSourceLink Decode(string xml)
        {
            Require(!String.IsNullOrWhiteSpace(xml) && xml.Length<=MaximumMetadataCharacters);
            try
            {
                var settings=new XmlReaderSettings{DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=MaximumMetadataCharacters};
                var document=new XmlDocument{XmlResolver=null};
                using(var reader=XmlReader.Create(new StringReader(xml),settings))document.Load(reader);
                var root=document.DocumentElement;
                Require(root!=null && root.LocalName=="DataSource" && root.NamespaceURI==Namespace && new[]{"1","2"}.Contains(root.GetAttribute("version")));
                foreach(XmlAttribute attribute in root.Attributes)
                    Require(attribute.NamespaceURI=="http://www.w3.org/2000/xmlns/" || attribute.NamespaceURI=="" && new[]{"version","id","format","linkedAtUtc"}.Contains(attribute.Name));
                bool mapped=root.GetAttribute("version")=="2";Require(root.ChildNodes.Count==(mapped?2:1) && root.FirstChild is XmlElement);
                var path=(XmlElement)root.FirstChild;
                Require(path.LocalName=="Path" && path.NamespaceURI==Namespace && path.Attributes.Count==0 && path.ChildNodes.Cast<XmlNode>().All(n=>n.NodeType==XmlNodeType.Text));
                DateTime time;
                Require(DateTime.TryParseExact(root.GetAttribute("linkedAtUtc"),"o",CultureInfo.InvariantCulture,DateTimeStyles.RoundtripKind,out time));
                var link=new DataSourceLink{Id=root.GetAttribute("id"),Format=root.GetAttribute("format"),Path=path.InnerText,LinkedAtUtc=time};
                if(mapped){Require(root.LastChild is XmlElement);link.Mapping=ExcelRangeMapping.Read((XmlElement)root.LastChild,Namespace);}
                Validate(link);return link;
            }
            catch(XmlException){throw new InvalidDataException("Invalid data source link XML.");}
            catch(ArgumentException){throw new InvalidDataException("Invalid data source link path.");}
        }
        internal static IReadOnlyList<DataSourceLink> Read(PowerPoint.Presentation presentation)
        {
            // PowerPoint exposes shape CustomerData in the presentation-wide collection
            // after reopen. Those are object references, not additional registry entries.
            var attached=new HashSet<string>(StringComparer.OrdinalIgnoreCase);int visited=0;
            Action<PowerPoint.Shape,int> visit=null;visit=(shape,depth)=>{
                if(++visited>10000||depth>32)throw new InvalidDataException("Source ownership inspection exceeds supported bounds.");
                foreach(Office.CustomXMLPart part in shape.CustomerData)attached.Add(Guid.Parse(part.Id).ToString("D"));
                if(shape.Type==Office.MsoShapeType.msoGroup)for(int i=1;i<=shape.GroupItems.Count;i++)visit(shape.GroupItems[i],depth+1);
            };
            foreach(PowerPoint.Slide slide in presentation.Slides)foreach(PowerPoint.Shape shape in slide.Shapes)visit(shape,0);
            var entries=new List<Tuple<string,bool>>();
            foreach(Office.CustomXMLPart part in presentation.CustomXMLParts.SelectByNamespace(Namespace))
                entries.Add(Tuple.Create(part.XML,attached.Contains(Guid.Parse(part.Id).ToString("D"))));
            return ReadRegistry(entries);
        }
        internal static IReadOnlyList<DataSourceLink> ReadRegistry(IEnumerable<Tuple<string,bool>> entries)
        {
            var links=new List<DataSourceLink>();var ids=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach(var entry in entries){if(entry.Item2)continue;var link=Decode(entry.Item1);Require(ids.Add(link.Id));links.Add(link);if(links.Count>256)throw new InvalidDataException("Source registry exceeds 256 entries.");}
            return links;
        }
        internal static DataSourceLink Save(PowerPoint.Presentation presentation,DataSourceLink link)
        {
            string xml=Encode(link);var existing=Read(presentation);
            var duplicate=existing.SingleOrDefault(x=>x.Format==link.Format && String.Equals(x.Path,link.Path,StringComparison.OrdinalIgnoreCase) && String.Equals(x.Mapping==null?null:x.Mapping.Key,link.Mapping==null?null:link.Mapping.Key,StringComparison.Ordinal));
            if(duplicate!=null)return duplicate;
            if(existing.Count>=256)throw new InvalidOperationException("This presentation already has 256 linked data sources.");
            Office.CustomXMLPart added=null;
            try
            {
                added=presentation.CustomXMLParts.Add(xml);
                var persisted=Decode(added.XML);
                Require(persisted.Id==link.Id && persisted.Path==link.Path && persisted.Format==link.Format && persisted.LinkedAtUtc==link.LinkedAtUtc && String.Equals(persisted.Mapping==null?null:persisted.Mapping.Key,link.Mapping==null?null:link.Mapping.Key,StringComparison.Ordinal));
                return persisted;
            }
            catch { if(added!=null)added.Delete();throw; }
        }
    }
}

