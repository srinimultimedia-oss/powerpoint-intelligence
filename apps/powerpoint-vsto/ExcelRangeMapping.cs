using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
namespace PowerPointIntelligence.AddIn
{
    internal sealed class ExcelRangeMapping
    {
        internal string Worksheet,Address;
        internal int Category;
        internal int[] Series;
        internal string[] Headers;
        internal void Validate()
        {
            if(String.IsNullOrWhiteSpace(Worksheet)||Worksheet.Length>31||Worksheet.Any(c=>Char.IsControl(c)||":\\/?*[]".Contains(c)))throw new InvalidDataException("Choose a valid worksheet name.");
            int rows,columns;Dimensions(Address,out rows,out columns);
            if(Headers==null||Headers.Length!=columns||Headers.Any(h=>String.IsNullOrWhiteSpace(h)||h.Length>200||h.Any(Char.IsControl))||Headers.Distinct(StringComparer.Ordinal).Count()!=Headers.Length)throw new InvalidDataException("The range needs distinct nonempty headers, each at most 200 characters.");
            if(Category<0||Category>=columns||Series==null||Series.Length==0||Series.Any(c=>c<0||c>=columns||c==Category)||Series.Distinct().Count()!=Series.Length)throw new InvalidDataException("Choose a category and different data columns.");
        }
        internal static void Dimensions(string address,out int rows,out int columns)
        {
            var match=Regex.Match(address??"",@"^\$?([A-Z]{1,3})\$?([1-9][0-9]{0,6}):\$?([A-Z]{1,3})\$?([1-9][0-9]{0,6})$");if(!match.Success)throw new InvalidDataException("Use a single A1 rectangle, for example A1:D12.");
            Func<string,int> column=s=>s.Aggregate(0,(v,c)=>v*26+c-'A'+1);int start=column(match.Groups[1].Value),end=column(match.Groups[3].Value),first=Int32.Parse(match.Groups[2].Value,CultureInfo.InvariantCulture),last=Int32.Parse(match.Groups[4].Value,CultureInfo.InvariantCulture);rows=last-first+1;columns=end-start+1;
            if(start>16384||end>16384||first>1048576||last>1048576||rows<2||rows>50||columns<2||columns>15)throw new InvalidDataException("Choose 2–50 rows including headers and 2–15 columns inside worksheet bounds.");
        }
        internal string Key {get{Validate();return Worksheet+"\n"+Address+"\n"+Category+"\n"+String.Join(",",Series)+"\n"+String.Join("\n",Headers);}}
        internal LinkedTable Project(LinkedTable raw,bool numeric)
        {
            Validate();int rows,columns;Dimensions(Address,out rows,out columns);
            if(raw==null||raw.Rows==null||raw.Rows.Length+1!=rows||raw.Headers==null||!raw.Headers.SequenceEqual(Headers,StringComparer.Ordinal)||raw.Rows.Any(r=>r==null||r.Length!=columns))throw new InvalidDataException("Source headers or dimensions changed. Refresh stopped; explicitly reselect and confirm the mapping.");
            if(numeric)return LinkedChartData.SelectColumns(raw,Category,Series);
            var selected=new[]{Category}.Concat(Series).ToArray();return new LinkedTable {Headers=selected.Select(c=>raw.Headers[c]).ToArray(),Rows=raw.Rows.Select(r=>selected.Select(c=>r[c]).ToArray()).ToArray()}.Validate(false);
        }
        internal void Write(XmlWriter writer,string ns)
        {Validate();writer.WriteStartElement("ExcelRange",ns);writer.WriteAttributeString("worksheet",Worksheet);writer.WriteAttributeString("address",Address);writer.WriteAttributeString("category",Category.ToString(CultureInfo.InvariantCulture));writer.WriteAttributeString("series",String.Join(",",Series));foreach(var h in Headers)writer.WriteElementString("Header",ns,h);writer.WriteEndElement();}
        internal static ExcelRangeMapping Read(XmlElement element,string ns)
        {
            if(element.LocalName!="ExcelRange"||element.NamespaceURI!=ns||element.Attributes.Count!=4||element.Attributes.Cast<XmlAttribute>().Any(a=>a.NamespaceURI!=""||!new[]{"worksheet","address","category","series"}.Contains(a.Name)))throw new InvalidDataException("Invalid Excel mapping metadata.");
            int category;var tokens=element.GetAttribute("series").Split(',');int[] series=new int[tokens.Length];if(tokens.Length>14||!Int32.TryParse(element.GetAttribute("category"),NumberStyles.None,CultureInfo.InvariantCulture,out category))throw new InvalidDataException("Invalid Excel column mapping.");for(int i=0;i<series.Length;i++)if(!Int32.TryParse(tokens[i],NumberStyles.None,CultureInfo.InvariantCulture,out series[i]))throw new InvalidDataException("Invalid Excel series mapping.");
            var headers=new List<string>();foreach(XmlNode node in element.ChildNodes){var child=node as XmlElement;if(child==null||child.LocalName!="Header"||child.NamespaceURI!=ns||child.Attributes.Count!=0||child.ChildNodes.Cast<XmlNode>().Any(n=>n.NodeType!=XmlNodeType.Text))throw new InvalidDataException("Invalid Excel header mapping.");headers.Add(child.InnerText);if(headers.Count>15)throw new InvalidDataException("Too many Excel headers.");}
            var result=new ExcelRangeMapping {Worksheet=element.GetAttribute("worksheet"),Address=element.GetAttribute("address"),Category=category,Series=series,Headers=headers.ToArray()};result.Validate();return result;
        }
    }
    internal static class MappedExcelData
    {
        internal static LinkedTable Read(string path,ExcelRangeMapping mapping)
        {
            mapping.Validate();return WithWorkbook(path,book=>ReadRange(book,mapping.Worksheet,mapping.Address));
        }
        private static T WithWorkbook<T>(string path,Func<object,T> action)
        {
            var file=new FileInfo(path);if(!file.Exists||file.Length>20*1024*1024)throw new InvalidDataException("The source is missing or exceeds 20 MB. Explicitly relink it.");
            dynamic excel=null,book=null;try{excel=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true));excel.Visible=false;excel.DisplayAlerts=false;excel.AutomationSecurity=3;excel.AskToUpdateLinks=false;book=excel.Workbooks.Open(path,0,true);return action((object)book);}finally{try{if(book!=null)book.Close(false);}finally{if(excel!=null)System.Runtime.InteropServices.Marshal.ReleaseComObject((object)excel);}}
        }
        private static LinkedTable ReadRange(object workbook,string worksheet,string address)
        {
            int rowCount,columnCount;ExcelRangeMapping.Dimensions(address,out rowCount,out columnCount);dynamic book=workbook;dynamic range=book.Worksheets[worksheet].Range[address];var values=(Array)(object)range.Value2;if(values==null||values.Rank!=2||values.GetLength(0)!=rowCount||values.GetLength(1)!=columnCount)throw new InvalidDataException("The selected range dimensions are unavailable.");var all=new string[rowCount][];for(int r=0;r<rowCount;r++){all[r]=new string[columnCount];for(int c=0;c<columnCount;c++){object value=values.GetValue(values.GetLowerBound(0)+r,values.GetLowerBound(1)+c);if(value is System.Runtime.InteropServices.ErrorWrapper||value is Int32&&Convert.ToBoolean(book.Application.WorksheetFunction.IsError(range.Cells[r+1,c+1]),CultureInfo.InvariantCulture))throw new InvalidDataException("The range contains an Excel error.");all[r][c]=Convert.ToString(value,CultureInfo.InvariantCulture);}}return new LinkedTable {Headers=all[0],Rows=all.Skip(1).ToArray()};
        }
        internal static DataSourceLink Choose(PowerPoint.Application app,bool persist=true)
        {
            string path=DataSources.PickFile(DataSources.Option("excel"));if(path==null)return null;
            var mapping=WithWorkbook(path,book=>{
                dynamic native=book;var names=new List<string>();foreach(dynamic sheet in native.Worksheets){if(names.Count>=256)throw new InvalidDataException("Workbook has more than 256 worksheets.");names.Add((string)sheet.Name);}
                using(var form=new Form {Text="Link explicit Excel range",Width=530,Height=235,StartPosition=FormStartPosition.CenterParent,FormBorderStyle=FormBorderStyle.FixedDialog,MinimizeBox=false,MaximizeBox=false})
                {
                    var sheets=new ComboBox {Left=145,Top=20,Width=350,DropDownStyle=ComboBoxStyle.DropDownList};sheets.Items.AddRange(names.Cast<object>().ToArray());if(names.Count==0)throw new InvalidDataException("Workbook has no worksheets.");sheets.SelectedIndex=0;
                    var range=new TextBox {Left=145,Top=60,Width=350,Text="A1:C10",MaxLength=32};var choose=new Button {Text="Choose columns…",Left=310,Top=132,Width=185};var cancel=new Button {Text="Cancel",DialogResult=DialogResult.Cancel,Left=210,Top=132,Width=90};form.Controls.AddRange(new Control[]{new Label {Text="Worksheet",Left=20,Top=23},new Label {Text="Range with headers",Left=20,Top=63,Width=125},new Label {Text="2–50 rows; 2–15 columns. Data is read-only.",Left=20,Top=100,Width=475},sheets,range,choose,cancel});form.CancelButton=cancel;form.AcceptButton=choose;
                    ExcelRangeMapping result=null;choose.Click+=(s,e)=>{try{string address=range.Text.Trim().ToUpperInvariant();var raw=ReadRange(book,(string)sheets.SelectedItem,address);var selected=LinkedChartData.PickMappedColumns(raw);if(selected==null)return;var candidate=new ExcelRangeMapping {Worksheet=(string)sheets.SelectedItem,Address=address,Category=selected[0],Series=selected.Skip(1).ToArray(),Headers=raw.Headers};candidate.Project(raw,true);result=candidate;form.DialogResult=DialogResult.OK;}catch(Exception ex) when(ex is InvalidDataException||ex is System.Runtime.InteropServices.COMException){MessageBox.Show(form,"Cannot map that range. Check its address, headers, data types and worksheet.","Excel mapping");}};
                    return form.ShowDialog()==DialogResult.OK?result:null;
                }
            });
            if(mapping==null)return null;var link=DataSources.Create("excel",path);link.Mapping=mapping;return persist?DataSources.Save(app.ActivePresentation,link):link;
        }
    }
}
