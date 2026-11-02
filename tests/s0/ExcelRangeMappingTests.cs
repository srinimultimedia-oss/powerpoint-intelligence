using System;
using System.IO;
using System.Linq;
using PowerPointIntelligence.AddIn;
internal static class ExcelRangeMappingTests
{
    private static ExcelRangeMapping Valid(){return new ExcelRangeMapping {Worksheet="Sales & planning",Address="B3:E6",Category=1,Series=new[]{0,3},Headers=new[]{"Actual","Period","Notes","Plan"}};}
    internal static void Run()
    {
        int passed=0;Action<bool> check=v=>{if(!v)throw new Exception("Excel mapping assertion failed.");passed++;};Action<Action> reject=action=>{try{action();}catch(InvalidDataException){passed++;return;}throw new Exception("Invalid Excel mapping accepted.");};
        var mapping=Valid();mapping.Validate();var raw=new LinkedTable {Headers=mapping.Headers,Rows=new[]{new[]{"10","Q1","text","12"},new[]{"20","Q2","more","22"},new[]{"30","Q3","extra","32"}}};var table=mapping.Project(raw,true);check(table.Headers.SequenceEqual(new[]{"Period","Actual","Plan"})&&table.Rows[2][2]=="32");
        var link=new DataSourceLink {Id=Guid.NewGuid().ToString("D"),Format="excel",Path=Path.GetFullPath("fixture.xlsx"),LinkedAtUtc=DateTime.UtcNow,Mapping=mapping};string xml=DataSources.Encode(link);var copy=DataSources.Decode(xml);check(copy.Mapping.Key==mapping.Key&&DataSources.Encode(copy)==xml);check(xml.Contains("version=\"2\"")&&xml.Contains("&amp;"));
        var legacy=new DataSourceLink {Id=link.Id,Format="excel",Path=link.Path,LinkedAtUtc=link.LinkedAtUtc};check(DataSources.Decode(DataSources.Encode(legacy)).Mapping==null);check(!DataLinkInventory.Same(link,legacy));
        var changed=Valid();changed.Series=new[]{3};check(!DataLinkInventory.Same(link,new DataSourceLink {Id=link.Id,Format=link.Format,Path=link.Path,LinkedAtUtc=link.LinkedAtUtc,Mapping=changed}));
        foreach(string address in new[]{"A1","A1:B1","A1:P3","A1:B51","A0:B2","B2:A3","A1:XFE3","A1048576:B1048577","A1:B3,C1:D3","[other.xlsx]Sheet!A1:B3","A:A","a1:b3"}){var bad=Valid();bad.Address=address;reject(()=>bad.Validate());}
        foreach(string name in new[]{"","Bad/Name",new String('x',32),"Bad\nName"}){var bad=Valid();bad.Worksheet=name;reject(()=>bad.Validate());}
        var badHeaders=Valid();badHeaders.Headers=new[]{"x","x","y","z"};reject(()=>badHeaders.Validate());var controlHeader=Valid();controlHeader.Headers[0]="Line\nBreak";reject(()=>controlHeader.Validate());var badColumns=Valid();badColumns.Series=new[]{1};reject(()=>badColumns.Validate());badColumns=Valid();badColumns.Series=new[]{0,0};reject(()=>badColumns.Validate());
        var renamed=new LinkedTable {Headers=new[]{"Changed","Period","Notes","Plan"},Rows=raw.Rows};reject(()=>mapping.Project(renamed,true));reject(()=>mapping.Project(new LinkedTable {Headers=raw.Headers,Rows=raw.Rows.Take(2).ToArray()},true));
        reject(()=>DataSources.Decode(xml.Replace("category=\"1\"","category=\"-1\"")));reject(()=>DataSources.Decode(xml.Replace("<Header>","<Header extra=\"x\">")));reject(()=>DataSources.Decode(xml.Replace("version=\"2\"","version=\"3\"")));reject(()=>DataSources.Decode("<!DOCTYPE x [<!ENTITY ext SYSTEM 'file:///never'>]>"+xml));
        Console.WriteLine("PASS "+passed+" explicit Excel range/mapping/version/shape-change assertions. Excel access and native persistence not executed.");
    }
}
