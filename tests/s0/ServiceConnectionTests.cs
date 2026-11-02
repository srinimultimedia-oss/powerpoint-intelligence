using System;
using System.IO;
using System.Linq;
using System.Xml;
using PowerPointIntelligence.AddIn;
internal static class ServiceConnectionTests
{
    private static int passed;
    [STAThread] public static int Main(string[] args){try{Run(args.Length==0?null:args[0]);return 0;}catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
    private static void Check(bool condition,string name){if(!condition)throw new Exception("Service connection: "+name);passed++;}
    private static void Reject(Action action,string name){bool rejected=false;try{action();}catch(Exception ex) when(ex is InvalidDataException||ex is XmlException||ex is ArgumentException){rejected=true;}Check(rejected,name);}
    private static ServiceConnectionSettings Valid()
    {
        var settings=ServiceConnectionSettings.Empty();var azure=settings.Profiles.Single(p=>p.Service==ConnectionService.AzureOpenAI);azure.Endpoint="https://approved-example.invalid/openai";azure.Deployment="reviewed-deployment";azure.Authentication=ConnectionAuthentication.MicrosoftEntraId;var data=settings.Profiles.Single(p=>p.Service==ConnectionService.OnPremDataService);data.Endpoint="https://intranet.example.invalid/api";data.Authentication=ConnectionAuthentication.WindowsIntegrated;return settings;
    }
    internal static void Run(string temporaryParent=null)
    {
        passed=0;var empty=ServiceConnectionSettings.Empty();Check(empty.Profiles.Count==3&&empty.Profiles.All(p=>ServiceConnections.Status(p)=="Not configured"),"three services default unconfigured");
        var settings=Valid();Check(ServiceConnections.Status(settings.Profiles[0])=="Configured, not connectivity-tested","metadata configuration never claims live connectivity");
        string xml=ServiceConnections.Encode(settings);var decoded=ServiceConnections.Decode(xml);Check(ServiceConnections.Encode(decoded)==xml,"versioned canonical round trip");Check(decoded.Profiles.Single(p=>p.Service==ConnectionService.AssetProvider).Endpoint=="","unconfigured services retained");
        var azure=settings.Profiles.Single(p=>p.Service==ConnectionService.AzureOpenAI);azure.Deployment="";Check(ServiceConnections.Status(azure)=="Not configured","Azure requires a deployment identifier");azure.Deployment="approved-name";azure.Authentication=ConnectionAuthentication.Unselected;Check(ServiceConnections.Status(azure)=="Not configured","auth mode must be selected");
        foreach(string endpoint in new[]{"http://example.invalid","file:///C:/private.txt","relative/path","https://user:password@example.invalid","https://example.invalid/?api-key=secret","https://example.invalid/#token","https://example.invalid/a b","https://example.invalid/\n"}){
            var bad=Valid();bad.Profiles[0].Endpoint=endpoint;Reject(()=>ServiceConnections.Encode(bad),"unsafe endpoint rejected");
        }
        var invalid=Valid();invalid.Profiles[0].Endpoint=new String('a',2049);Reject(()=>ServiceConnections.Encode(invalid),"endpoint bound");invalid=Valid();invalid.Profiles[0].Deployment=new String('x',129);Reject(()=>ServiceConnections.Encode(invalid),"deployment bound");invalid=Valid();invalid.Profiles[0].Deployment="secret\nvalue";Reject(()=>ServiceConnections.Encode(invalid),"control characters rejected");
        invalid=Valid();invalid.Profiles[2].Deployment="not-an-azure-profile";Reject(()=>ServiceConnections.Encode(invalid),"deployment restricted to Azure profile");invalid=Valid();invalid.Profiles[0].Authentication=ConnectionAuthentication.WindowsIntegrated;Reject(()=>ServiceConnections.Encode(invalid),"auth modes scoped to service");invalid=Valid();invalid.Profiles[0].Authentication=(ConnectionAuthentication)99;Reject(()=>ServiceConnections.Encode(invalid),"unknown auth enum");invalid=Valid();invalid.Profiles[0].Service=(ConnectionService)99;Reject(()=>ServiceConnections.Encode(invalid),"unknown service enum");
        invalid=Valid();invalid.Version=2;Reject(()=>ServiceConnections.Encode(invalid),"unknown object version");invalid=Valid();invalid.Profiles.RemoveAt(0);Reject(()=>ServiceConnections.Encode(invalid),"missing service");invalid=Valid();invalid.Profiles[1]=invalid.Profiles[0].Copy();Reject(()=>ServiceConnections.Encode(invalid),"duplicate service");
        Reject(()=>ServiceConnections.Decode(xml.Replace("version=\"1\"","version=\"2\"")),"unknown XML version");Reject(()=>ServiceConnections.Decode(xml.Replace(ServiceConnections.Namespace,"urn:unknown")),"unknown root namespace");Reject(()=>ServiceConnections.Decode(xml.Replace("<Profile ","<Profile apiKey=\"must-not-store\" ")),"secret/unknown attribute rejected");Reject(()=>ServiceConnections.Decode(xml.Replace("<Endpoint>","<ApiKey>")),"secret/unknown element rejected");Reject(()=>ServiceConnections.Decode(xml.Replace("service=\"AzureOpenAI\"","service=\"0\"")),"numeric service identifiers rejected");Reject(()=>ServiceConnections.Decode(xml.Replace("authentication=\"MicrosoftEntraId\"","authentication=\"1\"")),"numeric auth identifiers rejected");
        Reject(()=>ServiceConnections.Decode("<!DOCTYPE Connections [<!ENTITY x SYSTEM 'file:///never-read'>]>"+xml),"DTD/external entities prohibited");Reject(()=>ServiceConnections.Decode(new String('x',ServiceConnections.MaxCharacters+1)),"XML character bound");Reject(()=>ServiceConnections.Decode(xml.Replace("<Endpoint>","<Endpoint><Nested/>")),"nested endpoint data rejected");
        string root=Path.Combine(temporaryParent??Path.GetTempPath(),"svc-"+Guid.NewGuid().ToString("N").Substring(0,8));string file=Path.Combine(root,"connections.xml");
        Check(ServiceConnections.Load(file).Profiles.All(p=>ServiceConnections.Status(p)=="Not configured")&&!Directory.Exists(root),"read/cancel path creates no file or directory");ServiceConnections.Save(file,Valid());string first=File.ReadAllText(file);Check(ServiceConnections.Encode(ServiceConnections.Load(file))==first,"file round trip");var changed=Valid();changed.Profiles[0].Deployment="second-approved-deployment";ServiceConnections.Save(file,changed);Check(ServiceConnections.Load(file).Profiles[0].Deployment=="second-approved-deployment","existing config atomically replaced");Check(Directory.GetFiles(root).Length==1,"temporary files cleaned after save");
        var badSave=Valid();badSave.Profiles[0].Endpoint="https://example.invalid/?token=secret";string preserved=File.ReadAllText(file);Reject(()=>ServiceConnections.Save(file,badSave),"invalid save rejected");Check(File.ReadAllText(file)==preserved,"invalid save preserves previous config");
        string tooLarge=Path.Combine(root,"oversize.xml");File.WriteAllBytes(tooLarge,new byte[ServiceConnections.MaxFileBytes+1]);Reject(()=>ServiceConnections.Load(tooLarge),"saved byte bound");
        Check(!xml.Contains("ApiKey")&&!xml.Contains("Password")&&!xml.Contains("CredentialValue"),"schema has no secret storage fields");Console.WriteLine("PASS "+passed+" service-profile codec/validation/persistence assertions; no network or UI operation invoked.");
    }
}

