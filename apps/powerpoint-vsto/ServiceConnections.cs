using PowerPointIntelligence.Security;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
namespace PowerPointIntelligence.AddIn
{
    internal enum ConnectionService { AzureOpenAI, OnPremDataService, AssetProvider }
    internal enum ConnectionAuthentication { Unselected, MicrosoftEntraId, WindowsIntegrated, RuntimeCredential }
    internal sealed class ServiceConnectionProfile
    {
        internal ConnectionService Service;
        internal string Endpoint="";
        internal string Deployment="";
        internal ConnectionAuthentication Authentication;
        internal ServiceConnectionProfile Copy(){return new ServiceConnectionProfile {Service=Service,Endpoint=Endpoint,Deployment=Deployment,Authentication=Authentication};}
    }
    internal sealed class ServiceConnectionSettings
    {
        internal int Version=1;
        internal List<ServiceConnectionProfile> Profiles=new List<ServiceConnectionProfile>();
        internal static ServiceConnectionSettings Empty()
        {return new ServiceConnectionSettings {Profiles=Enum.GetValues(typeof(ConnectionService)).Cast<ConnectionService>().Select(s=>new ServiceConnectionProfile {Service=s}).ToList()};}
    }
    // Configuration preparation only. No HTTP client, credential store, discovery,
    // authentication attempt or simulated response is part of this component.
    internal static class ServiceConnections
    {
        internal const string Namespace="urn:powerpoint-intelligence:service-connections:v1";
        internal const int MaxCharacters=16384,MaxFileBytes=65536;
        internal static string DefaultPath {get{return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"PowerPointIntelligence","service-connections.v1.xml");}}
        internal static ConnectionAuthentication[] AllowedAuthentication(ConnectionService service)
        {return service==ConnectionService.OnPremDataService?new[]{ConnectionAuthentication.Unselected,ConnectionAuthentication.MicrosoftEntraId,ConnectionAuthentication.WindowsIntegrated,ConnectionAuthentication.RuntimeCredential}:new[]{ConnectionAuthentication.Unselected,ConnectionAuthentication.MicrosoftEntraId,ConnectionAuthentication.RuntimeCredential};}
        internal static void Validate(ServiceConnectionSettings settings)
        {
            if(settings==null)throw new ArgumentNullException(nameof(settings));
            if(settings.Version!=1||settings.Profiles==null||settings.Profiles.Count!=3)throw new InvalidDataException("Unknown configuration version or service count.");
            if(settings.Profiles.Any(p=>p==null)||settings.Profiles.Select(p=>p.Service).Distinct().Count()!=3)throw new InvalidDataException("Each service must appear exactly once.");
            foreach(var profile in settings.Profiles)ValidateProfile(profile);
        }
        private static void ValidateProfile(ServiceConnectionProfile profile)
        {
            if(profile==null||!Enum.IsDefined(typeof(ConnectionService),profile.Service)||!Enum.IsDefined(typeof(ConnectionAuthentication),profile.Authentication)||!AllowedAuthentication(profile.Service).Contains(profile.Authentication))throw new InvalidDataException("Unsupported service or authentication mode.");
            if(profile.Endpoint==null||profile.Endpoint.Length>2048||profile.Endpoint.Any(c=>Char.IsControl(c)||Char.IsWhiteSpace(c)))throw new InvalidDataException("Endpoint must be a bounded base HTTPS URI without whitespace.");
            if(profile.Endpoint.Length>0){Uri uri;if(!Uri.TryCreate(profile.Endpoint,UriKind.Absolute,out uri)||uri.Scheme!=Uri.UriSchemeHttps||String.IsNullOrEmpty(uri.Host)||uri.UserInfo.Length>0||uri.Query.Length>0||uri.Fragment.Length>0)throw new InvalidDataException("Use an absolute HTTPS endpoint without credentials, query parameters or fragments.");}
            if(profile.Deployment==null||profile.Deployment.Length>128||profile.Deployment.Any(Char.IsControl)||profile.Deployment!=profile.Deployment.Trim())throw new InvalidDataException("Deployment identifier must be at most 128 characters without surrounding whitespace or control characters.");
            if(profile.Service!=ConnectionService.AzureOpenAI&&profile.Deployment.Length>0)throw new InvalidDataException("Deployment identifiers apply only to Azure OpenAI.");
        }
        internal static string Status(ServiceConnectionProfile profile)
        {
            ValidateProfile(profile);bool complete=profile.Endpoint.Length>0&&profile.Authentication!=ConnectionAuthentication.Unselected&&(profile.Service!=ConnectionService.AzureOpenAI||profile.Deployment.Length>0);
            return complete?"Configured, not connectivity-tested":"Not configured";
        }
        internal static string Encode(ServiceConnectionSettings settings)
        {
            Validate(settings);var document=new XmlDocument {XmlResolver=null};var root=document.CreateElement("Connections",Namespace);document.AppendChild(root);root.SetAttribute("version","1");
            foreach(var profile in settings.Profiles.OrderBy(p=>p.Service)){
                var node=document.CreateElement("Profile",Namespace);node.SetAttribute("service",profile.Service.ToString());node.SetAttribute("authentication",profile.Authentication.ToString());root.AppendChild(node);
                foreach(var pair in new[]{Tuple.Create("Endpoint",profile.Endpoint),Tuple.Create("Deployment",profile.Deployment)}){var value=document.CreateElement(pair.Item1,Namespace);value.InnerText=pair.Item2;node.AppendChild(value);}
            }
            string xml=document.OuterXml;if(xml.Length>MaxCharacters)throw new InvalidDataException("Configuration exceeds size limit.");return xml;
        }
        private static void Attributes(XmlElement element,params string[] allowed)
        {
            foreach(XmlAttribute attribute in element.Attributes){if(attribute.Name=="xmlns"&&attribute.Value==Namespace)continue;if(attribute.NamespaceURI.Length!=0||!allowed.Contains(attribute.Name))throw new InvalidDataException("Unknown configuration attribute.");}
        }
        private static XmlElement[] Elements(XmlElement element)
        {
            foreach(XmlNode node in element.ChildNodes)if(node.NodeType!=XmlNodeType.Element&&node.NodeType!=XmlNodeType.Whitespace&&node.NodeType!=XmlNodeType.SignificantWhitespace)throw new InvalidDataException("Unexpected configuration content.");return element.ChildNodes.OfType<XmlElement>().ToArray();
        }
        internal static ServiceConnectionSettings Decode(string xml)
        {
            if(xml==null||xml.Length>MaxCharacters)throw new InvalidDataException("Configuration exceeds size limit.");var document=new XmlDocument {XmlResolver=null};
            using(var reader=XmlReader.Create(new StringReader(xml),new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=MaxCharacters,IgnoreComments=true,IgnoreProcessingInstructions=true}))document.Load(reader);
            var root=document.DocumentElement;if(root==null||root.LocalName!="Connections"||root.NamespaceURI!=Namespace||root.GetAttribute("version")!="1")throw new InvalidDataException("Unknown configuration version or namespace.");Attributes(root,"version");var settings=new ServiceConnectionSettings();
            foreach(var node in Elements(root)){
                if(node.LocalName!="Profile"||node.NamespaceURI!=Namespace)throw new InvalidDataException("Unknown service configuration element.");Attributes(node,"service","authentication");ConnectionService service;ConnectionAuthentication auth;
                if(!Enum.TryParse(node.GetAttribute("service"),false,out service)||!Enum.IsDefined(typeof(ConnectionService),service)||node.GetAttribute("service")!=service.ToString()||!Enum.TryParse(node.GetAttribute("authentication"),false,out auth)||!Enum.IsDefined(typeof(ConnectionAuthentication),auth)||node.GetAttribute("authentication")!=auth.ToString())throw new InvalidDataException("Unknown service or authentication mode.");
                var fields=Elements(node);if(fields.Length!=2||fields.Count(e=>e.LocalName=="Endpoint"&&e.NamespaceURI==Namespace)!=1||fields.Count(e=>e.LocalName=="Deployment"&&e.NamespaceURI==Namespace)!=1)throw new InvalidDataException("Expected endpoint and deployment fields only.");
                foreach(var field in fields){Attributes(field);if(field.ChildNodes.OfType<XmlElement>().Any())throw new InvalidDataException("Configuration fields must be plain text.");}
                settings.Profiles.Add(new ServiceConnectionProfile {Service=service,Authentication=auth,Endpoint=fields.Single(e=>e.LocalName=="Endpoint").InnerText,Deployment=fields.Single(e=>e.LocalName=="Deployment").InnerText});
            }
            Validate(settings);return settings;
        }
        internal static ServiceConnectionSettings Load(string path)
        {
            path=PathPolicy.ConfigFile(path);if(!File.Exists(path))return ServiceConnectionSettings.Empty();if(new FileInfo(path).Length>MaxFileBytes)throw new InvalidDataException("Saved configuration exceeds byte limit.");return Decode(File.ReadAllText(path,new UTF8Encoding(false,true)));
        }
        internal static void Save(string path,ServiceConnectionSettings settings)
        {
            string xml=Encode(settings);string target=PathPolicy.ConfigFile(path),directory=Path.GetDirectoryName(target);Directory.CreateDirectory(directory);string temp=PathPolicy.Within(directory,".sc-"+Guid.NewGuid().ToString("N")+".tmp");
            try{File.WriteAllText(temp,xml,new UTF8Encoding(false,true));if(File.Exists(target))File.Replace(temp,target,null);else File.Move(temp,target);}finally{if(File.Exists(temp))File.Delete(temp);}
        }
        internal static void ShowConfiguration()
        {
            ServiceConnectionSettings loaded;string loadStatus="No credentials are stored here. Use approved nonsecret endpoints and identifiers only. Nothing is connectivity-tested.";
            try{loaded=Load(DefaultPath);}catch(Exception ex) when(ex is IOException||ex is UnauthorizedAccessException||ex is InvalidDataException||ex is XmlException||ex is DecoderFallbackException){loaded=ServiceConnectionSettings.Empty();loadStatus="Saved configuration could not be read. Save will replace it; Cancel preserves it. No credentials or connection tests are supported.";}
            using(var form=new Form {Text="Service connections",Width=760,Height=650,MinimumSize=new System.Drawing.Size(690,580),StartPosition=FormStartPosition.CenterScreen})
            using(var body=new FlowLayoutPanel {Dock=DockStyle.Fill,AutoScroll=true,FlowDirection=FlowDirection.TopDown,WrapContents=false,Padding=new Padding(12)})
            using(var footer=new FlowLayoutPanel {Dock=DockStyle.Bottom,Height=48,FlowDirection=FlowDirection.RightToLeft,Padding=new Padding(8)})
            using(var save=new Button {Text="Save configuration",AutoSize=true})
            using(var cancel=new Button {Text="Cancel",DialogResult=DialogResult.Cancel,AutoSize=true})
            {
                var message=new Label {Text=loadStatus,Width=690,Height=58};body.Controls.Add(message);var editors=new List<Tuple<ServiceConnectionProfile,TextBox,TextBox,ComboBox,Label>>();
                foreach(var profile in loaded.Profiles.OrderBy(p=>p.Service)){
                    var group=new GroupBox {Text=Label(profile.Service),Width=690,Height=profile.Service==ConnectionService.AzureOpenAI?152:126};var grid=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=2,RowCount=4,Padding=new Padding(10)};grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,120));grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
                    var endpoint=new TextBox {Text=profile.Endpoint,Dock=DockStyle.Fill,MaxLength=2048,AccessibleName=Label(profile.Service)+" endpoint"};var deployment=new TextBox {Text=profile.Deployment,Dock=DockStyle.Fill,MaxLength=128,AccessibleName="Azure OpenAI deployment identifier"};var auth=new ComboBox {DropDownStyle=ComboBoxStyle.DropDownList,Dock=DockStyle.Fill,AccessibleName=Label(profile.Service)+" authentication mode"};auth.Items.AddRange(AllowedAuthentication(profile.Service).Cast<object>().ToArray());auth.SelectedItem=profile.Authentication;var state=new Label {AutoSize=true,Text=Status(profile)};
                    grid.Controls.Add(new Label {Text="HTTPS endpoint",AutoSize=true},0,0);grid.Controls.Add(endpoint,1,0);grid.Controls.Add(new Label {Text="Authentication",AutoSize=true},0,1);grid.Controls.Add(auth,1,1);if(profile.Service==ConnectionService.AzureOpenAI){grid.Controls.Add(new Label {Text="Deployment ID",AutoSize=true},0,2);grid.Controls.Add(deployment,1,2);}grid.Controls.Add(state,1,3);group.Controls.Add(grid);body.Controls.Add(group);var row=Tuple.Create(profile,endpoint,deployment,auth,state);editors.Add(row);
                    Action update=()=>{try{state.Text=Status(new ServiceConnectionProfile {Service=profile.Service,Endpoint=endpoint.Text,Deployment=deployment.Text,Authentication=(ConnectionAuthentication)auth.SelectedItem});}catch(InvalidDataException){state.Text="Not configured — invalid endpoint or identifier";}};endpoint.TextChanged+=(s,e)=>update();deployment.TextChanged+=(s,e)=>update();auth.SelectedIndexChanged+=(s,e)=>update();
                }
                save.Click+=(s,e)=>{try{var settings=new ServiceConnectionSettings {Profiles=editors.Select(row=>new ServiceConnectionProfile {Service=row.Item1.Service,Endpoint=row.Item2.Text,Deployment=row.Item3.Text,Authentication=(ConnectionAuthentication)row.Item4.SelectedItem}).ToList()};Save(DefaultPath,settings);form.DialogResult=DialogResult.OK;form.Close();}catch(Exception ex) when(ex is IOException||ex is UnauthorizedAccessException||ex is InvalidDataException){message.Text="Configuration not saved. "+ex.Message;}};
                footer.Controls.Add(cancel);footer.Controls.Add(save);form.Controls.Add(body);form.Controls.Add(footer);form.AcceptButton=save;form.CancelButton=cancel;form.ShowDialog();
            }
        }
        private static string Label(ConnectionService service)
        {return service==ConnectionService.AzureOpenAI?"Azure OpenAI":service==ConnectionService.OnPremDataService?"On-premises data service":"Asset provider";}
    }
}

