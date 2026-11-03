using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Windows.Forms;
using Microsoft.Win32;
internal static class IntelligenceSetup
{
    internal const string BundleHash = "__BUNDLE_HASH__";
    internal static string Digest(Stream input) { using(var sha=SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(input)).Replace("-",""); }
    internal static string SafePath(string root,string relative) {
        if(String.IsNullOrWhiteSpace(relative)||relative.Contains(":")||relative.Contains("\\")||relative.StartsWith("/"))throw new InvalidDataException("Invalid package path.");
        string full=Path.GetFullPath(Path.Combine(root,relative.Replace('/',Path.DirectorySeparatorChar)));
        if(!full.StartsWith(Path.GetFullPath(root)+Path.DirectorySeparatorChar,StringComparison.OrdinalIgnoreCase))throw new InvalidDataException("Package path escapes installation directory.");return full;
    }
    internal static string OfficeArchitecture() {
        var paths=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach(var view in Environment.Is64BitOperatingSystem?new[]{RegistryView.Registry64,RegistryView.Registry32}:new[]{RegistryView.Registry32})
        using(var registry=RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,view)){
            using(var key=registry.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE")){var value=key==null?null:key.GetValue(null) as string;if(!String.IsNullOrWhiteSpace(value)&&File.Exists(value.Trim('"')))paths.Add(value.Trim('"'));}
            using(var key=registry.OpenSubKey(@"SOFTWARE\Microsoft\Office\16.0\PowerPoint\InstallRoot")){var value=key==null?null:key.GetValue("Path") as string;if(!String.IsNullOrWhiteSpace(value)&&File.Exists(Path.Combine(value,"POWERPNT.EXE")))paths.Add(Path.Combine(value,"POWERPNT.EXE"));}
        }
        var architectures=new HashSet<string>();
        foreach(string path in paths)using(var input=new BinaryReader(File.OpenRead(path))){
            if(input.ReadUInt16()!=0x5A4D)throw new InvalidDataException("Invalid PowerPoint executable header.");input.BaseStream.Position=0x3C;int offset=input.ReadInt32();if(offset<64||offset>input.BaseStream.Length-6)throw new InvalidDataException("Invalid PowerPoint executable offset.");input.BaseStream.Position=offset;if(input.ReadUInt32()!=0x4550)throw new InvalidDataException("Invalid PowerPoint PE header.");ushort machine=input.ReadUInt16();if(machine!=0x8664&&machine!=0x14c)throw new NotSupportedException("Only x86 and x64 Office are supported.");if(FileVersionInfo.GetVersionInfo(path).FileMajorPart!=16)throw new NotSupportedException("Office 2016, Office 2019 or Microsoft 365 desktop is required.");architectures.Add(machine==0x8664?"x64":"x86");
        }
        if(architectures.Count!=1)throw new InvalidOperationException("A single supported desktop PowerPoint architecture could not be identified. Repair Office registration before installing.");return architectures.Single();
    }
    private static void RunVsto(string installer,string operation,string manifest){
        var process=Process.Start(new ProcessStartInfo(installer,operation+" \""+manifest+"\""){UseShellExecute=true});process.WaitForExit();if(process.ExitCode!=0)throw new InvalidOperationException("VSTO installer returned "+process.ExitCode+".");
    }
    [STAThread] private static int Main() {
        Application.EnableVisualStyles();string phase="prerequisites";
        string home=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"PowerPointIntelligence","easySetup");Directory.CreateDirectory(home);
        string log=Path.Combine(home,"setup-"+DateTime.UtcNow.ToString("yyyyMMdd-HHmmss")+".log");
        try {
            if(Process.GetProcessesByName("POWERPNT").Length!=0)throw new InvalidOperationException("Save your presentations and close PowerPoint normally, then run setup again. Setup does not close your work.");
            using(var framework=RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,RegistryView.Registry32))using(var full=framework.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full")){
                if(full==null||Convert.ToInt32(full.GetValue("Release",0))<528040)throw new InvalidOperationException("Microsoft .NET Framework 4.8 or later is required.");
            }
            string architecture=OfficeArchitecture();
            string installer=new[]{Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles),Environment.GetEnvironmentVariable("CommonProgramFiles(x86)")}.Where(x=>!String.IsNullOrEmpty(x)).Select(x=>Path.Combine(x,"microsoft shared","VSTO","10.0","VSTOInstaller.exe")).FirstOrDefault(File.Exists);
            if(installer==null)throw new FileNotFoundException("Microsoft Visual Studio 2010 Tools for Office Runtime is required. Install the Microsoft prerequisite, then retry.");
            if(MessageBox.Show("Install NexWave for "+architecture+" Office for this Windows user?\n\nThe included third-party templates are for your licensed computers only. A previous easySetup installation will be upgraded using the normal VSTO prompts.","NexWave Setup",MessageBoxButtons.OKCancel,MessageBoxIcon.Information)!=DialogResult.OK)return 2;
            phase="verify package";string destination=Path.Combine(home,"versions",BundleHash);
            using(var resource=Assembly.GetExecutingAssembly().GetManifestResourceStream("IntelligencePayload.zip")){
                if(resource==null||Digest(resource)!=BundleHash)throw new InvalidDataException("Embedded package checksum mismatch.");resource.Position=0;
                using(var zip=new ZipArchive(resource,ZipArchiveMode.Read)){
                    if(zip.Entries.Count>5000||zip.Entries.Sum(e=>e.Length)>536870912)throw new InvalidDataException("Package exceeds extraction bounds.");
                    var inventory=zip.GetEntry("payload.sha256");if(inventory==null||inventory.Length>2000000)throw new InvalidDataException("Missing package inventory.");
                    var expected=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
                    using(var reader=new StreamReader(inventory.Open())){string line;while((line=reader.ReadLine())!=null){var fields=line.Split('\t');if(fields.Length!=2||fields[0].Length!=64||!fields[0].All(Uri.IsHexDigit)||expected.ContainsKey(fields[1]))throw new InvalidDataException("Invalid package inventory.");SafePath(destination,fields[1]);expected.Add(fields[1],fields[0]);}}
                    if(zip.Entries.Count!=expected.Count+1)throw new InvalidDataException("Unexpected package entries.");
                    foreach(var item in expected){var entry=zip.GetEntry(item.Key);if(entry==null)throw new InvalidDataException("Missing packaged file.");using(var input=entry.Open())if(Digest(input)!=item.Value)throw new InvalidDataException("Packaged file checksum mismatch.");}
                    phase="extract";Directory.CreateDirectory(destination);
                    foreach(var item in expected){string target=SafePath(destination,item.Key);if(File.Exists(target)){using(var input=File.OpenRead(target))if(Digest(input)!=item.Value)throw new InvalidDataException("Existing version was modified; it was not overwritten.");continue;}Directory.CreateDirectory(Path.GetDirectoryName(target));using(var input=zip.GetEntry(item.Key).Open())using(var output=new FileStream(target,FileMode.CreateNew,FileAccess.Write))input.CopyTo(output);}
                }
            }
            phase="VSTO installation";string manifest=Path.Combine(destination,architecture,"PowerPointIntelligence.vsto");
            string receipt=Path.Combine(home,"installed-manifest.txt"),previous=File.Exists(receipt)?File.ReadAllText(receipt).Trim():null;
            if(previous!=null&&(!Path.GetFullPath(previous).StartsWith(Path.Combine(home,"versions")+Path.DirectorySeparatorChar,StringComparison.OrdinalIgnoreCase)||Path.GetFileName(previous)!="PowerPointIntelligence.vsto"||!File.Exists(previous)))throw new InvalidDataException("Previous setup receipt is unavailable or invalid. Repair/uninstall the previous package through Windows before retrying.");
            bool changing=previous!=null&&!String.Equals(previous,manifest,StringComparison.OrdinalIgnoreCase);
            if(changing)RunVsto(installer,"/Uninstall",previous);
            try{RunVsto(installer,"/Install",manifest);File.WriteAllText(receipt,manifest);}
            catch{if(changing)try{RunVsto(installer,"/Install",previous);}catch(Exception rollback){File.AppendAllText(log,"Rollback failed: "+rollback.GetType().Name+Environment.NewLine);}throw;}
            File.AppendAllText(log,"PASS VSTO installation; architecture="+architecture+"; bundle="+BundleHash+Environment.NewLine);
            MessageBox.Show("Installation completed. Open PowerPoint and check the NexWave ribbon.\n\nRuntime functionality still requires validation.\nUninstall through Windows Apps / Programs and Features.\n\nLog: "+log,"NexWave Setup");return 0;
        }catch(Exception error){File.AppendAllText(log,"FAIL phase="+phase+"; exception="+error.GetType().Name+"; HRESULT=0x"+error.HResult.ToString("X8")+Environment.NewLine);MessageBox.Show(error.Message+"\n\nLog: "+log,"NexWave Setup",MessageBoxButtons.OK,MessageBoxIcon.Error);return 1;}
    }
}
