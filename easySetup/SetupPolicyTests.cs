using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Collections.Generic;
using System.Linq;
internal static class SetupPolicyTests {
 private static int count;
 private static void Check(bool value){if(!value)throw new Exception("Setup policy assertion failed.");count++;}
 private static void Reject(Action action){try{action();}catch(InvalidDataException){count++;return;}throw new Exception("Unsafe package accepted.");}
 private static void RejectPath(string root,string name){Reject(()=>IntelligenceSetup.SafePath(root,name));}
 private static string Hash(byte[] bytes){using(var input=new MemoryStream(bytes))return IntelligenceSetup.Digest(input);}
 private static void Add(ZipArchive zip,string name,byte[] bytes){using(var output=zip.CreateEntry(name).Open())output.Write(bytes,0,bytes.Length);}
 private static MemoryStream Package(string separator,Action<ZipArchive> extra=null,string inventoryOverride=null){
  var data=new[]{new KeyValuePair<string,byte[]>("x86/PowerPointIntelligence.vsto",new byte[]{1,2,3}),new KeyValuePair<string,byte[]>("x64/templates/Nested folder/Test.potx",new byte[]{4,5,6}),new KeyValuePair<string,byte[]>("README.md",new byte[]{7})};
  var stream=new MemoryStream();
  using(var zip=new ZipArchive(stream,ZipArchiveMode.Create,true)){
   foreach(var item in data)Add(zip,item.Key.Replace("/",separator),item.Value);
   string inventory=String.Join("\n",data.Select(x=>Hash(x.Value)+"\t"+x.Key));
   Add(zip,"payload.sha256",Encoding.UTF8.GetBytes(inventoryOverride??inventory));if(extra!=null)extra(zip);
  }
  stream.Position=0;return stream;
 }
 private static void Verify(string root,string separator){using(var stream=Package(separator))using(var zip=new ZipArchive(stream,ZipArchiveMode.Read)){var files=IntelligenceSetup.VerifyPackage(zip,root);Check(files.Count==3&&files.ContainsKey("x86/PowerPointIntelligence.vsto")&&files.ContainsKey("x64/templates/Nested folder/Test.potx"));}}
 private static void Invalid(string root,Action<ZipArchive> extra=null,string inventory=null){using(var stream=Package("/",extra,inventory))using(var zip=new ZipArchive(stream,ZipArchiveMode.Read))Reject(()=>IntelligenceSetup.VerifyPackage(zip,root));}
 public static int Main(string[] args){
  string root=Path.Combine(Path.GetTempPath(),"ppi-setup-policy-"+Guid.NewGuid().ToString("N"));
  if(args.Length==2&&args[0]=="--verify-bundle"){
   // Verification is read-only. Use a short virtual destination so this test
   // does not depend on the agent sandbox TEMP prefix; GUI uses its actual destination.
   root=Path.Combine(Path.GetPathRoot(Path.GetTempPath()),"ppi-payload-verification");
   using(var input=File.OpenRead(args[1]))using(var zip=new ZipArchive(input,ZipArchiveMode.Read)){
    var files=IntelligenceSetup.VerifyPackage(zip,root);
    if(!files.Keys.Any(x=>x.StartsWith("x86/",StringComparison.Ordinal))||!files.Keys.Any(x=>x.StartsWith("x64/",StringComparison.Ordinal)))throw new InvalidDataException("Both Office payloads are required.");
    Console.WriteLine("PASS actual bundle: "+files.Count+" inventory files verified; raw backslash entries="+zip.Entries.Count(x=>x.FullName.Contains("\\"))+"; legacy forward lookup misses="+files.Keys.Count(x=>zip.GetEntry(x)==null)+"; both x86/x64 payloads present. No installation or extraction performed.");return 0;
   }
  }
  Check(IntelligenceSetup.SafePath(root,"x64/PowerPointIntelligence.vsto")==Path.Combine(root,"x64","PowerPointIntelligence.vsto"));
  foreach(string name in new[]{"../escape.exe","/absolute","C:/outside","x64/file:stream","x64\\file","","x64/../file","x64/./file","x64//file","x64/file.","x64/file "})RejectPath(root,name);
  Check(Hash(new byte[]{1,2,3})=="039058C6F2C0CB492C533B0A4D14EF77CC0F78ABCCCED5287D84A1A2011CFB81");
  Verify(root,"/");Verify(root,"\\");
  Invalid(root,z=>Add(z,"x86\\PowerPointIntelligence.vsto",new byte[]{1,2,3}));
  Invalid(root,z=>Add(z,"X86/POWERPOINTINTELLIGENCE.VSTO",new byte[]{1,2,3}));
  Invalid(root,z=>Add(z,"unexpected.txt",new byte[]{9}));
  Invalid(root,z=>Add(z,"..\\escape.exe",new byte[]{9}));
  Invalid(root,z=>Add(z,"C:\\outside",new byte[]{9}));
  Invalid(root,null,Hash(new byte[]{1,2,3})+"\tx86/missing.vsto\n"+Hash(new byte[]{4,5,6})+"\tx64/templates/Nested folder/Test.potx\n"+Hash(new byte[]{7})+"\tREADME.md");
  Invalid(root,null,new string('0',64)+"\tx86/PowerPointIntelligence.vsto\n"+Hash(new byte[]{4,5,6})+"\tx64/templates/Nested folder/Test.potx\n"+Hash(new byte[]{7})+"\tREADME.md");
  Invalid(root,null,Hash(new byte[]{1,2,3})+"\tx86\\PowerPointIntelligence.vsto");
  Invalid(root,null,"invalid\tREADME.md");
  Invalid(root,null,Hash(new byte[]{7})+"\tREADME.md\n"+Hash(new byte[]{7})+"\tREADME.md");
  Invalid(root,z=>Add(z,"folder/",new byte[0]));
  // Exercise the same returned entry objects used by GUI extraction, including
  // existing-file verification. Never launch VSTO or touch a user installation.
  try{
   using(var stream=Package("\\"))using(var zip=new ZipArchive(stream,ZipArchiveMode.Read)){
    var files=IntelligenceSetup.VerifyPackage(zip,root);Check(!Directory.Exists(root));
    IntelligenceSetup.ExtractPackage(files,root);
    Check(File.ReadAllBytes(Path.Combine(root,"x86","PowerPointIntelligence.vsto")).SequenceEqual(new byte[]{1,2,3}));
    Check(File.ReadAllBytes(Path.Combine(root,"x64","templates","Nested folder","Test.potx")).SequenceEqual(new byte[]{4,5,6}));
    IntelligenceSetup.ExtractPackage(files,root);Check(File.ReadAllBytes(Path.Combine(root,"README.md")).SequenceEqual(new byte[]{7}));
    File.WriteAllBytes(Path.Combine(root,"README.md"),new byte[]{99});Reject(()=>IntelligenceSetup.ExtractPackage(files,root));
    Check(File.ReadAllBytes(Path.Combine(root,"README.md")).SequenceEqual(new byte[]{99}));
   }
  }finally{if(Directory.Exists(root)){// Exact owned fixture path, never recursive cleanup of a user path.
   foreach(var file in Directory.GetFiles(root,"*",SearchOption.AllDirectories))File.Delete(file);
   foreach(var dir in Directory.GetDirectories(root,"*",SearchOption.AllDirectories).OrderByDescending(x=>x.Length))Directory.Delete(dir);
   Directory.Delete(root);
  }}
  Console.WriteLine(count+" setup policy assertions passed; normalized ZIP verification/extraction tested for both Office payload names; no Office activation or installation performed.");return 0;
 }
}
