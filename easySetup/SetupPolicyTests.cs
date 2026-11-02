using System;
using System.IO;
internal static class SetupPolicyTests {
 private static int count;
 private static void Check(bool value){if(!value)throw new Exception("Setup policy assertion failed.");count++;}
 private static void Reject(string root,string name){try{IntelligenceSetup.SafePath(root,name);}catch(InvalidDataException){count++;return;}throw new Exception("Unsafe package path accepted.");}
 public static int Main(){string root=Path.Combine(Path.GetTempPath(),"ppi-setup-policy");Check(IntelligenceSetup.SafePath(root,"x64/PowerPointIntelligence.vsto")==Path.Combine(root,"x64","PowerPointIntelligence.vsto"));Reject(root,"../escape.exe");Reject(root,"/absolute");Reject(root,"C:/outside");Reject(root,"x64/file:stream");Reject(root,"x64\\file");Reject(root,"");using(var stream=new MemoryStream(new byte[]{1,2,3})){Check(IntelligenceSetup.Digest(stream)=="039058C6F2C0CB492C533B0A4D14EF77CC0F78ABCCCED5287D84A1A2011CFB81");}Console.WriteLine(count+" setup policy assertions passed; no Office activation or installation performed.");return 0;}
}
