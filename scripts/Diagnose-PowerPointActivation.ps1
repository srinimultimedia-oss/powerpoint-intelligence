param([string]$OutputDirectory=(Join-Path $PSScriptRoot '../artifacts/activation-diagnostics'))
$ErrorActionPreference='Stop'
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class ActivationToken {
 [DllImport("advapi32.dll", SetLastError=true)] static extern bool GetTokenInformation(IntPtr h,int c,out int v,int n,out int needed);
 public static int Read(IntPtr h,int c) { int v,n; if(!GetTokenInformation(h,c,out v,4,out n)) throw new System.ComponentModel.Win32Exception(); return v; }
}
"@
$identity=[Security.Principal.WindowsIdentity]::GetCurrent()
$report=[ordered]@{Time=(Get-Date).ToString('o');OS=[Environment]::OSVersion.ToString();Process64=[Environment]::Is64BitProcess;User=$identity.Name;LocalAppData=$env:LOCALAPPDATA;Token=@{};Registration=@{};Activation=@{}}
foreach($entry in @{IsAppContainer=29;ElevationType=18;IsElevated=20;HasRestrictions=21}.GetEnumerator()) {
 try {$report.Token[$entry.Key]=[ActivationToken]::Read($identity.Token,$entry.Value)} catch {$report.Token[$entry.Key]=$_.Exception.Message}
}
foreach($view in @([Microsoft.Win32.RegistryView]::Registry64,[Microsoft.Win32.RegistryView]::Registry32)) {
 $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot,$view)
 try {
  $key=$base.OpenSubKey('PowerPoint.Application\CLSID');$clsid=if($key){$key.GetValue('')};if($key){$key.Dispose()}
  $server=if($clsid){$base.OpenSubKey("CLSID\$clsid\LocalServer32")}
  $report.Registration[$view.ToString()]=@{CLSID=$clsid;LocalServer32=if($server){$server.GetValue('')}}
  if($server){$server.Dispose()}
 } finally {$base.Dispose()}
}
$application=$null
try {
 $report.Activation.Stage='ProgID resolution';$type=[Type]::GetTypeFromProgID('PowerPoint.Application',$true)
 $report.Activation.Stage='COM class factory / instance creation';$application=[Activator]::CreateInstance($type)
 $report.Activation.Stage='Application.Version';$report.Activation.Version=$application.Version
 $report.Activation.Status='PASS'
} catch {
 $errorRoot=$_.Exception;while($errorRoot.InnerException){$errorRoot=$errorRoot.InnerException}
 $report.Activation.Status='FAIL';$report.Activation.Exception=$errorRoot.GetType().FullName
 $report.Activation.HResult=('0x{0:X8}' -f $errorRoot.HResult);$report.Activation.Message=$errorRoot.Message
} finally {if($application){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($application)}}
# No Quit: this diagnostic must never close an existing user application.
$path=Join-Path $OutputDirectory 'activation.json'
$report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $path
$report | ConvertTo-Json -Depth 8
Write-Output "Evidence: $path"
