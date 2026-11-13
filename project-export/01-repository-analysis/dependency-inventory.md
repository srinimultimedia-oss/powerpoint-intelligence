# Dependency inventory

| Manifest | Kind | Name | Version | Resolution/integrity |
|---|---|---|---|---|
| apps/powerpoint-vsto/PowerPointIntelligence.csproj | Reference | Newtonsoft.Json | project/reference identity | ..\..\artifacts\refs\newtonsoft\lib\net45\Newtonsoft.Json.dll |
| apps/powerpoint-vsto/PowerPointIntelligence.csproj | Reference | System.IO.Compression | project/reference identity | framework/tooling resolution |
| apps/powerpoint-vsto/PowerPointIntelligence.csproj | Reference | System.IO.Compression.FileSystem | project/reference identity | framework/tooling resolution |
| apps/powerpoint-vsto/PowerPointIntelligence.csproj | Reference | System | project/reference identity | framework/tooling resolution |
| apps/powerpoint-vsto/PowerPointIntelligence.csproj | Reference | System.Core | project/reference identity | framework/tooling resolution |
| apps/powerpoint-vsto/PowerPointIntelligence.csproj | Reference | System.Xml | project/reference identity | framework/tooling resolution |
| apps/powerpoint-vsto/PowerPointIntelligence.csproj | Reference | System.Xml.Linq | project/reference identity | framework/tooling resolution |
| apps/powerpoint-vsto/PowerPointIntelligence.csproj | Reference | System.Configuration | project/reference identity | framework/tooling resolution |
| apps/powerpoint-vsto/PowerPointIntelligence.csproj | Reference | System.Runtime.Serialization | project/reference identity | framework/tooling resolution |
| apps/powerpoint-vsto/PowerPointIntelligence.csproj | Reference | System.Windows.Forms | project/reference identity | framework/tooling resolution |
| apps/powerpoint-vsto/PowerPointIntelligence.csproj | Reference | System.Drawing | project/reference identity | framework/tooling resolution |
| apps/powerpoint-vsto/PowerPointIntelligence.csproj | Reference | Microsoft.CSharp | project/reference identity | framework/tooling resolution |
| apps/powerpoint-vsto/PowerPointIntelligence.csproj | Reference | office, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71 e9bce111e9429c | project/reference identity | framework/tooling resolution |
| apps/powerpoint-vsto/PowerPointIntelligence.csproj | Reference | Microsoft.Office.Interop.PowerPoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71 e9bce111e9429c | project/reference identity | framework/tooling resolution |
| apps/powerpoint-vsto/PowerPointIntelligence.csproj | Reference | Microsoft.Office.Tools.Common, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a | project/reference identity | framework/tooling resolution |
| apps/powerpoint-vsto/PowerPointIntelligence.csproj | Reference | Microsoft.Office.Tools, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a | project/reference identity | framework/tooling resolution |
| apps/powerpoint-vsto/PowerPointIntelligence.csproj | Reference | Microsoft.Office.Tools.v4.0.Framework, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a | project/reference identity | framework/tooling resolution |
| apps/powerpoint-vsto/PowerPointIntelligence.csproj | Reference | Microsoft.Office.Tools.Common.v4.0.Utilities | project/reference identity | framework/tooling resolution |
| apps/powerpoint-vsto/PowerPointIntelligence.csproj | Reference | Microsoft.VisualStudio.Tools.Applications.Runtime, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a | project/reference identity | framework/tooling resolution |
| src/Chart.Domain/Chart.Domain.csproj | Reference | System | project/reference identity | framework/tooling resolution |
| src/Chart.Domain/Chart.Domain.csproj | Reference | System.Core | project/reference identity | framework/tooling resolution |
| src/Chart.Domain/Chart.Domain.csproj | Reference | System.Xml | project/reference identity | framework/tooling resolution |
| src/Chart.Domain/Chart.Domain.csproj | PackageReference | Newtonsoft.Json | 13.0.3 | framework/tooling resolution |
| packages/chart-domain/package-lock.json | locked npm | node_modules/ajv | 8.20.0 | integrity recorded |
| packages/chart-domain/package-lock.json | locked npm | node_modules/fast-deep-equal | 3.1.3 | integrity recorded |
| packages/chart-domain/package-lock.json | locked npm | node_modules/fast-uri | 3.1.9 | integrity recorded |
| packages/chart-domain/package-lock.json | locked npm | node_modules/json-schema-traverse | 1.0.0 | integrity recorded |
| packages/chart-domain/package-lock.json | locked npm | node_modules/require-from-string | 2.0.2 | integrity recorded |
| packages/chart-engine/package-lock.json | locked npm | node_modules/ajv | 8.20.0 | integrity recorded |
| packages/chart-engine/package-lock.json | locked npm | node_modules/fast-deep-equal | 3.1.3 | integrity recorded |
| packages/chart-engine/package-lock.json | locked npm | node_modules/fast-uri | 3.1.9 | integrity recorded |
| packages/chart-engine/package-lock.json | locked npm | node_modules/json-schema-traverse | 1.0.0 | integrity recorded |
| packages/chart-engine/package-lock.json | locked npm | node_modules/require-from-string | 2.0.2 | integrity recorded |
| packages/chart-engine/package-lock.json | locked npm | node_modules/typescript | 5.9.3 | integrity recorded |

Installed dependency trees and DLLs are separately inventoried as binary/vendor material, not source-authored features. Office/VSTO and winsqlite3 are platform/runtime prerequisites. CI uses actions/checkout@v4 and actions/setup-node@v4; immutable action commit pinning is not present. Lockfiles pin npm versions/integrity but do not authenticate every NuGet/Office/platform dependency. Test-S1 has an independent NuGet download path requiring review (security gap SEC-01), so a lockfile elsewhere must not be generalized to all consumers. No dependencies were updated/installed during this export.
