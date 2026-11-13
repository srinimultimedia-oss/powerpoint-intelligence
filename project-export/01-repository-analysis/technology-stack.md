# Verified technology stack

| Technology | Version/target | Role and limitation | Source |
|---|---|---|---|
| C# | 7.3; .NET Framework4.8 | VSTO and standalone domain projects, explicit compile lists | apps/powerpoint-vsto/PowerPointIntelligence.csproj:9 |
| Office | PowerPoint COM and VSTO; Office15 PIAs embedded | Desktop Office16 API baseline policy; profiles do not certify2016/2019/365 builds | apps/powerpoint-vsto/OfficeCompatibility.cs:1 |
| UI | Windows Forms/System.Drawing | Taskpane, dialogs, floating properties, procedural icons | apps/powerpoint-vsto/ThisAddIn.cs:88 |
| TypeScript/JS | Node engine>=24; chart-core/domain>=20 | Pure deterministic calculations/strict schemas and offlineCLI; current Nodev24.14.0 | packages/chart-engine/package.json:1 |
| SQLite | OS winsqlite3.dll; FTS5 capability required | Native bridge and exact PIIDX schema checks; not shippedsqlitebundle | src/Library.Core/NativeSqlite.cs:1 |
| JSON | Newtonsoft.Json13.0.3; AJV8.20.0 | Runtime domain decoding vs Node schema developer validation | src/Chart.Domain/Chart.Domain.csproj:12 |
| Build | VisualStudio/MSBuild/Roslyn, PowerShell | WindowsOffice tools and signing trust; no solutionfile | scripts/Install-Local.ps1:1 |
| AI/API | Schema/profile scaffold only | No HTTP server, SDK client, liveAzure calls or identity middleware | services/ai-gateway/README.md:1 |

No React/Electron/browser frontend, Office.js add-in, SmartArt authoring engine, cloud function, ASP.NET runtime service or enterprise database implementation was found in production source. Their presence in requirements does not establish implementation. Dependency currency/CVE status was not queried online; versions below are the inspected lock/project declarations.
