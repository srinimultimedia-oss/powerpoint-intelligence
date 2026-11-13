# External interfaces

## Office object models

The deployment references Office/PowerPoint 15 PIAs with EmbedInteropTypes and requires the common Office 16 API at runtime. This does not mean Office 15 is the supported runtime edition. Excel operations use dynamic COM rather than a referenced Excel interop package. `apps/powerpoint-vsto/PowerPointIntelligence.csproj:30` `apps/powerpoint-vsto/OfficeCompatibility.cs:69`.

Source workbooks are opened read-only for input. Mapped Excel reads explicitly set AutomationSecurity=3 and AskToUpdateLinks=false before opening; column selection and range signature/header checks guard refresh. The older DemoLinks path is separate and stores a version-1 file association; it does not provide general mapped-range semantics. `apps/powerpoint-vsto/ExcelRangeMapping.cs:59` `apps/powerpoint-vsto/DemoLinks.cs:38`.

## File and local database integrations

| Integration | Input / output | Authority / failure behavior |
|---|---|---|
| XLSX/CSV/JSON/XML data source | Bounded local table → native chart/workbook, presentation link metadata | Explicitly selected source path, parse/model/range validation; no sample fallback on bad source |
| PPTX/POTX library source | ZIP/XML source metadata → PIIDX SQLite catalog | Local file parsing; source URI/path and hashes; no proprietary index dependency |
| Native source template access | Hidden read-only PowerPoint source → insertion into current presentation | Separate source read access authorized in decisions; native formatting/geometry acceptance needs Office tests |
| VSTOInstaller process | /Install or /Uninstall plus quoted .vsto path → exit code/publisher prompt | Required runtime already installed; no prerequisite download or trust-store changes by bootstrapper |
| winsqlite3 native API | Parameterized SQL/database handles → local catalog | Windows system capability; no server DB or network database connection |
| NuGet/npm restoration | Locked JS dependencies and pinned Newtonsoft version → development dependencies | Build/test tooling, not runtime service calls; supply-chain gap below |

## Actual network use

There is no implemented live Azure, Tableau, MCP or asset-provider client. `Invoke-WebRequest` in Test-S1.ps1 downloads a pinned Newtonsoft.Json 13.0.3 NuGet package if the expected DLL is absent. This is a development dependency restoration path, not an AI API. It lacks a pinned package checksum before expansion in that script. `scripts/Test-S1.ps1:18`. npm lockfiles contain resolved package URLs/integrity; Restore-LockedDependencies uses npm ci --ignore-scripts. `scripts/Restore-LockedDependencies.ps1:6`.

Browser-visible hyperlinks and descriptive service labels must not be counted as live data integrations. No REST status codes or OAuth redirect URIs can be derived for absent services.
