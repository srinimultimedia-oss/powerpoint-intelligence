# NexWave — easySetup

Version: 1.0.0.17
Source SHA-256: 59B1BFC8498086DC7FA433071DC5CC0EEE34B9832473ED00EA290069A6315256
Generated UTC: 2026-11-03T06:22:29.1837301Z

Run PowerPointIntelligenceSetup.exe after saving and closing PowerPoint. Setup detects Office x86/x64, verifies/extracts its embedded payload to the current user's LocalAppData, and opens the normal VSTO publisher prompt. Keep extracted version folders: VSTO may depend on them. Installation is per Windows user. Updates use a newer signed deployment version: the prior easySetup package is uninstalled through VSTO and the previous package is offered for rollback if installation fails. A first migration from an older repository-based installation may require uninstalling that old package normally if VSTO reports a different installation location. Uninstall through Windows Apps / Programs and Features. No administrator rights, certificate trust stores, antivirus exclusions or security policies are changed by this bootstrapper. A company policy may still require administrator assistance.

## Requirements
Windows desktop, .NET Framework 4.8, supported x86/x64 PowerPoint 2016/2019/Microsoft 365, and Microsoft VSTO Runtime already installed. Prerequisites are not downloaded or silently installed. These product profiles do not establish certification on every Office build. The build computer also requires Visual Studio Office tools and a usable signing certificate. Self-signed development publishers may not be trusted on another computer; use an approved publisher certificate for distribution. Private keys are never included.

## Features in this source
Editable native charts and grouped-shape rendering; Excel/CSV/JSON/XML data sources; explicit Excel worksheet/range/column mapping, refresh and confirmed relinking; thumbnail template library and palette/alignment controls; agendas; text elements, process layouts, universal connectors, Harvey balls and checkboxes; contextual properties; layout, formatting and quality tools. External services need separate endpoints and credentials.

Command labels discovered from the current source:
Generate linked report, Clustered column, Clustered bar, Stacked column, 100% stacked column, Stacked bar, 100% stacked bar, Line, Profile, Area, Pie, Doughnut, Scatter, Bubble, Column and line, Waterfall diagram, Mekko diagram (%), Mekko diagram (units), Gantt diagram, Text box, Rounded rectangle, Process flow, Checkbox, Harvey ball, Table, Agenda chapter, Image from file, Insert symbol, Grid layout, Paste lines into objects, Replace font, Resize font, Switch decimal separators, Proofing language, Edit chart data, Toggle legend, Add value labels, Toggle data table, Chart style, Value line annotation, Difference arrow annotation, CAGR annotation, Chart scanner, File path on slide, Import slides, Apply template, Export slide image, Apply local brand colors, Toggle value-axis title, Style chart and plot area, Chart title, Duplicate slide, Save selected slides, Remove presentation comments, Round chart workbook data, Style chart workbook header, Insert chart from JSON, Export selected chart JSON, Quick arrange toolbar, Edit table datasheet, Browse local asset library, Agenda table of contents, Apply style file, Insert Excel CAGR formula, Chart from Excel file, Refresh selected Excel link, Manage local workbook links, Current slide workbook list

## Included assets and validation limits
Included template packages: 62, licensed-computer use only. Third-party templates are not independently authored assets and must not be redistributed outside your licenses. Native fidelity evidence covers 16 tested original slides, not the complete library. The focused desktop mapped Excel test passed six assertions and the focused linked-chart test passed refresh/edit/save-reopen. Full native suite acceptance remains open; packaging does not prove runtime acceptance. Review docs/implementation for detailed evidence.

## Rebuild and automatic refresh
Run .\easySetup\Build-EasySetup.ps1 from normal Windows PowerShell. Optionally supply -CertificateThumbprint. Signed Release builds call this script automatically; recursive packaging is disabled. Start .\easySetup\Watch-EasySetup.ps1 for rebuilding after source edits. Saving source alone does not produce a signed installer unless the watcher is running. A failed build leaves the previous EXE unchanged; use Test-EasySetup.ps1 to reject stale packages before sharing.

## Logs and deployment files
Setup logs: %LOCALAPPDATA%\PowerPointIntelligence\easySetup. The EXE contains both architecture payloads, signed manifests and template assets. easySetup/build.json records source/package hashes and signer details. README.md is regenerated with each successful package. Build work/evidence remains in artifacts. Do not copy source credentials into the package.
