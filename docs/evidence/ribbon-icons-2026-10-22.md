# Original ribbon icons — 22 October 2026

Added original geometric icons to all 105 buttons and menus in the Intelligence
ribbon, including all 90 local tool commands. Chart icons distinguish clustered,
stacked, 100%, line, area, pie/doughnut, scatter/bubble, combination, waterfall,
Mekko and Gantt tools. Arrangement, text, data, slide, asset and workspace tools
have corresponding pictograms. Labels and existing tooltips remain available.

`RibbonIcons.cs` draws directly at 16 or 32 pixels into transparent ARGB bitmaps.
The ribbon XML declares one `loadImage` callback and an explicit image key for
each button/menu. Images are converted to OLE pictures and cached for the add-in
lifetime. There are no external icon files or downloads. Office 2016, 2019 and
Microsoft 365 use the same callback path; runtime certification for those exact
builds is still pending. Microsoft documents the ribbon's cached image mechanism
and accepted image types in [Ribbon LoadImage](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.tools.ribbon.ribbonbase.loadimage?view=vsto-2022)
and [RibbonLoadImageEventArgs.Image](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.tools.ribbon.ribbonloadimageeventargs.image?view=vsto-2022).

The compiled ribbon contract test verifies every image binding, both dimensions,
nonempty artwork, transparent pixels, callback presence and cached OLE picture
creation. It generates a contact sheet with light/dark backgrounds. That contact
sheet was visually reviewed; this is not a screenshot of the installed ribbon.
Office theme/high-DPI behavior and the actual host callback still need desktop
validation. Build marker is `PPI-DEMO-RIBBON-20261022-v3`.

Installation attempt: `artifacts/install-20261022-150002.log`. It stopped at
`New-SelfSignedCertificate`, returning `0x80070005` before building or invoking
VSTOInstaller. No current-source installation or runtime loading is claimed.
From the repository in the desktop user's PowerShell session, close PowerPoint
after saving work and run `./scripts/Install-Local.ps1`, then reopen PowerPoint
and select Intelligence. Do not use `-SkipBuild` for this source update.

All 16 noninteractive checks passed in
`artifacts/validation-20261022-150002-46fb4f9d`, including ribbon icon tests and
AnyCPU/x86/x64 compilation. Source snapshot SHA256:
`84FF55973F6317406AF8029FEE0BFCF8013A2EADF3D5857E02FAAA858AB455BC`.
S0's intermittent native Office crash remains unresolved.
