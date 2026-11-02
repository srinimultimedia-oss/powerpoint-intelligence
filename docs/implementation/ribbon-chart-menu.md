# Large Charts menu beside Library

29 October 2026. Build: PPI-DEMO-RIBBON-20261028-v17-properties8-chart-menu.

Charts is now a large 32-pixel menu immediately after the large Library button in the Elements group. It uses original geometric charts-gallery artwork and the existing chart-family commands, callbacks and enablement. The previous small Charts entry is moved rather than duplicated. Keytip: CH under the Intelligence tab. All 90 legacy bindings and 29 workflows remain present.

Validation: Test-Ribbon checks adjacency/size/image identity, unique menu ID, icon rendering at 16/32 pixels, transparency, distinct artwork and existing command bindings. The generated light/dark icon sheet was visually inspected. Test-DemoFeatures compiles the native fixture. x64 compilation is logged in artifacts/large-charts-menu-build-x64.log; installed visual appearance has not been executed in this agent session. Test-InstalledRibbon now verifies the loaded large-menu binding beside Library.

Install-Local.ps1 builds and installs the new package from normal desktop PowerShell. Restart PowerPoint, open a presentation and choose Intelligence to see Library and Charts. The Charts dropdown retains linked-source insertion in the current presentation. No proprietary image asset is reused. Narrow-window ribbon collapsing remains controlled by Office.

The preceding property-core desktop acceptance (23 contextual + 8 font assertions) is preserved in artifacts/demo-features-97c75dce36974d3ea08640c33105af95/acceptance.json and docs/implementation/context-properties.md. A ribbon icon/layout change does not establish runtime acceptance for every chart choice or Office version.