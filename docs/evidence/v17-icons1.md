# v17 icon correction

Applied icon-only artwork/binding fixes to the restored v17 functionality. Build identity: PPI-DEMO-RIBBON-20261026-v17-icons1. The prior packaging freshness fix remains enabled.

All 26 workflow actions now have distinct semantic icons, rather than reusing category glyphs. Library, template colors/alignment, layout links, native editable columns, clustered charts and Excel/CSV/JSON/XML controls have appropriate representations. Artwork is original, transparent and rendered at native 16/32 px sizes.

Validation: Test-Ribbon.ps1 PASS for all 90 legacy and 26 workflow commands, all 140 image bindings, transparency/cache checks, distinct workflow pixels and relevant action pairs. AnyCPU/x86/x64 forced Compile targets PASS. Actual compiled DLL byte-marker checks PASS for all three targets, avoiding timestamp-based incremental false positives. git diff --check PASS. Logs: artifacts/v17-icons1-ribbon.log and artifacts/v17-icons1-compile-*.log.

Hash comparison confirms TemplateLibrary.cs, LocalLibraryGallery.cs, FeatureInventory.cs and ThisAddIn.cs remain identical to the v17 snapshot. No v19 branding or folder-setting behavior was reapplied. The DataSources.cs change only selects the Excel file icon. Existing SceneContractTests/TemplateLibraryTests executables were not rebuilt or copied for this task.

Previews: artifacts/demo-contracts/ribbon-icons.png (all controls), ribbon-icons-focus.png (workflow controls). These are generator renders, not installed-Office screenshots. Native ribbon appearance/loading still requires desktop verification.

With PowerPoint closed, run .\scripts\Install-Local.ps1 in normal Windows PowerShell. This forces a freshly compiled signed payload and increments the deployment revision. Installed-ribbon validation requires the v17-icons1 marker, so an older v17 package cannot be mistaken for the icon correction. No installed success is claimed by source/compile checks.
