# v17-icons2: collapsed-group icon fix

## Root cause confirmed in PowerPoint

The existing desktop receipt reports successful installation of deployment 1.0.0.26, and the deployed DLL contains v17-icons1. Native UI inspection opened an empty, owned presentation and selected Intelligence. At the observed 1426-pixel window width, Data, Text, Layout, Color, Quality, Collaboration and More collapsed into seven identical generic-circle buttons. Opening Data showed that its actual command icons rendered correctly. The problem was missing image attributes on the ribbon groups, not missing command artwork or another stale DLL.

Earlier generator tests selected only button/menu nodes. They therefore missed the collapsed group controls. Disabled, grey command icons were also observed and correctly reflect missing chart/object selections; no compatibility/selection guard was weakened.

## Change

All ten workflow groups now have explicit 32 px semantic images through the existing loadImage callback. More has a dedicated nine-tile symbol; each group has distinct rendered artwork. Existing command artwork, v17 behavior and the packaging freshness fix remain intact. Marker: PPI-DEMO-RIBBON-20261026-v17-icons2.

Regression coverage now includes groups alongside buttons/menus, requires the ten group images at 32 px and verifies that their pixels differ. The installed-ribbon script requires icons2 and all ten group image bindings. Microsoft specifies image/getImage support on group elements: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui/190f51e6-3a3c-4cf8-affb-cec9b906aae4

## Executed validation

- Test-Ribbon.ps1 PASS: all 150 button/menu/group image bindings, transparency/cache behavior, ten distinct group images, 26 distinct workflow icons and all existing command bindings.
- Forced AnyCPU/x86/x64 compilation PASS; actual compiled DLL byte-marker checks confirm icons2 in all three outputs.
- git diff --check PASS.
- Generator preview on light/dark backgrounds reviewed: artifacts/demo-contracts/ribbon-icons-focus.png. All-control preview: ribbon-icons.png in the same folder.
- Logs: artifacts/v17-icons2-ribbon.log and artifacts/v17-icons2-compile-*.log.

Native observation establishes the pre-fix failure. The new group icons have not yet been installed or observed after installation. No signed package or native post-fix success is claimed. The owned blank inspection presentation was closed without saving or changing user presentations.

## Apply and verify

With PowerPoint closed, run .\scripts\Install-Local.ps1 in normal Windows PowerShell and complete the normal publisher prompt. Open Intelligence at a normal/narrow window width: the collapsed Data/Text/Layout/Color/Quality/Collaboration/More controls should now show distinct symbols. Open each group and confirm its command images still render. Then run .\scripts\Test-InstalledRibbon.ps1 for source/loaded-package identity and group binding checks; that structural check does not replace visual inspection.
