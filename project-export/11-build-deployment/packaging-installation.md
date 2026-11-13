# Packaging, installation and upgrades

## Layout and identities

One AnyCPU managed bootstrap EXE embeds IntelligencePayload.zip: x86/x64 folders, DLL/config/application/deployment manifests, dependencies, licensed POTX assets, generated README and payload.sha256. Public file/storage/COM identities retain PowerPointIntelligence; display product/ribbon is NexWave. Private signing keys are not payload assets. `easySetup/Build-EasySetup.ps1:74` `easySetup/Bootstrapper.cs:13`.

## Bootstrap transaction

Main is STA/WinForms. It refuses running POWERPNT processes without closing work; requires .NET4.8 registry Release>=528040, a registered single PowerPoint x86/x64 PE with major16 and installed VSTOInstaller. No prerequisite is downloaded. It asks visible confirmation. `easySetup/Bootstrapper.cs:36` `easySetup/Bootstrapper.cs:20`.

```mermaid
sequenceDiagram
  actor User
  participant Setup
  participant Package
  participant Cache as Per-user versions cache
  participant VSTO as VSTOInstaller
  User->>Setup: Run EXE and confirm
  Setup->>Setup: Verify closed host and prerequisites
  Setup->>Package: Verify bundle and per-file SHA inventory
  Setup->>Cache: Bounded safe extraction into bundle hash version
  alt Previous different manifest
    Setup->>VSTO: Uninstall prior package
  end
  Setup->>VSTO: Install selected architecture manifest
  alt Success
    Setup->>Cache: Save installed-manifest receipt
    Setup-->>User: Completed; runtime validation remains
  else Install failure after upgrade removal
    Setup->>VSTO: Attempt previous manifest install
    Setup-->>User: Error and rollback log
  end
```

Bounds: <=5000 entries, declared expanded sum<=512 MiB, inventory<=2,000,000 bytes, exact expected entry count+1, each entry digest checked. SafePath rejects colon/backslash/root escape. Existing version files must hash identically and are not overwritten. Destination is named by bundle hash. `easySetup/Bootstrapper.cs:53` `easySetup/Bootstrapper.cs:15`.

Cache/logs/installed-manifest.txt live under per-user LocalAppData/PowerPointIntelligence/easySetup. Keep VSTO-referenced version folders. Upgrade validates prior receipt under versions, exact .vsto filename and existence; stale prior receipt requires normal repair/uninstall. Different location from a repository-local install may need normal migration/uninstall. `easySetup/Bootstrapper.cs:64`.

## Local install alternative

Install-Local -Platform x64/x86 signs/builds the repository manifest then launches visible VSTOInstaller; SkipBuild still requires payload freshness/manifests. If no certificate selected, it may create a 90-day nonexportable development certificate, but imports no trusted root and changes no policy. This differs from builder's approved-chain requirement. `scripts/Install-Local.ps1:38`.

## Uninstall/rollback

Use Windows Apps/Programs and Features/normal VSTO removal. Changed-package failure attempts previous VSTO installation; rollback failure is logged separately. No central MSI/remote repair/database rollback service exists. Do not delete cache folders before removing installed VSTO dependencies.
