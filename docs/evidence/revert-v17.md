# Restore v17

Restored product/test/script source byte-for-byte to the verified v17 snapshot at artifacts/validation-20261027-000225-a2768410/source.

Verified source SHA256: DF745CA44CDA8C2B69347E24860FD782028A19F41DC4DB84546D0879CEC4C8D5.

The v18 icon redesign and v19 branding/configured-library changes were reversed. v17 functional work was preserved. Later source/doc changes and the v18 focus preview were preserved under artifacts/revert-v17-backup-20261027-003943; no executable was moved, duplicated or replaced to evade a runner guard.

Executed checks: exact v17 source fingerprint PASS; Test-Ribbon.ps1 PASS; AnyCPU/x86/x64 compilation PASS; git diff --check PASS. Logs: artifacts/revert-v17-ribbon.log and artifacts/revert-v17-compile-*.log.

The earlier v17 full local run remains 20/21 with the x64 scene provenance/quarantine hold; the full suite and native Office tests were not rerun for this exact-source restoration. No installer, Office presentation or antivirus setting was modified. Apply the restored version to the desktop with scripts/Install-Local.ps1 in normal Windows PowerShell.
