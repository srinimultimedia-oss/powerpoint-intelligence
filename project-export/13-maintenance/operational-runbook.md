# Operational runbook

**Audience:** interactive Windows support/developer staff. No server deployment or administrative action is performed by this export.

Before rollout, verify exact code/package hashes, architecture and publisher trust in the intended Windows user. Save/close presentations, use the normal installer prompt and then confirm NexWave ribbon loading. Retain extraction folders referenced by VSTO. Use the current package validation script rather than judging readiness from a copied filename. `easySetup/Test-EasySetup.ps1:1`.

For a managed chart failure, capture only fixed stage/error metadata and the owned test evidence directory. Preserve native objects and original local source files. If refresh sees a moved file or changed headers, require confirmed relinking; do not automatically guess replacements. For library failures, check source registration, source root permissions and exact schema/integrity validation before replacing a catalog. User source folders are authoritative; an index is derived local state. `apps/powerpoint-vsto/MappedLinkOperations.cs:1`; `src/Library.Core/PiidxCatalog.cs:1`.

Artifact cleanup is a separate reviewed operation. Preserve active deployment files, stable runner caches/receipts, signing material and scanner history. Historical source/evidence archives are not executable recovery sources. The documentation export is never a source-distribution/install bundle. `scripts/Maintain-Artifacts.py:1`.

Uninstall and rollback use VSTO/easySetup's implemented mechanisms. There is no enterprise background scheduler, central telemetry, RBAC service or documented HA/DR runtime to operate. Backup and retention of presentations, linked files and source libraries need local organizational policy; default JSON and PIIDX paths are not a substitute. Open operational decisions remain Q-005/Q-006/Q-007/Q-008/Q-012.
