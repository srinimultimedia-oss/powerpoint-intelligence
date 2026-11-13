# Security gap register

Source-review assessment only; no new vulnerability scan or source remediation. Historical scanner classifications are retained separately from current source controls.

| ID / assessed severity | Evidence | Impact/status | Resolution / validation recommendation |
|---|---|---|---|
| SEC-01 High — dependency verification gap | `scripts/Test-S1.ps1:18` | Missing DLL triggers pinned-version NuGet download/expansion/downstream compilation or execution without independent package digest/signature in that script. HTTPS/version pin is not independent identity. Later scene checks do not cover every S1 consumer. | Approved digest/publisher before expansion/load; tampered download and offline-cache tests. |
| SEC-02 High — PIIDX scanner scope gap | `scripts/Collect-SceneThreatEvidence.ps1:23` `scripts/Collect-SceneThreatEvidence.ps1:42` | Regex filters SceneContractTests/TemplateLibraryTests and scene/template folders, not PiidxTestSuite.dll. Historical no-PIIDX-match cannot prove PIIDX was searched. Recovery reproduction controls remain, but evidence is incomplete. | Bounded explicit PIIDX path/hash matching; positive PIIDX and unrelated-data exclusion fixtures; retain old report as limited. |
| SEC-03 Medium — implicit XML readers | `scripts/Install-Local.ps1:21` `scripts/Test-DemoPackage.ps1:27` | Builder uses Safe-Xml; related installer/test readers retain implicit casts. No external fetch exploit asserted; local provenance/runtime defaults matter. | Consistent bounded no-DTD/resolver-null readers, version/hash and PS5.1 regressions. |
| SEC-04 High when connected — auth/action-policy absent | `services/ai-gateway/README.md:1` `services/ai-gateway/action-plan.schema.json:37` | No identity/provider/action executor now. Generic arguments/model status cannot be future authorization. Not a current live-Azure exploit. | Typed action schemas, authorization, approval state machine, egress/minimization before executor. |
| SEC-05 Medium — local-data lifecycle policy | `apps/powerpoint-vsto/ServiceConnections.cs:35` `src/Library.Core/PiidxCatalog.cs:32` | Catalog/text/paths and profile metadata can be sensitive; no application encryption/retention/central audit found. Windows user/ACL is boundary. | Classification, workstation ACL/backup/retention/deletion and encryption decision. |
| SEC-06 Medium — publisher distribution trust | `docs/DECISIONS.md:35` `scripts/Build-Integrity.ps1:39` | Local development publisher success does not establish target estate trust/renewal. | Approved publisher, timestamp/renewal policy, target x86/x64 acceptance. |
| SEC-07 Medium — native release evidence gaps | `scripts/Test-All.ps1:89` | CI/noninteractive success cannot establish installed ribbon/native save-reopen/build variants/DPI/focus. | Desktop negative/mutation/roundtrip matrix with retained evidence. |
| SEC-08 Low/compatibility — host fingerprints | `scripts/Source-Fingerprint.ps1:21` `easySetup/Common.ps1:14` | ICU/NLS ordering may classify unchanged package as stale; historical deployment notes demonstrate NLS reproduction. Not an integrity bypass claim. | Canonical ordinal ordering/shared PS5/PS7 vectors and new signed package after recipe change. |

## Historical Checkmarx state

The 3-Nov11:35 remediation document preserves High similarity IDs -2018192911/-355656028 (recovery paths), 1923684989/1239593030 (XML),1072025071 (fingerprint hash-only sink). Source contains artifact/safe-builder XML controls. The fingerprint flow is a documented potential false positive pending query/rescan review. One High 2 ms result had no details and was owner-deferred. `docs/security/CHECKMARX_20261103_113558_REMEDIATION.md:1`.

No “resolved”/“zero vulnerabilities” count is inferred. A source control, a unit assertion and scanner clearance are different evidence. Severity here is qualitative impact prioritization, not exploit proof. Nothing was disabled, suppressed or changed in application source by this export.
