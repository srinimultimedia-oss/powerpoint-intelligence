# SECURITY REMEDIATION PLAN — Run 1

**Execution update:** the owner has approved feature-preserving remediation. H1, M1 and related L1 controls are implemented with local verification described in [SECURITY_REMEDIATION_STATUS.md](SECURITY_REMEDIATION_STATUS.md). This document remains the original broad plan; conditional/enterprise packages were not implicitly approved or marked complete.

Date: 2 November 2026 (Asia/Dubai). **Proposed plan, not implementation authorization. Stop until the owner approves Run 2.** Source assessment and original identities are in [CHECKMARX_FINDINGS.md](CHECKMARX_FINDINGS.md); execution handoff is [CODEX_REMEDIATION_INSTRUCTIONS.md](CODEX_REMEDIATION_INSTRUCTIONS.md).

## Executive assessment

Original report: **118 findings, 0 Critical, 116 High, 2 Medium; Low/Information excluded**. Available detail: 117 records; one High SCS aggregate lacks details. Four records statically confirm two root causes (corpus filename escape and unbounded repair work); three are conditional executable/toolchain risks; 110 are potential false positives tied to intended CLI/config/source selection; one SCS entry is blocked. These are assessment classifications, not scanner changes. No finding is fixed. The overview's download/untrusted-functionality mismatch needs the raw export; do not infer a new ID or alter CWE/severity to reconcile it.

Repository is Windows .NET Framework 4.8 VSTO, not a hosted multi-tenant file server. User-picked source files/roots are authorized input; generated artifacts require an owned-output boundary. Preserve that distinction throughout remediation.

## Approval and prerequisite decisions (H0)

1. Obtain original machine-readable SAST/SCS results, original scanner IDs/query IDs/engine identities, redacted SCS evidence, scanner/rules version and scanned ZIP checksum or commit mapping. Scan ID bed87600-7c2c-45b7-8b17-15b0b268ec82; source endpoints match but entire ZIP equivalence is not proven.
2. Record whether developer/CI caches or command arguments can be written by less-trusted actors while tests/signing execute with a more trusted token. Record actual cache/configuration directory ACLs. Do not invent a privilege boundary from a user selecting their own CLI output path.
3. Approve source-root/UNC policy separately from owned catalog/configuration/output policy. Determine legitimate redirected profiles and enterprise shares before enforcing physical-path rules.
4. Approve independent dependency trust material, distribution signing identity, Windows/Office x86/x64 test availability and resource budgets. Do not auto-trust a certificate or disable antivirus/Application Control.
5. Approve which conditional hardening packages proceed. Missing SCS detail blocks its exact remediation, not the independent corpus/repair fixes.

## Critical priority

**No Critical findings are included in the original report.** No Critical implementation package is invented. A later SCS export can change prioritization only through an explicitly documented reassessment.

## High priority

Original High severity is retained for every mapped record, even tentative false positives. The following packages are ordered by code-execution/signing impact, established attack surface and dependencies, rather than raw duplicate count.

### H2 — Executable/toolchain provenance (conditional)

Exact likely changes: `scripts/Piidx-TestSuite.ps1`, `scripts/Scene-RunnerCache.ps1`, `scripts/Test-SceneContracts.ps1`, `easySetup/Build-EasySetup.ps1`; shared tool-discovery callers in `scripts/Install-Local.ps1`, `Test-S1.ps1`, `Test-Piidx*.ps1`, `Test-TemplateLibrary.ps1`, `Test-All.ps1`; if a common helper is approved, add it under `scripts` and wire it explicitly.

Approach: first characterize the actor/ACL boundary. Protect trusted caches against lower-trust writes, authenticate cross-boundary artifacts independently of colocated receipts, validate approved tool roots/publishers, and reduce check/use races. A final hash check is defense-in-depth, not proof of immutable execution. Use an approved custom toolchain configuration for enterprise paths instead of this machine's literal paths.

Preserve: source/compiler/dependency hashes, stable exe names, bounded waits, scanner evidence/refusal, legacy audits, missing-file refusal, signed Release/easySetup behavior and explicit compile-versus-execution reporting. No file is restored from quarantine and no signature gate is weakened.

Tests/acceptance: existing runner cache/review tests remain; add tool publisher/environment substitution, trusted nondefault installation, combined receipt/binary swaps, final-window races, missing/evidence tamper, and legitimate Windows x86/x64 execution. Accept only after the named attacker cannot cause substituted code execution under the approved trust model. If no distinct actor boundary exists, preserve the finding as a documented conditional/potential false positive; do not fake a vulnerability or suppress the query.

Dependencies: H0 actor model, actual directory permissions and signing/CI policies. Complexity: **medium/large, roughly 3–5 engineer-days**, excluding enterprise policy rollout and native validation.

### H4 — Independently verify dependency restore (additional AR-01)

Exact likely changes: `scripts/Test-S1.ps1:16-24`, a proposed checked-in dependency integrity manifest/helper under `scripts`, callers that consume `artifacts/refs/newtonsoft/lib/net45/Newtonsoft.Json.dll`, and targeted restore tests. Production `PowerPointIntelligence.csproj` HintPath remains unless an approved dependency-management refactor requires adjustment.

Approach: pin independently established package and extracted DLL identity for the **existing 13.0.3 version**; verify before unpacking, copying and reuse. Stage uniquely and publish verified output atomically. Use safe extraction against path traversal/duplicates, validate legitimate package layout and return nonzero on any verification failure. Do not upgrade dependencies merely to clear an unrelated path finding.

Preserve: deterministic offline verified-cache use, assembly version/reference behavior, typed JSON/scene serialization and every downstream test. Tests: tampered archive/DLL, changed cached DLL, wrong identity, partial download, unsafe ZIP member and successful current-version offline restore. Acceptance: unverified bytes are never expanded into the accepted cache or loaded; altered cache fails closed. Blocker: approved independent trust/hash material, not a fabricated digest. Complexity: **medium, 1–2 days** plus approval latency.

### H1 — Contain data-derived corpus output (confirmed)

Exact likely changes: `tests/s3/SceneContractTests.cs:67-72`; add a minimal reusable safe-child/owned-output helper under `tests` or approved shared code; update `scripts/Test-SceneContracts.ps1`, `tests/s3/scene-fixtures.mjs`, `tests/s3/validate-roundtrip.mjs` only if their argument/name contracts require coordinated changes. Inspect the paired Node producer/consumer before changing names.

Approach: validate a bounded basename; reject all path/drive/UNC/ADS/device/dot forms, canonicalize root/candidate and enforce separator-aware containment. Never silently sanitize names into collisions. Prefer create-new or explicit owned replacement rules; preserve valid existing corpus identifiers and filenames. This is a fix to a concrete test data write primitive, not a reason to restrict all intentional CLI input paths.

Tests/acceptance: malicious names (both separators, root/drive-relative, UNC/extended, parent, sibling-prefix, ADS, reserved names, trailing dot/space, duplicates, oversized names) must be rejected before creating/modifying any outside file; valid shared fixtures still roundtrip across .NET and Node. Use disposable directories and an outside sentinel. Junction tests belong to M2 if hostile physical ancestry is in scope. Complexity: **small/medium, 0.5–1.5 days**. Dependency: characterize existing corpus naming; no SCS detail or enterprise API is needed.

### H3 — Review intended path authority and targeted hardening (potential false positives)

Exact likely files: `src/Library.Core/LibraryConfiguration.cs`, `PiidxCatalog.cs`, `PresentationSource.cs`; `apps/powerpoint-vsto/ServiceConnections.cs`, `DataSources.cs`, `ExcelRangeMapping.cs`, `DemoLinks.cs`, `OfficeCompatibility.cs`, `LocalLibraryWorkspace.cs`; test entrypoints/helpers listed per record in the register, especially `tests/s0/DemoOfficeTests.cs`, `DemoContractTests.cs`, `ServiceConnectionTests.cs`, `DataSourceTests.cs`, `DemoScenario.cs`, `LinkedChartDiagnostic.cs`, `Context*Tests.cs`, `HostIntegrationTests.cs`, `tests/library/TemplateLibraryTests.cs`, `TemplateFidelity*.cs`, `tests/s1/DomainTests.cs`, `tests/s3/NativeRendererTests.cs`, `SceneContractTests.cs` and `tests/compatibility/OfficeCompatibilityTests.cs`. Change only callers that actually need a stronger contract; this is an impact list, not an instruction to edit every file.

Approach: characterize production call graphs and define separate authorities: (1) authorized external input/source roots, (2) deployed application config, (3) owned mutable catalog/nonsecret config, (4) owned generated test/evidence output. An internal OwnedWorkspace/context or store abstraction can remove test-only arbitrary-path injection without inventing a server sandbox. Reject unsafe derived children; preserve explicit input selection and approved external roots. Validate before side effects or Office startup. Protect atomic temp cleanup and return nonzero with content-free diagnostics on invalid ownership.

Tests/acceptance: valid external/Unicode/space/long-path sources and custom indexes work; header/moved-source refusal, source association identity, migration, concurrency, atomic save/cancel and existing evidence survive. Each original source/sink must be re-reviewed against the new authority; retain original IDs. A demonstrated intended same-user path may remain a potential false positive with evidence; do not exclude test files, suppress queries or claim 0 scanner results without rerunning the scanner.

Dependencies: H0 config/root/CI authority decisions, H1's minimal safe-child primitive where reusable, M2 for hostile physical ancestry. Complexity: **large if broadly refactored, 3–6 days**; disposition-only evidence may require substantially less. A blanket allowlist forcing all user documents under the plugin folder is unacceptable.

## Medium priority

### M1 — Bound and transactionally publish offline repairs (confirmed)

Exact likely files: `scripts/Repair-TemplateAlignment.py:18-66`, a new bounded Python repair regression suite under `tests`, and its validation-script integration if approved. No package upgrade is needed to introduce limits.

Implementation approach: inspect ZIP metadata before any large read; reject duplicate/ambiguous members and enforce actual streamed bytes, not only claimed ZIP lengths. Introduce per-member/package bytes, member count, XML depth/node/text/shape, aggregate work, elapsed/cancellation budgets. Apply the same explicit entity/network-disabled parser in validation. Bound or redesign the quadratic container-to-label search. Validate a temporary new deck and sidecar completely before publication. Fail with explicit errors and keep input/previous files intact.

Proposed starting limits for characterization, **not already approved product limits**: package/member counts and total inflated bytes calibrated on current decks; align XML bounds with existing readers where appropriate (40 MiB maximum is already used by PresentationSource), but account for aggregate members and shape work. Record measured valid corpus maxima before choosing exact defaults; changing limits needs tests/owner-visible documentation rather than silent truncation.

Preserve: native editable text/shapes/groups, alignment function, original text values, every untouched part, source hash, timings/transitions and output validation semantics. Test small synthetic ZIP bombs, excessive valid nodes, quadratic shapes, validation-pass exhaustion, truncation, duplicate names, cancellation, failure cleanup and valid deck roundtrip in native PowerPoint. Acceptance: predictable measured memory/work bounds, no false success/partial published deck, original unchanged and functional/native checks passed. Complexity: **medium, 2–4 days** plus native test time. Independent of missing SCS metadata. Pair L1 to avoid duplicate work.

### M2 — Physical path and reparse hardening (additional AR-02, conditional)

Likely files: `PiidxCatalog.cs`, `PresentationSource.cs`, `LibraryConfiguration.cs`, `ServiceConnections.cs`, `easySetup/Bootstrapper.cs`, common owned-path helper if approved; new disposable Windows filesystem integration tests. Runtime .NET Framework 4.8 compatibility is mandatory.

Approach: specify where lexical validation suffices and where an actual opened-handle identity/final-path check is needed. Handle parent junctions/aliases and check-open races consistently; protect creation/replace operations in owned directories. Keep explicitly approved redirected profiles/UNC roots, source read permissions, SQLite locks/journals and installer version directories. No indiscriminate deletion or reparse-policy changes.

Acceptance: an explicitly modeled lower-trust actor cannot redirect protected operations to outside/unowned files; valid redirected/source scenarios pass. Test physical junction ancestors, races, canonical aliases, existing sentinels and normal paths. If privileges are insufficient to create an integration fixture, report BLOCKED rather than substituting a string-only test. Dependencies: H0 authority model and Windows test availability. Complexity: **large, 2–4 days** plus policy/test setup.

## Low priority

No original Low findings were included (filter exclusion is not absence). **L1 is additional AR-03 only**, preferably implemented with M1.

Likely files: `Repair-TemplateAlignment.py:43,46,53-66` and its new regression tests. Replace preservation asserts with explicit failures, reserve deck/sidecar safely, never overwrite existing sidecars, and remove/retain temporary diagnostics according to an explicit failure policy. Test optimized Python, deliberately failed comparisons, sidecar conflict and interrupted publication. Acceptance: validation cannot disappear under `-O`, no false success JSON and no clobbered unrelated output. Complexity: **small, 0.5–1 day incremental to M1**.

## Cross-cutting architecture and acceptance

- Keep source input authorization separate from output containment. `GetFullPath`, extensions and an unkeyed checksum are not independent authorization/provenance controls.
- Preserve .NET Framework 4.8, x86/x64 Office configuration, native chart/shape editability, current-deck insertion, data-source mapping, palette/alignment, template fidelity, agendas, connectors/indicators and optional animation/build behavior.
- Preserve explicit user selection/confirmed relinking, no sample fallback, metadata identity/corrupt-object preservation, readonly source access and native save/reopen.
- Do not log workbook contents, secrets, tokens or complete sensitive path graphs. Redact security fixtures and keep malicious cases inert.
- No automatic certificate trust, security-policy relaxation, antivirus exclusion, test removal, query suppression or unapproved dependency changes.
- Native validation is a desktop task; do not bypass the agent's AppContainer COM boundary. Use an approved normal Windows environment with current protection active.

## Regression and scanner acceptance gates

1. Add focused negative tests first for the approved change; initially demonstrate the relevant assertion fails on the baseline without unsafe writes/resource stress. Do not run a destructive proof on user files.
2. Run affected current suites: `scripts/Test-SceneRunnerCache.ps1`, `Test-SceneContracts.ps1` for AnyCPU/x86/x64 (respect their provenance/scanner gates), `Test-Piidx.ps1`, `Test-PiidxIsolation.ps1`, `Test-Ribbon.ps1`, `Test-S0.ps1`, `Test-S1.ps1` as appropriate; `easySetup/Test-EasySetup.ps1 -SourceOnly` for packaging source changes. A dependency restore test must not fetch unverified artifacts just to establish a green baseline.
3. After targeted checks pass, `scripts/Test-All.ps1` for complete local regression. Then desktop `Install-Local.ps1` and `Test-All.ps1 -InteractiveOffice` with presentations saved/closed. Use focused `Test-DemoFeatures.ps1 -InteractiveOffice -MappedExcelOnly` / `-LinkedChartDiagnostic` when relevant and `Test-TemplateLibrary.ps1 -InteractiveOffice` for fidelity. Signed installer changes require `easySetup/Build-EasySetup.ps1` and `Test-EasySetup.ps1`, plus actual install/upgrade/rollback/uninstall smoke tests.
4. Re-run the same Checkmarx preset/scope and compare original **result IDs + query ordinals + Similarity IDs**, including duplicates; obtain Low/Information separately if expanded coverage is approved. Track raw SCS details and untrusted-functionality taxonomy to closure. No scan is run or finding resolved in Run 1.
5. Security acceptance requires negative path/provenance/resource tests and code review, not only a green scanner. Functional acceptance requires executed Office x86/x64 evidence; unavailable machine/locale/build coverage stays explicit. Zero filtered SCA results is not a dependency attestation.

## Proposed sequence and complexity

Start H0 immediately after approval. Address H2/H4 first where a real code-execution/signing boundary is established; independently fix H1 and M1/L1 rather than waiting for unrelated enterprise decisions. Extend H3/M2 only under approved source/CI/ownership policies. Re-run functional/scanner gates and reconcile SCS before claiming complete security acceptance.

Confirmed corpus + repair fixes are roughly **3–6 engineer-days** before broad native/scanner validation. The full proposed provenance/path hardening and regression program is roughly **10–18 engineer-days**, plus **2–4 days** of available Windows/Office validation and external review latency. These are planning ranges, not measured delivery promises. Missing SCS evidence, enterprise decisions, signing/CI access or new findings can change scope. Breaking-change risks include rejected fixture names, narrower config/output locations, new large-deck rejection thresholds, custom toolchain restrictions and redirected-path handling; each requires compatibility tests and explicit release notes.

## STOP

Run 1 ends with these documents. **Wait for the owner's approval before modifying application code, tools, dependencies or scanner configuration in Run 2.**


## Exact original-record work-package index

Every detailed record is mapped once below. This is implementation grouping, not scan deduplication. Individual source/sink/fix/test records remain in CHECKMARX_FINDINGS.md.


### H2 / EXEC-PIIDX — Conditional risk

CX-R001 (Similarity 333933924)


### H2 / EXEC-SCENE — Conditional risk

CX-R002 (Similarity -810172859)


### H2 / EXEC-TOOL — Conditional risk

CX-R003 (Similarity 515670905)


### H3 / CONFIG-ROOT — Potential false positive

CX-R004 (Similarity -1305902851), CX-R005 (Similarity 531211230), CX-R006 (Similarity -1886406541)


### H3 / DISCOVERY — Potential false positive

CX-R008 (Similarity -415229814)


### H1 / CORPUS-NAME — Confirmed weakness

CX-R007 (Similarity 889475886), CX-R080 (Similarity -2077706329)


### M1 / ZIP-REPAIR — Confirmed weakness

CX-R116 (Similarity -900544832), CX-R117 (Similarity -1745516950)


### H3 / DATA-INPUT — Potential false positive

CX-R009 (Similarity -1401055022), CX-R010 (Similarity 314803386), CX-R011 (Similarity -763093660), CX-R012 (Similarity -576431034), CX-R013 (Similarity 936672265)


### H3 / SERVICE-CONFIG — Potential false positive

CX-R014 (Similarity 280573853), CX-R017 (Similarity 333146485), CX-R018 (Similarity -1699090574), CX-R019 (Similarity 280573853), CX-R020 (Similarity 333146485), CX-R024 (Similarity 706674842), CX-R025 (Similarity -1651521647), CX-R026 (Similarity 1188711609), CX-R027 (Similarity 507564216), CX-R028 (Similarity -47860084), CX-R029 (Similarity -1477818700), CX-R030 (Similarity -56861040), CX-R031 (Similarity -1477818700), CX-R032 (Similarity -47860084), CX-R036 (Similarity -1372051600), CX-R037 (Similarity -315100207), CX-R038 (Similarity -1654445831)


### H3 / CATALOG-PATH — Potential false positive

CX-R015 (Similarity 1291543521), CX-R016 (Similarity 1291543521), CX-R021 (Similarity 955143290), CX-R022 (Similarity -1947506698), CX-R023 (Similarity -1885568552)


### H3 / OFFICE-CONFIG — Potential false positive

CX-R091 (Similarity 1144353721)


### H3 / TEST-ARTIFACT — Potential false positive

CX-R033 (Similarity 1027859451), CX-R034 (Similarity -419043739), CX-R035 (Similarity 601440768), CX-R039 (Similarity 1760295004), CX-R040 (Similarity -90541740), CX-R041 (Similarity 1947077620), CX-R042 (Similarity 46418524), CX-R043 (Similarity -1299226345), CX-R044 (Similarity -1761994869), CX-R045 (Similarity -1720959708), CX-R046 (Similarity 462301936), CX-R047 (Similarity 1407586478), CX-R048 (Similarity 758400680), CX-R049 (Similarity -1517092728), CX-R050 (Similarity 2097746304), CX-R051 (Similarity 1237596672), CX-R052 (Similarity -152178112), CX-R053 (Similarity 934697572), CX-R054 (Similarity -1719932254), CX-R055 (Similarity -1345387132), CX-R056 (Similarity -86767479), CX-R057 (Similarity -695368680), CX-R058 (Similarity -1532911912), CX-R059 (Similarity -1429002613), CX-R060 (Similarity 2129211647), CX-R061 (Similarity -1429002613), CX-R062 (Similarity 1703303725), CX-R063 (Similarity 1552682661), CX-R064 (Similarity -1579543158), CX-R065 (Similarity -473184483), CX-R066 (Similarity -82668503), CX-R067 (Similarity 1966682696), CX-R068 (Similarity -485459168), CX-R069 (Similarity 361596616), CX-R070 (Similarity -756336288), CX-R071 (Similarity -1176590899), CX-R072 (Similarity -1657118664), CX-R073 (Similarity -753889073), CX-R074 (Similarity 711231073), CX-R075 (Similarity -2135577675), CX-R076 (Similarity 1335748697), CX-R077 (Similarity 745660673), CX-R078 (Similarity 920427747), CX-R079 (Similarity 2145587874), CX-R081 (Similarity -1426692475), CX-R082 (Similarity -1599453999), CX-R083 (Similarity -498998273), CX-R084 (Similarity 1235437117), CX-R085 (Similarity 1025575657), CX-R086 (Similarity -1153363084), CX-R087 (Similarity -620214048), CX-R088 (Similarity -1926027703), CX-R089 (Similarity -1549767052), CX-R090 (Similarity -1612927088), CX-R092 (Similarity 1398503197), CX-R093 (Similarity -457272103), CX-R094 (Similarity -293388255), CX-R095 (Similarity -771679709), CX-R096 (Similarity -661553133), CX-R097 (Similarity 1201223223), CX-R098 (Similarity 596844632), CX-R099 (Similarity 1018546989), CX-R100 (Similarity 596844632), CX-R101 (Similarity 596844632), CX-R102 (Similarity 341466547), CX-R103 (Similarity 596844632), CX-R104 (Similarity 596844632), CX-R105 (Similarity 1855789791), CX-R106 (Similarity -1574684448), CX-R107 (Similarity 360136018), CX-R108 (Similarity 596844632), CX-R109 (Similarity -2135358301), CX-R110 (Similarity 661522298), CX-R111 (Similarity 793367280), CX-R112 (Similarity 1739815605), CX-R113 (Similarity -1880250405), CX-R114 (Similarity -169465736), CX-R115 (Similarity 724762689)


High SCS summary slot: H0, blocked pending the original ID and trace. Additional AR-01 → H4; AR-02 → M2; AR-03 → L1.
