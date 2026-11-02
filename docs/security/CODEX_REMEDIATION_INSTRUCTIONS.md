# CODEX REMEDIATION INSTRUCTIONS — Run 2 handoff

**Handoff update:** feature-preserving H1/M1/L1 work has subsequently been authorized and implemented. Read [SECURITY_REMEDIATION_STATUS.md](SECURITY_REMEDIATION_STATUS.md) before continuing so completed local work is not repeated and missing scanner/native evidence is not mistaken for a pass.

Prepared **2 November 2026 (Asia/Dubai)**. This is an implementation handoff **only after explicit owner approval**. The existence of this document is not approval. Run 1 did not change application code, dependencies, signing/trust or scanner settings and did not resolve a finding.

## Fixed investigation context

- Work only in `C:/Srini/powerpoint-intelligence/powerpoint-intelligence`, not the historical Codex work copy. Follow AGENTS.md and docs/DECISIONS.md. Runtime remains Windows desktop VSTO / .NET Framework 4.8 with x86/x64 Office profiles.
- Authoritative report: `C:/Users/srini/Downloads/scan report.md`; raw SHA-256 `82ecb7f276330342f20da4518bb9b853650bc30b973ce11a6d603addb4a62ff9`; scan `bed87600-7c2c-45b7-8b17-15b0b268ec82`, 2 November 2026 09:21 Asia/Dubai. Baseline HEAD `1f3259c15caf5affba44df37d9b43f2d6f00d205` plus working files. All 234 source/sink snippets matched their reported lines at assessment time.
- Original summary: 118, Critical 0 / High 116 / Medium 2; Low and Information filtered out. Detailed rows: 117 (115 High / 2 Medium), plus one unidentifiable High SCS aggregate (engine 2ms). Query overview says two downloads + one untrusted-functionality result, while details group three downloads. Obtain raw original IDs/query/engine export; never invent a missing record, change CWE/severity to fit totals or merge SCS by assumption.
- Four static confirmed records represent two root causes; three conditional executable/toolchain risks; 110 potential false positives; one blocked SCS slot. Keep those categories separate from original severity. All detailed records remain New / To Verify in the original report. No CVE identifiers were supplied.
- Read CHECKMARX_FINDINGS.md for per-record original Similarity ID/result URL/query ordinal, source/sink, evidence, fix, regression risk and tests. SECURITY_REMEDIATION_PLAN.md contains approved-scope candidates H0/H1/H2/H3/H4/M1/M2/L1 and exact identity groups. No need to repeat the entire investigation unless source/hash or threat model changed.

## Before touching implementation

1. Confirm the owner's Run 2 approval and precise work packages. Record new HEAD/working tree state; preserve unrelated existing changes. If reported lines have drifted, locate the exact snippets and update references without changing original IDs.
2. Obtain missing SCS detail/redacted evidence and scanner source identity. SCS remediation stays blocked without it; independent confirmed fixes may proceed under approval.
3. Record actual runtime boundaries: who writes corpus/config/cache/CLI args, who runs builds/signing, permissions and CI trust. Owner-selected source files are not automatically untrusted output fragments.
4. Establish a baseline without unsafe network restores or arbitrary executable execution. Preserve existing stable cache filenames and every scanner/legacy-recovery gate. Do not rebuild a quarantined/unreceipted runner merely to get green tests. Coordinate any necessary source-reviewed rebuild through the established approved process.

## Implementation recipe: H1 — confirmed corpus output escape

Original records: CX-R007, Stored_Path_Traversal Result 4 of 5, Similarity **889475886**; CX-R080, Path_Traversal Result 72 of 107, Similarity **-2077706329**. Both destination `tests/s3/SceneContractTests.cs:72`; actual data path at lines 67–72 is cases.json → item.name → Path.Combine(output,name+suffix) → WriteAllText.

1. Characterize existing fixture names in `tests/s3/scene-fixtures.mjs` and roundtrip expectations in `tests/s3/validate-roundtrip.mjs`. Keep all valid names and shared C#/Node semantics.
2. Add focused negative tests first using an owned temporary root and a separate sentinel. Use a valid existing scene payload with hostile name values; no destructive paths outside disposable test storage.
3. Introduce a testable safe-child helper with explicit owned output root, bounded basename rules and separator-aware full-path containment. Reject parent/root/drive-relative/UNC/extended paths, slash/backslash, ADS, device names, trailing dots/spaces and collisions. `GetFileName` stripping or replacing characters is not a safe remediation because it silently aliases inputs.
4. Ensure validation precedes any write, using create-new or explicit owned artifact replacement. Preserve stable executable names; scripts may retain their approved output-root argument without treating every user-chosen root as an exploit.
5. Run valid and malicious corpus/roundtrip tests for supported architectures through existing provenance controls. Acceptance: invalid data cannot create/modify outside files, sentinels unchanged and all valid scene contracts still roundtrip. Map both original records to evidence, not one deduplicated entry.

## Implementation recipe: M1 + L1 — repair budgets and integrity

Original records: CX-R116, Similarity **-900544832**, read at `Repair-TemplateAlignment.py:48` to bodies loop :31; CX-R117, Similarity **-1745516950**, validation read :56 to XPath :62. Both Medium/CWE-606. Additional AR-03 covers asserts/output-sidecar behavior at :43,:46,:53–66.

1. Add bounded fixture-driven Python tests; do not generate a resource-exhausting archive in the agent/user environment. Small archives should exceed injected test limits deterministically.
2. Introduce one explicit budget/cancellation context shared by repair and verification: archive members, per-member and total streamed inflated bytes, XML depth/nodes/text/shape work, elapsed time and geometry matching cost. ZIP metadata estimates are only preflight, not the final limit. Reject ambiguous/duplicate members; keep package-relative names as package identifiers rather than writing them as paths.
3. Calibrate and document default limits from valid corpus statistics. Do not silently truncate slides/fields. Bound/redesign container-label quadratic matching, not just the final body loop.
4. Use explicit entity/network-disabled parsers consistently; replace preservation asserts with explicit checked exceptions so optimized Python cannot skip them.
5. Reserve new deck and sidecar outputs transactionally. Reject existing sidecars, publish success only after complete validation, and handle interruption/failure without modifying input or publishing a partial success.
6. Preserve all original text, untouched part bytes, geometry where not intentionally repaired, timings/transitions and native editability. Tests include `-O`, failed preservation, sidecar collision, truncated ZIP, huge/deep XML under small injected limits, cancellation and valid native PowerPoint reopen. Acceptance requires measured bounds and source hash preservation, not only no exception on a normal deck.

## Implementation recipe: H2/H4 — provenance, if approved

Original records CX-R001/CX-R002/CX-R003 have Similarity IDs **333933924**, **-810172859**, **515670905**. They are conditional, not proof of a network download: sources/sinks are `scripts/Piidx-TestSuite.ps1:11→26`, `scripts/Scene-RunnerCache.ps1:18→scripts/Test-SceneContracts.ps1:58`, and `easySetup/Build-EasySetup.ps1:9→10`.

Preserve existing hash/source/evidence gates; define the missing actor boundary before adding controls. Use approved cache ACL/ownership, authenticated cross-boundary artifacts, validated toolchain roots/publisher and immediate pre-execution checks. Do not replace scanner review, restore detected files, adopt legacy receipts, relocate stable executable names, or claim hashing a same-folder receipt authenticates its author. Ensure the installer watcher/Release hook still rebuilds both signed architectures and regenerates README after approved plugin changes.

Additional AR-01 is the actual unchecked network restore at `scripts/Test-S1.ps1:16–24`. Obtain approved independent digest/signature material for the existing Newtonsoft.Json 13.0.3 package/DLL, verify before extraction/load and every accepted reuse, use bounded safe extraction and atomic cache publication. Keep version stable unless dependency changes are separately approved. No guessed hash, no checksum fetched from the same compromised response, no automatic root trust or execution-policy bypass.

Acceptance tests: corrupted ZIP/DLL, cache substitution, wrong signer/tool/environment, missing/legacy/scanner evidence, interrupted restore and legitimate offline restore/both-architecture builds. Document the actual privileges and residual races; a green hash test is not proof of physical directory ownership.

## Implementation recipe: H3/M2 — intended paths and conditional physical identity

The 110 potential false positives have exact original identities in the mapping below and the register. Many are test Main(args) → a selected fixture directory → a fixed artifact or intentional input read. Production service UI uses its fixed per-user DefaultPath; managed library configuration intentionally registers external roots; Office startup reads its own deployed configuration. Preserve those facts instead of implementing a universal deny-outside-app rule.

If approved, introduce separate source, deployed-config, owned-catalog and owned-output authorities. Validate before side effects, use explicit per-run workspace/store context and safe child generation, and preserve CLI/wrapper contracts. For AR-02 physical identity, model actual hostile ancestor/TOCTOU access and use suitable Windows handles/final-path/ownership checks; lexical containment alone is not enough. Maintain approved redirected profiles/network shares and SQLite lock/journal behavior. Windows integration tests requiring unavailable privileges remain blocked, not replaced by a passing string test.

Review legacy `DemoLinks.Transfer` Excel-open security policy compared with the modern readers, if that command remains exposed; do not claim macro execution merely from missing explicit AutomationSecurity. Current external-service code is configuration-only, not an implemented HTTP/auth server. Actual service authorization/credentials/retention must follow Q-003/Q-004/Q-005, not ad hoc assumptions during a path fix.

## Required preserved behavior

- All existing ribbon commands/icons/task pane/context controls and native selection/STA mutation behavior.
- Native editable charts/grouped shapes and metadata identity; corrupted/unsupported objects preserved.
- Excel/CSV/JSON/XML linking, explicit worksheet/range/columns, changed-header/moved-file refusal, confirmed relinking and no sample fallback.
- Template thumbnails/source fidelity/palette/alignment, current-deck insertion and optional builds/transitions; no change to third-party asset licensing.
- Agendas, text shapes, connectors, process grids, indicators, layout/quality and saved native editing.
- PIIDX identities, digest/schema verification, search/preview cache, readonly source indexing, cancellation, writer locks and migration rollback.
- Signed x86/x64 deployment/easySetup freshness and README generation, normal VSTO prompts and no automatic security-store changes.

## Validation and reporting contract

Run only tests appropriate to approved edits; once targeted checks pass broaden to the complete local suite. Existing commands include Test-SceneRunnerCache, Test-SceneContracts (AnyCPU/x86/x64 with evidence gates), Test-Piidx, Test-PiidxIsolation, Test-Ribbon, Test-S0, Test-S1, Test-All, and easySetup/Test-EasySetup -SourceOnly. Do not fetch unverified dependencies to make a baseline pass.

For native acceptance use approved normal Windows desktop PowerShell, with protection active and presentations saved/closed: Install-Local then Test-All -InteractiveOffice, plus focused MappedExcelOnly/LinkedChartDiagnostic, native renderer and original template fidelity when affected. Never work around the agent AppContainer COM/signing boundary. Record Office build, actual process architecture, source fingerprint, per-scenario results, evidence paths and blocked coverage. Existing historical passes are not new remediation evidence.

Re-run Checkmarx with the original preset/filter/scope and retain the 117 individual original result URLs, all repeated Similarity IDs and the SCS record once supplied. No file exclusion, query suppression, severity rewrite or unexplained finding drop. A result is fixed only after source change, focused negative tests, preserved functional behavior and scanner/reviewer evidence are reconciled. Potential false positives remain tentative until reviewed; no automatic scanner disposition is authorized by this plan.

Deliver a per-original-record before/after status, exact changed files, security/functional tests, source fingerprint, scan identity, residual risk and unavailable tests. If a newly exposed boundary changes assumptions, stop that package and ask rather than silently changing enterprise configuration or supported behavior. No claim of full completion while the SCS detail or required native tests remain blocked.

## Per-record execution index

Use the following index with the detailed register: package, local alias and original Similarity ID. Duplicate Similarity IDs remain separate and must be matched by the original result URL/query ordinal. These groups cover all 117 detailed records exactly once.


### H2 / EXEC-PIIDX

CX-R001 (Similarity 333933924)


### H2 / EXEC-SCENE

CX-R002 (Similarity -810172859)


### H2 / EXEC-TOOL

CX-R003 (Similarity 515670905)


### H3 / CONFIG-ROOT

CX-R004 (Similarity -1305902851), CX-R005 (Similarity 531211230), CX-R006 (Similarity -1886406541)


### H3 / DISCOVERY

CX-R008 (Similarity -415229814)


### H1 / CORPUS-NAME

CX-R007 (Similarity 889475886), CX-R080 (Similarity -2077706329)


### M1 / ZIP-REPAIR

CX-R116 (Similarity -900544832), CX-R117 (Similarity -1745516950)


### H3 / DATA-INPUT

CX-R009 (Similarity -1401055022), CX-R010 (Similarity 314803386), CX-R011 (Similarity -763093660), CX-R012 (Similarity -576431034), CX-R013 (Similarity 936672265)


### H3 / SERVICE-CONFIG

CX-R014 (Similarity 280573853), CX-R017 (Similarity 333146485), CX-R018 (Similarity -1699090574), CX-R019 (Similarity 280573853), CX-R020 (Similarity 333146485), CX-R024 (Similarity 706674842), CX-R025 (Similarity -1651521647), CX-R026 (Similarity 1188711609), CX-R027 (Similarity 507564216), CX-R028 (Similarity -47860084), CX-R029 (Similarity -1477818700), CX-R030 (Similarity -56861040), CX-R031 (Similarity -1477818700), CX-R032 (Similarity -47860084), CX-R036 (Similarity -1372051600), CX-R037 (Similarity -315100207), CX-R038 (Similarity -1654445831)


### H3 / CATALOG-PATH

CX-R015 (Similarity 1291543521), CX-R016 (Similarity 1291543521), CX-R021 (Similarity 955143290), CX-R022 (Similarity -1947506698), CX-R023 (Similarity -1885568552)


### H3 / OFFICE-CONFIG

CX-R091 (Similarity 1144353721)


### H3 / TEST-ARTIFACT

CX-R033 (Similarity 1027859451), CX-R034 (Similarity -419043739), CX-R035 (Similarity 601440768), CX-R039 (Similarity 1760295004), CX-R040 (Similarity -90541740), CX-R041 (Similarity 1947077620), CX-R042 (Similarity 46418524), CX-R043 (Similarity -1299226345), CX-R044 (Similarity -1761994869), CX-R045 (Similarity -1720959708), CX-R046 (Similarity 462301936), CX-R047 (Similarity 1407586478), CX-R048 (Similarity 758400680), CX-R049 (Similarity -1517092728), CX-R050 (Similarity 2097746304), CX-R051 (Similarity 1237596672), CX-R052 (Similarity -152178112), CX-R053 (Similarity 934697572), CX-R054 (Similarity -1719932254), CX-R055 (Similarity -1345387132), CX-R056 (Similarity -86767479), CX-R057 (Similarity -695368680), CX-R058 (Similarity -1532911912), CX-R059 (Similarity -1429002613), CX-R060 (Similarity 2129211647), CX-R061 (Similarity -1429002613), CX-R062 (Similarity 1703303725), CX-R063 (Similarity 1552682661), CX-R064 (Similarity -1579543158), CX-R065 (Similarity -473184483), CX-R066 (Similarity -82668503), CX-R067 (Similarity 1966682696), CX-R068 (Similarity -485459168), CX-R069 (Similarity 361596616), CX-R070 (Similarity -756336288), CX-R071 (Similarity -1176590899), CX-R072 (Similarity -1657118664), CX-R073 (Similarity -753889073), CX-R074 (Similarity 711231073), CX-R075 (Similarity -2135577675), CX-R076 (Similarity 1335748697), CX-R077 (Similarity 745660673), CX-R078 (Similarity 920427747), CX-R079 (Similarity 2145587874), CX-R081 (Similarity -1426692475), CX-R082 (Similarity -1599453999), CX-R083 (Similarity -498998273), CX-R084 (Similarity 1235437117), CX-R085 (Similarity 1025575657), CX-R086 (Similarity -1153363084), CX-R087 (Similarity -620214048), CX-R088 (Similarity -1926027703), CX-R089 (Similarity -1549767052), CX-R090 (Similarity -1612927088), CX-R092 (Similarity 1398503197), CX-R093 (Similarity -457272103), CX-R094 (Similarity -293388255), CX-R095 (Similarity -771679709), CX-R096 (Similarity -661553133), CX-R097 (Similarity 1201223223), CX-R098 (Similarity 596844632), CX-R099 (Similarity 1018546989), CX-R100 (Similarity 596844632), CX-R101 (Similarity 596844632), CX-R102 (Similarity 341466547), CX-R103 (Similarity 596844632), CX-R104 (Similarity 596844632), CX-R105 (Similarity 1855789791), CX-R106 (Similarity -1574684448), CX-R107 (Similarity 360136018), CX-R108 (Similarity 596844632), CX-R109 (Similarity -2135358301), CX-R110 (Similarity 661522298), CX-R111 (Similarity 793367280), CX-R112 (Similarity 1739815605), CX-R113 (Similarity -1880250405), CX-R114 (Similarity -169465736), CX-R115 (Similarity 724762689)


SCS aggregate: High/open, original ID unavailable; H0 blocks exact remediation. AR-01→H4, AR-02→M2, AR-03→L1 remain separately identified review findings.

**STOP: do not begin Run 2 until the human owner explicitly approves it.**
