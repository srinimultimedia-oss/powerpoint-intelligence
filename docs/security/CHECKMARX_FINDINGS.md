# CHECKMARX FINDINGS — Run 1

**Later implementation note (2 November 2026):** the owner subsequently approved feature-preserving fixes. This file retains the original Run 1 assessment; see [SECURITY_REMEDIATION_STATUS.md](SECURITY_REMEDIATION_STATUS.md) for the four locally addressed records, evidence and remaining gates. Original Checkmarx states have not been changed and no rescan result is asserted.

Assessment date: **2 November 2026, Asia/Dubai**. Assessment/planning only. **No finding is resolved. No application code, dependency, trust policy, scanner disposition or suppression was changed. Run 2 requires explicit owner approval.**

## Provenance, discovery and scope

Authoritative report: `C:/Users/srini/Downloads/scan report.md`, raw SHA-256 **82ecb7f276330342f20da4518bb9b853650bc30b973ce11a6d603addb4a62ff9**. All 5,394 logical lines were processed, including every query record, overview, category table, vulnerability-description appendix and the final SCS table. Repository/accessible Codex Markdown searches (including generated directories, excluding `.git`, dependency caches, `node_modules`, `bin` and `obj`) found no second Checkmarx Scan Report. This explicitly supplied report was identified by its contents, not its filename.

Scan ID: **bed87600-7c2c-45b7-8b17-15b0b268ec82**, last scanned **2 November 2026 09:21 Asia/Dubai**, report created 09:27. Full Scan / ASA Premium, source origin **ZIP**, branch **N/A**. Project is titled `Srinivas Kikkuru`; its main-branch label `visa cancellation checklist v1.0` does not establish repository identity or a commit. Scope: 323 files (SAST 297, IaC 26), 25,408 LOC (22,858 / 2,550). All six scanner types are marked Completed.

Review checkout: `C:/Srini/powerpoint-intelligence/powerpoint-intelligence`, local HEAD **1f3259c15caf5affba44df37d9b43f2d6f00d205** plus reviewed working files. Every one of the **234** reported source/destination snippets matches the current file and reported line. Report prefix `/powerpoint-intelligence-main/` maps to this root. This proves endpoint content correspondence, not equivalence of the complete scanned ZIP to this commit. Source/configuration/dependency-file hashes were captured for unchanged-source verification.

## Counts and report discrepancies

| Item | Count / scope |
|---|---|
| Original reported total | **118** |
| Critical | **0** |
| High | **116** |
| Medium | **2** |
| Low / Information | Excluded by report filters; **0 included**, whole-scan totals unknown |
| Detailed records | **117**, comprising 115 High + 2 Medium |
| Summary-only SCS | **1 High**, engine `2ms`, no individual ID/location/trace supplied |
| Confirmed weakness records | **4**, representing **2 root causes**: corpus output-name traversal and repair resource exhaustion |
| Conditional risk records | **3**, executable/toolchain provenance dependent on an unestablished actor boundary |
| Potential false-positive records | **110**, tentative source-based assessments, not approved scanner dispositions |
| Blocked entries | **1**, summary-only SCS; raw query taxonomy reconciliation is also outstanding |
| Remediated / verified fixed | **0 / 0** |

Detailed groups: 3 Download of Code Without Integrity Check (CWE-494), 5 Stored_Path_Traversal (CWE-22), 107 Path_Traversal (CWE-22), 2 Unchecked_Input_for_Loop_Condition (CWE-606). The overview instead lists **2** downloads and **1 Inclusion of Functionality from Untrusted Control Sphere**; no independently identifiable detail row exists for the latter. The third detailed download is vswhere invocation. **Do not silently relabel its CWE, merge it with SCS or invent a 119th finding.** Obtain original result/query/engine export to reconcile this.

All 117 detailed result URLs are distinct. Similarity IDs are not unique: 280573853, 1291543521, 333146485, -47860084, -1477818700 and -1429002613 each occur twice; 596844632 occurs six times. Every record is retained. `CX-Rnnn` is an assessment alias, not an original scanner ID; original Similarity ID, query ordinal and result URL remain attached. Original detailed Status=New, State=To Verify, found first/last 2 November 2026, first scan equal to the scan above. No CVE IDs are supplied. Zero filtered SCA results do not establish CVE-free dependencies or complete SCA coverage.

“Confirmed” means the static primitive/control gap is verified, not an executed exploit or production reachability. “Potential false positive” needs threat-model/reviewer confirmation: developer CLI paths intentionally chosen by the same user are not inherently a desktop privilege boundary. “Conditional risk” requires a more-trusted caller/artifact boundary not demonstrated in the report. Original scanner severity is preserved independently of these assessments.

No malicious deck, build/test executable or Office automation was run for this assessment; no package was downloaded/upgraded, credential exposed, trust changed or scanner exclusion introduced. The corpus witness was a side-effect-free path calculation only.

## Every detailed finding


### CX-R001 — Download of Code Without Integrity Check — Result 1 of 3

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **333933924**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=%2BAF5VlQ5BgcEk%2BqbUNsAC0P2HI8%3D); report line 326 |
| Original severity / CWE / CVE | **High** / **CWE-494** / no CVE supplied |
| Component | PIIDX developer assembly cache |
| Source | `scripts/Piidx-TestSuite.ps1:11` — function Import-PpiPiidxSuite {, `$prior` |
| Sink | `scripts/Piidx-TestSuite.ps1:26` — function Import-PpiPiidxSuite {, `$dll` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Conditional risk**, open / not remediated |
| Plan mapping | **H2**, profile `EXEC-PIIDX` |

**Original description:** No query description available

**Observed source → sink:**

```text
$prior=$null;if(Test-Path -LiteralPath $receipt){$prior=Get-Content -LiteralPath $receipt -Raw | ConvertFrom-Json}
→ $null=[Reflection.Assembly]::LoadFrom($dll)
```

**Root cause / controls:** Piidx-TestSuite.ps1:12-15 already compares cached DLL and source hashes against its receipt. There is no network download in this reported path. The receipt and executable assembly share a mutable developer cache; independent provenance is not established by a colocated checksum.

**Exploitability / business impact:** Requires an attacker who can replace both DLL and receipt, or win a verification/load race, across a more-trusted build/CI boundary. That boundary is not demonstrated. DLL-only or receipt-only corruption is rejected. Conditional impact: code execution with the runner token; no elevation proved.

**Recommendation:** Preserve fail-closed source/hash/missing-file controls. Establish cache ownership/ACL and actor model; authenticate shared artifacts against independent provenance, validate physical path ancestry and minimize verify/load races. Another adjacent hash is not independent authentication.

**Regression risk:** Stable DLL names, reviewed recovery, source drift detection and fresh-session assembly loading must remain intact.

**Required validation:** DLL-only, receipt-only and combined tamper; missing/unreceipted files; changed sources/toolchain; concurrent replacement; approved Windows PIIDX execution with no security-policy changes.


### CX-R002 — Download of Code Without Integrity Check — Result 2 of 3

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-810172859**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=g91dl7ETsuHNIeAfUDetKfCH1ms%3D); report line 367 |
| Original severity / CWE / CVE | **High** / **CWE-494** / no CVE supplied |
| Component | Scene runner provenance and execution |
| Source | `scripts/Scene-RunnerCache.ps1:18` — function Get-PpiSceneCacheState([string]$ReceiptPath,[string]$BuildKey,[System.Collections.IDictionary]$Files) {, `$Files[$name]` |
| Sink | `scripts/Test-SceneContracts.ps1:58` — script scope, `$executable` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Conditional risk**, open / not remediated |
| Plan mapping | **H2**, profile `EXEC-SCENE` |

**Original description:** No query description available

**Observed source → sink:**

```text
if(!(Test-Path -LiteralPath $Files[$name])){throw 'A cached test runner file is missing. Check scanner evidence; no automatic restoration or rebuild.'}
→ & $executable tests/s3/fixtures/all-primitives.v1.json $roundtrip $corpus 2>&1 | Tee-Object -FilePath (Join-Path $results 'execution.log')
```

**Root cause / controls:** Scene-RunnerCache.ps1:15-20 verifies the build key and every binary/dependency SHA-256 before Test-SceneContracts.ps1 executes the runner. Corpus generation separates that check and launch. The residual question is a mutable shared cache/race, not absence of all integrity checks.

**Exploitability / business impact:** An attacker needs write access to the runner/receipt during a trusted invocation. Same-owner artifacts alone do not establish an exploit. Conditional impact is code substitution in the runner context.

**Recommendation:** Keep all hash/scanner/legacy-audit gates. Define protected cache ownership, revalidate immediately before execution and use authenticated immutable artifacts if lower-trust shared caches are supported. Final hashing alone cannot solve a writable-directory race.

**Regression risk:** Never rename stable test executables, restore quarantine content, adopt old receipts implicitly or downgrade execution checks to compile checks.

**Required validation:** Retain all existing cache/review tests; add last-moment executable/dependency swaps, untrusted RunnerRoot, modified receipt and junction-parent cases. Verify legitimate x86/x64 execution.


### CX-R003 — Download of Code Without Integrity Check — Result 3 of 3

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **515670905**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=MVbT4JHiROszlNYb4utc2GriBRA%3D); report line 408 |
| Original severity / CWE / CVE | **High** / **CWE-494** / no CVE supplied |
| Component | easySetup toolchain discovery |
| Source | `easySetup/Build-EasySetup.ps1:9` — script scope, `$vswhere` |
| Sink | `easySetup/Build-EasySetup.ps1:10` — script scope, `$vswhere` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Conditional risk**, open / not remediated |
| Plan mapping | **H2**, profile `EXEC-TOOL` |

**Original description:** No query description available

**Observed source → sink:**

```text
$vswhere=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
→ $vs=& $vswhere -latest -products '*' -property installationPath
```

**Root cause / controls:** The ProgramFiles(x86) environment value determines vswhere.exe; Build-EasySetup.ps1 executes it without an independent canonical-root/Microsoft publisher check. No download occurs at this source/sink.

**Exploitability / business impact:** Requires lower-trust control of the environment/tool location while a more trusted build token runs it. A developer selecting their own toolchain is not independently an exploit. If the boundary exists, malicious tooling can compromise signed output.

**Recommendation:** Use approved OS installation roots/toolchain configuration, canonical executable paths and expected publisher validation for vswhere, MSBuild and compiler. Reconcile the report query taxonomy using the raw export; do not silently relabel original CWE-494 as CWE-829.

**Regression risk:** Support approved nondefault VS/CI installations; do not hardcode this workstation or modify certificate trust automatically.

**Required validation:** Spoofed environment, unsigned/wrong-publisher tool, valid nondefault VS and both signed architecture payloads.


### CX-R004 — Stored_Path_Traversal — Result 1 of 5

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1305902851**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=BmsXsxXaff6KgcaEcGsdk5oV0%2Fw%3D); report line 469 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Library configuration and legacy migration |
| Source | `src/Library.Core/LibraryConfiguration.cs:49` — Load, `ReadAllText` |
| Sink | `src/Library.Core/PiidxCatalog.cs:59` — Register, `Exists` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `CONFIG-ROOT` |

**Original description:** Method Load at line 49 of /powerpoint-intelligence-main/src/Library.Core/LibraryConfiguration.cs gets dynamic data from the ReadAllText element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Register at line 59 of /powerpoint-intelligence-main/src/Library.Core/PiidxCatalog.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
result.Sources.AddRange(ParseSources(File.ReadAllText(file),Path.GetDirectoryName(file)));
→ root=Root(root);if(!Directory.Exists(root))throw new DirectoryNotFoundException("Library root is unavailable.");
```

**Root cause / controls:** LibraryConfiguration.PathValue/LegacyFolder canonicalize a configured source root. PiidxCatalog.Register:59-60 checks directory existence and rejects a root that is itself a reparse point. Selecting external source directories is intended; these reported sinks probe metadata, not arbitrary file content.

**Exploitability / business impact:** No escape from an already authorized source root or file overwrite is demonstrated. A hostile configuration or UNC/reparse policy violation would require a separately established trust boundary. Configuration ACLs and enterprise source policy are not known.

**Recommendation:** Keep external roots and migration. Separate authorization to register a source from derived-path containment. Verify config provenance and approved network/reparse policy. Do not force all input libraries under the application directory.

**Regression risk:** Blanket restrictions break external libraries, network shares, relocation and preserved legacy registrations.

**Required validation:** Authorized external/relative roots, duplicates, missing roots, leaf/ancestor links, approved network roots, migration interruption and preservation of prior catalog/registration.


### CX-R005 — Stored_Path_Traversal — Result 2 of 5

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **531211230**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=EN8FLgdDeBmuWaPQImmuPEtmWcs%3D); report line 510 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Library configuration and legacy migration |
| Source | `src/Library.Core/LibraryConfiguration.cs:83` — MigrateLegacy, `ReadAllText` |
| Sink | `src/Library.Core/PiidxCatalog.cs:60` — Register, `GetAttributes` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `CONFIG-ROOT` |

**Original description:** Method MigrateLegacy at line 83 of /powerpoint-intelligence-main/src/Library.Core/LibraryConfiguration.cs gets dynamic data from the ReadAllText element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Register at line 60 of /powerpoint-intelligence-main/src/Library.Core/PiidxCatalog.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
string root=LegacyFolder(File.ReadAllText(LegacyCachePath));if(!Directory.Exists(root))throw new DirectoryNotFoundException("Legacy library source folder is unavailable; original registration preserved.");
→ if((File.GetAttributes(root)&FileAttributes.ReparsePoint)!=0)throw new InvalidDataException("Register a direct library directory.");
```

**Root cause / controls:** LibraryConfiguration.PathValue/LegacyFolder canonicalize a configured source root. PiidxCatalog.Register:59-60 checks directory existence and rejects a root that is itself a reparse point. Selecting external source directories is intended; these reported sinks probe metadata, not arbitrary file content.

**Exploitability / business impact:** No escape from an already authorized source root or file overwrite is demonstrated. A hostile configuration or UNC/reparse policy violation would require a separately established trust boundary. Configuration ACLs and enterprise source policy are not known.

**Recommendation:** Keep external roots and migration. Separate authorization to register a source from derived-path containment. Verify config provenance and approved network/reparse policy. Do not force all input libraries under the application directory.

**Regression risk:** Blanket restrictions break external libraries, network shares, relocation and preserved legacy registrations.

**Required validation:** Authorized external/relative roots, duplicates, missing roots, leaf/ancestor links, approved network roots, migration interruption and preservation of prior catalog/registration.


### CX-R006 — Stored_Path_Traversal — Result 3 of 5

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1886406541**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=Ji%2FFpjiCThgyQwqBiu1jvvnYvMg%3D); report line 551 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Library configuration and legacy migration |
| Source | `src/Library.Core/LibraryConfiguration.cs:83` — MigrateLegacy, `ReadAllText` |
| Sink | `src/Library.Core/PiidxCatalog.cs:59` — Register, `Exists` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `CONFIG-ROOT` |

**Original description:** Method MigrateLegacy at line 83 of /powerpoint-intelligence-main/src/Library.Core/LibraryConfiguration.cs gets dynamic data from the ReadAllText element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Register at line 59 of /powerpoint-intelligence-main/src/Library.Core/PiidxCatalog.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
string root=LegacyFolder(File.ReadAllText(LegacyCachePath));if(!Directory.Exists(root))throw new DirectoryNotFoundException("Legacy library source folder is unavailable; original registration preserved.");
→ root=Root(root);if(!Directory.Exists(root))throw new DirectoryNotFoundException("Library root is unavailable.");
```

**Root cause / controls:** LibraryConfiguration.PathValue/LegacyFolder canonicalize a configured source root. PiidxCatalog.Register:59-60 checks directory existence and rejects a root that is itself a reparse point. Selecting external source directories is intended; these reported sinks probe metadata, not arbitrary file content.

**Exploitability / business impact:** No escape from an already authorized source root or file overwrite is demonstrated. A hostile configuration or UNC/reparse policy violation would require a separately established trust boundary. Configuration ACLs and enterprise source policy are not known.

**Recommendation:** Keep external roots and migration. Separate authorization to register a source from derived-path containment. Verify config provenance and approved network/reparse policy. Do not force all input libraries under the application directory.

**Regression risk:** Blanket restrictions break external libraries, network shares, relocation and preserved legacy registrations.

**Required validation:** Authorized external/relative roots, duplicates, missing roots, leaf/ancestor links, approved network roots, migration interruption and preservation of prior catalog/registration.


### CX-R007 — Stored_Path_Traversal — Result 4 of 5

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **889475886**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=gaaduQJpGVLLRv2qGkcFheVwEf0%3D); report line 592 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Scene corpus artifact filename |
| Source | `tests/s3/SceneContractTests.cs:67` — Main, `ReadAllText` |
| Sink | `tests/s3/SceneContractTests.cs:72` — Lambda, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Confirmed weakness**, open / not remediated |
| Plan mapping | **H1**, profile `CORPUS-NAME` |

**Original description:** Method Main at line 67 of /powerpoint-intelligence-main/tests/s3/SceneContractTests.cs gets dynamic data from the ReadAllText element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Lambda at line 72 of /powerpoint-intelligence-main/tests/s3/SceneContractTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
foreach (var item in JArray.Parse(File.ReadAllText(Path.Combine(corpus, "cases.json"))))
→ if (expected) { model = SceneCodec.Decode(json); File.WriteAllText(Path.Combine(output, name + ".v1.json"), SceneCodec.Encode(model)); }
```

**Root cause / controls:** SceneContractTests.cs:67-72 reads item["name"] from cases.json and combines output with name+".v1.json" without a basename/containment check. This data-derived child path differs from an intentional CLI output root.

**Exploitability / business impact:** A malicious corpus name such as ../escaped can write outside the chosen output directory when its scene JSON is valid. Side-effect-free Windows path evaluation yields C:\escaped.v1.json from C:\owned-output plus ..\escaped.v1.json. This is a verified static write primitive, not an executed exploit. Impact is runner-writable file corruption; production exposure/elevation are not demonstrated. Two reported records share this sink.

**Recommendation:** Reject invalid bounded basenames, rooted/drive/UNC/extended paths, separators, dot segments, ADS/device and trailing-dot/space forms. Enforce separator-aware canonical containment in an owned output root and no unowned-file overwrite. Reject rather than sanitize collisions; retain valid current fixture names.

**Regression risk:** C#/Node fixture naming and roundtrip validators must remain synchronized. Silent renaming may hide/collide cases.

**Required validation:** All current valid corpus names plus parent traversal, absolute/drive-relative/UNC, sibling-prefix, ADS/devices, duplicates, existing targets and junction parents. Verify no outside creation/overwrite and all valid cross-language roundtrips.


### CX-R008 — Stored_Path_Traversal — Result 5 of 5

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-415229814**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=OWTVboggHA76LM%2Bp56U1nQsRfew%3D); report line 633 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Library file discovery |
| Source | `src/Library.Core/PresentationSource.cs:64` — Discover, `EnumerateFiles` |
| Sink | `src/Library.Core/PresentationSource.cs:67` — Discover, `GetAttributes` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `DISCOVERY` |

**Original description:** Method Discover at line 64 of /powerpoint-intelligence-main/src/Library.Core/PresentationSource.cs gets dynamic data from the EnumerateFiles element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Discover at line 67 of /powerpoint-intelligence-main/src/Library.Core/PresentationSource.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
foreach(string file in Directory.EnumerateFiles(folder))
→ if((File.GetAttributes(file)&FileAttributes.ReparsePoint)!=0)continue;
```

**Root cause / controls:** PresentationSource.Discover obtains child names from Directory.EnumerateFiles(folder), rather than appending stored attacker names. It canonicalizes the root, skips reparse leaves, bounds files/directories and checks cancellation.

**Exploitability / business impact:** GetAttributes does not itself read contents or escape the enumerated directory. Concurrent ancestor replacement is a separate conditional race (AR-02), not proof of this reported stored traversal.

**Recommendation:** Preserve bounded source discovery; establish approved physical-root/handle identity policy when hostile shared folders are supported. Revalidate the source at open, not only enumeration.

**Regression risk:** Overbroad junction/UNC rejection can break redirected profiles and managed libraries.

**Required validation:** Direct children, skipped links, cancellation, directory/file limits, source mutation and disposable Windows ancestor/race tests; unavailable tests remain blocked.


### CX-R009 — Path_Traversal — Result 1 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1401055022**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=SlXDUirAGs0QUBO9FnqwYV1Ybok%3D); report line 697 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Explicit linked-data input |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `apps/powerpoint-vsto/ExcelRangeMapping.cs:57` — WithWorkbook, `Exists` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `DATA-INPUT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in WithWorkbook at line 57 of /powerpoint-intelligence-main/apps/powerpoint-vsto/ExcelRangeMapping.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ var file=new FileInfo(path);if(!file.Exists||file.Length>20*1024*1024)throw new InvalidDataException("The source is missing or exceeds 20 MB. Explicitly relink it.");
```

**Root cause / controls:** The reported source is a test output argument. DataSourceTests/DemoScenario create fixtures there and intentionally load them. Production Create/Validate requires canonical rooted paths and approved extensions; Load limits size and mapped Excel validates range/columns. No fixed input sandbox is promised.

**Exploitability / business impact:** An explicitly selected user file is intended desktop behavior, not itself disclosure/traversal. Presentation-carried source references raise separate consent/network questions not established by this CLI-fixture path. No privilege boundary or exfiltration is demonstrated.

**Recommendation:** Keep picking, explicit mapping, moved-source refusal and confirmed relinking. Define ExternalSource authority separately from OwnedOutput; validate at use and apply an approved consent/network policy for document-carried references. Do not ban supported external files.

**Regression risk:** Preserve all four formats, Unicode paths, non-A1 mappings, native edits and save/reopen.

**Required validation:** Excel/CSV/JSON/XML, size/type/schema failures, missing/moved source without fallback, header-change refusal, relinking, document-carried path policy and native mapped refresh/reopen with no unintended network probe.


### CX-R010 — Path_Traversal — Result 2 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **314803386**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=E8U812I3JbLBCSES8%2BWsmkgSqCs%3D); report line 738 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Explicit linked-data input |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `apps/powerpoint-vsto/DataSources.cs:95` — Load, `Exists` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `DATA-INPUT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Load at line 95 of /powerpoint-intelligence-main/apps/powerpoint-vsto/DataSources.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ var info=new FileInfo(link.Path);if(!info.Exists||info.Length>20*1024*1024)throw new InvalidDataException("The linked source is missing or exceeds 20 MB. Relink it using Data Source.");
```

**Root cause / controls:** The reported source is a test output argument. DataSourceTests/DemoScenario create fixtures there and intentionally load them. Production Create/Validate requires canonical rooted paths and approved extensions; Load limits size and mapped Excel validates range/columns. No fixed input sandbox is promised.

**Exploitability / business impact:** An explicitly selected user file is intended desktop behavior, not itself disclosure/traversal. Presentation-carried source references raise separate consent/network questions not established by this CLI-fixture path. No privilege boundary or exfiltration is demonstrated.

**Recommendation:** Keep picking, explicit mapping, moved-source refusal and confirmed relinking. Define ExternalSource authority separately from OwnedOutput; validate at use and apply an approved consent/network policy for document-carried references. Do not ban supported external files.

**Regression risk:** Preserve all four formats, Unicode paths, non-A1 mappings, native edits and save/reopen.

**Required validation:** Excel/CSV/JSON/XML, size/type/schema failures, missing/moved source without fallback, header-change refusal, relinking, document-carried path policy and native mapped refresh/reopen with no unintended network probe.


### CX-R011 — Path_Traversal — Result 3 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-763093660**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=lSQxym12i2teZcSyuQPiNoRAp%2BI%3D); report line 779 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Explicit linked-data input |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `apps/powerpoint-vsto/DataSources.cs:95` — Load, `Exists` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `DATA-INPUT` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Load at line 95 of /powerpoint-intelligence-main/apps/powerpoint-vsto/DataSources.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ var info=new FileInfo(link.Path);if(!info.Exists||info.Length>20*1024*1024)throw new InvalidDataException("The linked source is missing or exceeds 20 MB. Relink it using Data Source.");
```

**Root cause / controls:** The reported source is a test output argument. DataSourceTests/DemoScenario create fixtures there and intentionally load them. Production Create/Validate requires canonical rooted paths and approved extensions; Load limits size and mapped Excel validates range/columns. No fixed input sandbox is promised.

**Exploitability / business impact:** An explicitly selected user file is intended desktop behavior, not itself disclosure/traversal. Presentation-carried source references raise separate consent/network questions not established by this CLI-fixture path. No privilege boundary or exfiltration is demonstrated.

**Recommendation:** Keep picking, explicit mapping, moved-source refusal and confirmed relinking. Define ExternalSource authority separately from OwnedOutput; validate at use and apply an approved consent/network policy for document-carried references. Do not ban supported external files.

**Regression risk:** Preserve all four formats, Unicode paths, non-A1 mappings, native edits and save/reopen.

**Required validation:** Excel/CSV/JSON/XML, size/type/schema failures, missing/moved source without fallback, header-change refusal, relinking, document-carried path policy and native mapped refresh/reopen with no unintended network probe.


### CX-R012 — Path_Traversal — Result 4 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-576431034**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=WAvFdTyXfcQjiTSQEkOV3migVkg%3D); report line 820 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Explicit linked-data input |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `apps/powerpoint-vsto/DataSources.cs:97` — Load, `ReadAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `DATA-INPUT` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Load at line 97 of /powerpoint-intelligence-main/apps/powerpoint-vsto/DataSources.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ string text=File.ReadAllText(link.Path);
```

**Root cause / controls:** The reported source is a test output argument. DataSourceTests/DemoScenario create fixtures there and intentionally load them. Production Create/Validate requires canonical rooted paths and approved extensions; Load limits size and mapped Excel validates range/columns. No fixed input sandbox is promised.

**Exploitability / business impact:** An explicitly selected user file is intended desktop behavior, not itself disclosure/traversal. Presentation-carried source references raise separate consent/network questions not established by this CLI-fixture path. No privilege boundary or exfiltration is demonstrated.

**Recommendation:** Keep picking, explicit mapping, moved-source refusal and confirmed relinking. Define ExternalSource authority separately from OwnedOutput; validate at use and apply an approved consent/network policy for document-carried references. Do not ban supported external files.

**Regression risk:** Preserve all four formats, Unicode paths, non-A1 mappings, native edits and save/reopen.

**Required validation:** Excel/CSV/JSON/XML, size/type/schema failures, missing/moved source without fallback, header-change refusal, relinking, document-carried path policy and native mapped refresh/reopen with no unintended network probe.


### CX-R013 — Path_Traversal — Result 5 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **936672265**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=wpEpE33%2FztzhfYCL4nk3Gb%2Fe9yU%3D); report line 861 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Explicit linked-data input |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `apps/powerpoint-vsto/DemoLinks.cs:51` — Transfer, `Exists` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `DATA-INPUT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Transfer at line 51 of /powerpoint-intelligence-main/apps/powerpoint-vsto/DemoLinks.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ var info=new FileInfo(path);if(!info.Exists||!String.Equals(info.Extension,".xlsx",StringComparison.OrdinalIgnoreCase)||info.Length>20*1024*1024)throw new InvalidDataException("Choose an existing .xlsx file under 20 MB.");
```

**Root cause / controls:** The reported source is a test output argument. DataSourceTests/DemoScenario create fixtures there and intentionally load them. Production Create/Validate requires canonical rooted paths and approved extensions; Load limits size and mapped Excel validates range/columns. No fixed input sandbox is promised.

**Exploitability / business impact:** An explicitly selected user file is intended desktop behavior, not itself disclosure/traversal. Presentation-carried source references raise separate consent/network questions not established by this CLI-fixture path. No privilege boundary or exfiltration is demonstrated.

**Recommendation:** Keep picking, explicit mapping, moved-source refusal and confirmed relinking. Define ExternalSource authority separately from OwnedOutput; validate at use and apply an approved consent/network policy for document-carried references. Do not ban supported external files.

**Regression risk:** Preserve all four formats, Unicode paths, non-A1 mappings, native edits and save/reopen.

**Required validation:** Excel/CSV/JSON/XML, size/type/schema failures, missing/moved source without fallback, header-change refusal, relinking, document-carried path policy and native mapped refresh/reopen with no unintended network probe.


### CX-R014 — Path_Traversal — Result 6 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **280573853**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=ALSofsJKSenebk0GQG9v41BKwyI%3D); report line 902 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Nonsecret service configuration |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `apps/powerpoint-vsto/ServiceConnections.cs:95` — Save, `Delete` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `SERVICE-CONFIG` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Save at line 95 of /powerpoint-intelligence-main/apps/powerpoint-vsto/ServiceConnections.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ try{File.WriteAllText(temp,xml,new UTF8Encoding(false,true));if(File.Exists(target))File.Replace(temp,target,null);else File.Move(temp,target);}finally{if(File.Exists(temp))File.Delete(temp);}
```

**Root cause / controls:** The path originates in CLI test arguments. ServiceConnectionTests makes a GUID-named child and constant connections.xml. Production UI uses the fixed per-user DefaultPath; Save uses a random sibling temp and atomic Replace/Move. The Delete sink cleans that owned temp. Strict schema/HTTPS/size validation already exists.

**Exploitability / business impact:** No untrusted production caller selects Save(path) in the reviewed call graph. Replace/delete/existence operations are ordinary owned persistence. A separate ownership/reparse issue or future API would need analysis. No live HTTP/auth bypass/credential storage exists here.

**Recommendation:** Keep nonsecret-only strict metadata, atomic saves and cancellation. Encapsulate fixed production store separately from test-only path injection or use an explicitly owned directory context. Handle ancestry under approved policy; do not invent connectivity/authentication success.

**Regression risk:** Atomic replacement, environment-independent fixtures and preservation after invalid save/cancel must remain; do not delete user config or send new traffic.

**Required validation:** Existing strict service cases, roundtrip, invalid save preserves bytes, temp cleanup/write failure, fixed UI destination, unsafe parent/junction and no plaintext credentials.


### CX-R015 — Path_Traversal — Result 7 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **1291543521**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=KfOWj7OfedjLABQ40B1BQ5H5P50%3D); report line 943 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | PIIDX index and writer lock |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `src/Library.Core/PiidxCatalog.cs:43` — WriterLock, `Exists` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `CATALOG-PATH` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in WriterLock at line 43 of /powerpoint-intelligence-main/src/Library.Core/PiidxCatalog.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ if(File.Exists(path)&&(File.GetAttributes(path)&FileAttributes.ReparsePoint)!=0)throw new InvalidDataException("Index writer lock cannot be a reparse point.");
```

**Root cause / controls:** A test workspace/configured .piidx path reaches index/lock probes. The catalog canonicalizes its path, restricts .piidx, checks leaf reparse points and uses an exclusive bounded-wait writer lock. Selecting a different authorized index is intended.

**Exploitability / business impact:** A developer-selected fixture path does not establish remote arbitrary-file access. A hostile writable config/ancestor is a separate conditional threat (AR-02); no content-derived lock filename appears on the reported path.

**Recommendation:** Preserve configured index placement, single-writer/cancellation and read-only source roots. Define owned catalog authority and physical ancestry/handle checks where hostile locations are in scope; never silently migrate/create a rejected database.

**Regression risk:** Migration, journals, concurrency, cancellation, readonly libraries and redirected LocalAppData must remain supported.

**Required validation:** Schema/digest/tamper/corpus tests, lock timeout/concurrency, corrupt index preservation, valid custom paths, ancestor/alias/network policy and no partial catalog updates.


### CX-R016 — Path_Traversal — Result 8 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **1291543521**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=vqoPH574mShX9Wvf6IirQyZ71XQ%3D); report line 984 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | PIIDX index and writer lock |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `src/Library.Core/PiidxCatalog.cs:43` — WriterLock, `GetAttributes` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `CATALOG-PATH` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in WriterLock at line 43 of /powerpoint-intelligence-main/src/Library.Core/PiidxCatalog.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ if(File.Exists(path)&&(File.GetAttributes(path)&FileAttributes.ReparsePoint)!=0)throw new InvalidDataException("Index writer lock cannot be a reparse point.");
```

**Root cause / controls:** A test workspace/configured .piidx path reaches index/lock probes. The catalog canonicalizes its path, restricts .piidx, checks leaf reparse points and uses an exclusive bounded-wait writer lock. Selecting a different authorized index is intended.

**Exploitability / business impact:** A developer-selected fixture path does not establish remote arbitrary-file access. A hostile writable config/ancestor is a separate conditional threat (AR-02); no content-derived lock filename appears on the reported path.

**Recommendation:** Preserve configured index placement, single-writer/cancellation and read-only source roots. Define owned catalog authority and physical ancestry/handle checks where hostile locations are in scope; never silently migrate/create a rejected database.

**Regression risk:** Migration, journals, concurrency, cancellation, readonly libraries and redirected LocalAppData must remain supported.

**Required validation:** Schema/digest/tamper/corpus tests, lock timeout/concurrency, corrupt index preservation, valid custom paths, ancestor/alias/network policy and no partial catalog updates.


### CX-R017 — Path_Traversal — Result 9 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **333146485**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=1MCKQ9paYIKIa1ch8rvS2c9YTrU%3D); report line 1025 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Nonsecret service configuration |
| Source | `tests/s0/ServiceConnectionTests.cs:9` — Main, `args` |
| Sink | `apps/powerpoint-vsto/ServiceConnections.cs:95` — Save, `Delete` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `SERVICE-CONFIG` |

**Original description:** Method Main at line 9 of /powerpoint-intelligence-main/tests/s0/ServiceConnectionTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Save at line 95 of /powerpoint-intelligence-main/apps/powerpoint-vsto/ServiceConnections.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
[STAThread] public static int Main(string[] args){try{Run(args.Length==0?null:args[0]);return 0;}catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
→ try{File.WriteAllText(temp,xml,new UTF8Encoding(false,true));if(File.Exists(target))File.Replace(temp,target,null);else File.Move(temp,target);}finally{if(File.Exists(temp))File.Delete(temp);}
```

**Root cause / controls:** The path originates in CLI test arguments. ServiceConnectionTests makes a GUID-named child and constant connections.xml. Production UI uses the fixed per-user DefaultPath; Save uses a random sibling temp and atomic Replace/Move. The Delete sink cleans that owned temp. Strict schema/HTTPS/size validation already exists.

**Exploitability / business impact:** No untrusted production caller selects Save(path) in the reviewed call graph. Replace/delete/existence operations are ordinary owned persistence. A separate ownership/reparse issue or future API would need analysis. No live HTTP/auth bypass/credential storage exists here.

**Recommendation:** Keep nonsecret-only strict metadata, atomic saves and cancellation. Encapsulate fixed production store separately from test-only path injection or use an explicitly owned directory context. Handle ancestry under approved policy; do not invent connectivity/authentication success.

**Regression risk:** Atomic replacement, environment-independent fixtures and preservation after invalid save/cancel must remain; do not delete user config or send new traffic.

**Required validation:** Existing strict service cases, roundtrip, invalid save preserves bytes, temp cleanup/write failure, fixed UI destination, unsafe parent/junction and no plaintext credentials.


### CX-R018 — Path_Traversal — Result 10 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1699090574**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=jj%2Bo%2Fz1NNYBEpYfc9s806DLD%2BIg%3D); report line 1066 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Nonsecret service configuration |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `apps/powerpoint-vsto/ServiceConnections.cs:95` — Save, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `SERVICE-CONFIG` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Save at line 95 of /powerpoint-intelligence-main/apps/powerpoint-vsto/ServiceConnections.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ try{File.WriteAllText(temp,xml,new UTF8Encoding(false,true));if(File.Exists(target))File.Replace(temp,target,null);else File.Move(temp,target);}finally{if(File.Exists(temp))File.Delete(temp);}
```

**Root cause / controls:** The path originates in CLI test arguments. ServiceConnectionTests makes a GUID-named child and constant connections.xml. Production UI uses the fixed per-user DefaultPath; Save uses a random sibling temp and atomic Replace/Move. The Delete sink cleans that owned temp. Strict schema/HTTPS/size validation already exists.

**Exploitability / business impact:** No untrusted production caller selects Save(path) in the reviewed call graph. Replace/delete/existence operations are ordinary owned persistence. A separate ownership/reparse issue or future API would need analysis. No live HTTP/auth bypass/credential storage exists here.

**Recommendation:** Keep nonsecret-only strict metadata, atomic saves and cancellation. Encapsulate fixed production store separately from test-only path injection or use an explicitly owned directory context. Handle ancestry under approved policy; do not invent connectivity/authentication success.

**Regression risk:** Atomic replacement, environment-independent fixtures and preservation after invalid save/cancel must remain; do not delete user config or send new traffic.

**Required validation:** Existing strict service cases, roundtrip, invalid save preserves bytes, temp cleanup/write failure, fixed UI destination, unsafe parent/junction and no plaintext credentials.


### CX-R019 — Path_Traversal — Result 11 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **280573853**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=l6ThIyGxJKTuohxmThFQAzRQ7sI%3D); report line 1107 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Nonsecret service configuration |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `apps/powerpoint-vsto/ServiceConnections.cs:95` — Save, `Exists` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `SERVICE-CONFIG` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Save at line 95 of /powerpoint-intelligence-main/apps/powerpoint-vsto/ServiceConnections.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ try{File.WriteAllText(temp,xml,new UTF8Encoding(false,true));if(File.Exists(target))File.Replace(temp,target,null);else File.Move(temp,target);}finally{if(File.Exists(temp))File.Delete(temp);}
```

**Root cause / controls:** The path originates in CLI test arguments. ServiceConnectionTests makes a GUID-named child and constant connections.xml. Production UI uses the fixed per-user DefaultPath; Save uses a random sibling temp and atomic Replace/Move. The Delete sink cleans that owned temp. Strict schema/HTTPS/size validation already exists.

**Exploitability / business impact:** No untrusted production caller selects Save(path) in the reviewed call graph. Replace/delete/existence operations are ordinary owned persistence. A separate ownership/reparse issue or future API would need analysis. No live HTTP/auth bypass/credential storage exists here.

**Recommendation:** Keep nonsecret-only strict metadata, atomic saves and cancellation. Encapsulate fixed production store separately from test-only path injection or use an explicitly owned directory context. Handle ancestry under approved policy; do not invent connectivity/authentication success.

**Regression risk:** Atomic replacement, environment-independent fixtures and preservation after invalid save/cancel must remain; do not delete user config or send new traffic.

**Required validation:** Existing strict service cases, roundtrip, invalid save preserves bytes, temp cleanup/write failure, fixed UI destination, unsafe parent/junction and no plaintext credentials.


### CX-R020 — Path_Traversal — Result 12 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **333146485**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=LAaNk9cpfK7yLN280XnEsf1XZL8%3D); report line 1148 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Nonsecret service configuration |
| Source | `tests/s0/ServiceConnectionTests.cs:9` — Main, `args` |
| Sink | `apps/powerpoint-vsto/ServiceConnections.cs:95` — Save, `Exists` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `SERVICE-CONFIG` |

**Original description:** Method Main at line 9 of /powerpoint-intelligence-main/tests/s0/ServiceConnectionTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Save at line 95 of /powerpoint-intelligence-main/apps/powerpoint-vsto/ServiceConnections.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
[STAThread] public static int Main(string[] args){try{Run(args.Length==0?null:args[0]);return 0;}catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
→ try{File.WriteAllText(temp,xml,new UTF8Encoding(false,true));if(File.Exists(target))File.Replace(temp,target,null);else File.Move(temp,target);}finally{if(File.Exists(temp))File.Delete(temp);}
```

**Root cause / controls:** The path originates in CLI test arguments. ServiceConnectionTests makes a GUID-named child and constant connections.xml. Production UI uses the fixed per-user DefaultPath; Save uses a random sibling temp and atomic Replace/Move. The Delete sink cleans that owned temp. Strict schema/HTTPS/size validation already exists.

**Exploitability / business impact:** No untrusted production caller selects Save(path) in the reviewed call graph. Replace/delete/existence operations are ordinary owned persistence. A separate ownership/reparse issue or future API would need analysis. No live HTTP/auth bypass/credential storage exists here.

**Recommendation:** Keep nonsecret-only strict metadata, atomic saves and cancellation. Encapsulate fixed production store separately from test-only path injection or use an explicitly owned directory context. Handle ancestry under approved policy; do not invent connectivity/authentication success.

**Regression risk:** Atomic replacement, environment-independent fixtures and preservation after invalid save/cancel must remain; do not delete user config or send new traffic.

**Required validation:** Existing strict service cases, roundtrip, invalid save preserves bytes, temp cleanup/write failure, fixed UI destination, unsafe parent/junction and no plaintext credentials.


### CX-R021 — Path_Traversal — Result 13 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **955143290**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=1ZRyY2YrA5n5%2BGl53f6HgGX%2BCPc%3D); report line 1189 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | PIIDX index and writer lock |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `src/Library.Core/PiidxCatalog.cs:33` — Connect, `Exists` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `CATALOG-PATH` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Connect at line 33 of /powerpoint-intelligence-main/src/Library.Core/PiidxCatalog.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ if(!File.Exists(IndexPath))throw new FileNotFoundException("PIIDX index is missing.",IndexPath);
```

**Root cause / controls:** A test workspace/configured .piidx path reaches index/lock probes. The catalog canonicalizes its path, restricts .piidx, checks leaf reparse points and uses an exclusive bounded-wait writer lock. Selecting a different authorized index is intended.

**Exploitability / business impact:** A developer-selected fixture path does not establish remote arbitrary-file access. A hostile writable config/ancestor is a separate conditional threat (AR-02); no content-derived lock filename appears on the reported path.

**Recommendation:** Preserve configured index placement, single-writer/cancellation and read-only source roots. Define owned catalog authority and physical ancestry/handle checks where hostile locations are in scope; never silently migrate/create a rejected database.

**Regression risk:** Migration, journals, concurrency, cancellation, readonly libraries and redirected LocalAppData must remain supported.

**Required validation:** Schema/digest/tamper/corpus tests, lock timeout/concurrency, corrupt index preservation, valid custom paths, ancestor/alias/network policy and no partial catalog updates.


### CX-R022 — Path_Traversal — Result 14 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1947506698**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=HOfLTBVYrOlPEuxCtQhGij%2FwFYQ%3D); report line 1230 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | PIIDX index and writer lock |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `src/Library.Core/PiidxCatalog.cs:44` — WriterLock, `path` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `CATALOG-PATH` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in WriterLock at line 44 of /powerpoint-intelligence-main/src/Library.Core/PiidxCatalog.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ try{return new FileStream(path,FileMode.OpenOrCreate,FileAccess.ReadWrite,FileShare.None);}
```

**Root cause / controls:** A test workspace/configured .piidx path reaches index/lock probes. The catalog canonicalizes its path, restricts .piidx, checks leaf reparse points and uses an exclusive bounded-wait writer lock. Selecting a different authorized index is intended.

**Exploitability / business impact:** A developer-selected fixture path does not establish remote arbitrary-file access. A hostile writable config/ancestor is a separate conditional threat (AR-02); no content-derived lock filename appears on the reported path.

**Recommendation:** Preserve configured index placement, single-writer/cancellation and read-only source roots. Define owned catalog authority and physical ancestry/handle checks where hostile locations are in scope; never silently migrate/create a rejected database.

**Regression risk:** Migration, journals, concurrency, cancellation, readonly libraries and redirected LocalAppData must remain supported.

**Required validation:** Schema/digest/tamper/corpus tests, lock timeout/concurrency, corrupt index preservation, valid custom paths, ancestor/alias/network policy and no partial catalog updates.


### CX-R023 — Path_Traversal — Result 15 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1885568552**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=OA5wzNfJTLDrGJu4HqOkJ2NW3Xs%3D); report line 1271 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | PIIDX index and writer lock |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `src/Library.Core/PiidxCatalog.cs:34` — Connect, `GetAttributes` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `CATALOG-PATH` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Connect at line 34 of /powerpoint-intelligence-main/src/Library.Core/PiidxCatalog.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ if((File.GetAttributes(IndexPath)&FileAttributes.ReparsePoint)!=0)throw new InvalidDataException("PIIDX index cannot be a reparse point.");
```

**Root cause / controls:** A test workspace/configured .piidx path reaches index/lock probes. The catalog canonicalizes its path, restricts .piidx, checks leaf reparse points and uses an exclusive bounded-wait writer lock. Selecting a different authorized index is intended.

**Exploitability / business impact:** A developer-selected fixture path does not establish remote arbitrary-file access. A hostile writable config/ancestor is a separate conditional threat (AR-02); no content-derived lock filename appears on the reported path.

**Recommendation:** Preserve configured index placement, single-writer/cancellation and read-only source roots. Define owned catalog authority and physical ancestry/handle checks where hostile locations are in scope; never silently migrate/create a rejected database.

**Regression risk:** Migration, journals, concurrency, cancellation, readonly libraries and redirected LocalAppData must remain supported.

**Required validation:** Schema/digest/tamper/corpus tests, lock timeout/concurrency, corrupt index preservation, valid custom paths, ancestor/alias/network policy and no partial catalog updates.


### CX-R024 — Path_Traversal — Result 16 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **706674842**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=qKUebLmfH0bGeTh9uviMMK3sUV8%3D); report line 1312 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Nonsecret service configuration |
| Source | `tests/s0/ServiceConnectionTests.cs:9` — Main, `args` |
| Sink | `apps/powerpoint-vsto/ServiceConnections.cs:95` — Save, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `SERVICE-CONFIG` |

**Original description:** Method Main at line 9 of /powerpoint-intelligence-main/tests/s0/ServiceConnectionTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Save at line 95 of /powerpoint-intelligence-main/apps/powerpoint-vsto/ServiceConnections.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
[STAThread] public static int Main(string[] args){try{Run(args.Length==0?null:args[0]);return 0;}catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
→ try{File.WriteAllText(temp,xml,new UTF8Encoding(false,true));if(File.Exists(target))File.Replace(temp,target,null);else File.Move(temp,target);}finally{if(File.Exists(temp))File.Delete(temp);}
```

**Root cause / controls:** The path originates in CLI test arguments. ServiceConnectionTests makes a GUID-named child and constant connections.xml. Production UI uses the fixed per-user DefaultPath; Save uses a random sibling temp and atomic Replace/Move. The Delete sink cleans that owned temp. Strict schema/HTTPS/size validation already exists.

**Exploitability / business impact:** No untrusted production caller selects Save(path) in the reviewed call graph. Replace/delete/existence operations are ordinary owned persistence. A separate ownership/reparse issue or future API would need analysis. No live HTTP/auth bypass/credential storage exists here.

**Recommendation:** Keep nonsecret-only strict metadata, atomic saves and cancellation. Encapsulate fixed production store separately from test-only path injection or use an explicitly owned directory context. Handle ancestry under approved policy; do not invent connectivity/authentication success.

**Regression risk:** Atomic replacement, environment-independent fixtures and preservation after invalid save/cancel must remain; do not delete user config or send new traffic.

**Required validation:** Existing strict service cases, roundtrip, invalid save preserves bytes, temp cleanup/write failure, fixed UI destination, unsafe parent/junction and no plaintext credentials.


### CX-R025 — Path_Traversal — Result 17 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1651521647**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=1cTEJIMTKXCpgthmQYJZaKivL8A%3D); report line 1353 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Nonsecret service configuration |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `apps/powerpoint-vsto/ServiceConnections.cs:94` — Save, `CreateDirectory` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `SERVICE-CONFIG` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Save at line 94 of /powerpoint-intelligence-main/apps/powerpoint-vsto/ServiceConnections.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ string xml=Encode(settings);string target=Path.GetFullPath(path),directory=Path.GetDirectoryName(target);Directory.CreateDirectory(directory);string temp=Path.Combine(directory,".sc-"+Guid.NewGuid().ToString("N")+".tmp");
```

**Root cause / controls:** The path originates in CLI test arguments. ServiceConnectionTests makes a GUID-named child and constant connections.xml. Production UI uses the fixed per-user DefaultPath; Save uses a random sibling temp and atomic Replace/Move. The Delete sink cleans that owned temp. Strict schema/HTTPS/size validation already exists.

**Exploitability / business impact:** No untrusted production caller selects Save(path) in the reviewed call graph. Replace/delete/existence operations are ordinary owned persistence. A separate ownership/reparse issue or future API would need analysis. No live HTTP/auth bypass/credential storage exists here.

**Recommendation:** Keep nonsecret-only strict metadata, atomic saves and cancellation. Encapsulate fixed production store separately from test-only path injection or use an explicitly owned directory context. Handle ancestry under approved policy; do not invent connectivity/authentication success.

**Regression risk:** Atomic replacement, environment-independent fixtures and preservation after invalid save/cancel must remain; do not delete user config or send new traffic.

**Required validation:** Existing strict service cases, roundtrip, invalid save preserves bytes, temp cleanup/write failure, fixed UI destination, unsafe parent/junction and no plaintext credentials.


### CX-R026 — Path_Traversal — Result 18 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **1188711609**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=ADeHxJwvKDLoSPd2zpynn4VeAqA%3D); report line 1394 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Nonsecret service configuration |
| Source | `tests/s0/ServiceConnectionTests.cs:9` — Main, `args` |
| Sink | `apps/powerpoint-vsto/ServiceConnections.cs:94` — Save, `CreateDirectory` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `SERVICE-CONFIG` |

**Original description:** Method Main at line 9 of /powerpoint-intelligence-main/tests/s0/ServiceConnectionTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Save at line 94 of /powerpoint-intelligence-main/apps/powerpoint-vsto/ServiceConnections.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
[STAThread] public static int Main(string[] args){try{Run(args.Length==0?null:args[0]);return 0;}catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
→ string xml=Encode(settings);string target=Path.GetFullPath(path),directory=Path.GetDirectoryName(target);Directory.CreateDirectory(directory);string temp=Path.Combine(directory,".sc-"+Guid.NewGuid().ToString("N")+".tmp");
```

**Root cause / controls:** The path originates in CLI test arguments. ServiceConnectionTests makes a GUID-named child and constant connections.xml. Production UI uses the fixed per-user DefaultPath; Save uses a random sibling temp and atomic Replace/Move. The Delete sink cleans that owned temp. Strict schema/HTTPS/size validation already exists.

**Exploitability / business impact:** No untrusted production caller selects Save(path) in the reviewed call graph. Replace/delete/existence operations are ordinary owned persistence. A separate ownership/reparse issue or future API would need analysis. No live HTTP/auth bypass/credential storage exists here.

**Recommendation:** Keep nonsecret-only strict metadata, atomic saves and cancellation. Encapsulate fixed production store separately from test-only path injection or use an explicitly owned directory context. Handle ancestry under approved policy; do not invent connectivity/authentication success.

**Regression risk:** Atomic replacement, environment-independent fixtures and preservation after invalid save/cancel must remain; do not delete user config or send new traffic.

**Required validation:** Existing strict service cases, roundtrip, invalid save preserves bytes, temp cleanup/write failure, fixed UI destination, unsafe parent/junction and no plaintext credentials.


### CX-R027 — Path_Traversal — Result 19 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **507564216**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=YCKo6RMDlItmuAFFUkwKmUIBPkQ%3D); report line 1435 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Nonsecret service configuration |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `apps/powerpoint-vsto/ServiceConnections.cs:95` — Save, `Move` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `SERVICE-CONFIG` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Save at line 95 of /powerpoint-intelligence-main/apps/powerpoint-vsto/ServiceConnections.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ try{File.WriteAllText(temp,xml,new UTF8Encoding(false,true));if(File.Exists(target))File.Replace(temp,target,null);else File.Move(temp,target);}finally{if(File.Exists(temp))File.Delete(temp);}
```

**Root cause / controls:** The path originates in CLI test arguments. ServiceConnectionTests makes a GUID-named child and constant connections.xml. Production UI uses the fixed per-user DefaultPath; Save uses a random sibling temp and atomic Replace/Move. The Delete sink cleans that owned temp. Strict schema/HTTPS/size validation already exists.

**Exploitability / business impact:** No untrusted production caller selects Save(path) in the reviewed call graph. Replace/delete/existence operations are ordinary owned persistence. A separate ownership/reparse issue or future API would need analysis. No live HTTP/auth bypass/credential storage exists here.

**Recommendation:** Keep nonsecret-only strict metadata, atomic saves and cancellation. Encapsulate fixed production store separately from test-only path injection or use an explicitly owned directory context. Handle ancestry under approved policy; do not invent connectivity/authentication success.

**Regression risk:** Atomic replacement, environment-independent fixtures and preservation after invalid save/cancel must remain; do not delete user config or send new traffic.

**Required validation:** Existing strict service cases, roundtrip, invalid save preserves bytes, temp cleanup/write failure, fixed UI destination, unsafe parent/junction and no plaintext credentials.


### CX-R028 — Path_Traversal — Result 20 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-47860084**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=HGXbeMnvFmC%2FCQCusrf%2FtNviNbs%3D); report line 1476 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Nonsecret service configuration |
| Source | `tests/s0/ServiceConnectionTests.cs:9` — Main, `args` |
| Sink | `apps/powerpoint-vsto/ServiceConnections.cs:95` — Save, `Replace` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `SERVICE-CONFIG` |

**Original description:** Method Main at line 9 of /powerpoint-intelligence-main/tests/s0/ServiceConnectionTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Save at line 95 of /powerpoint-intelligence-main/apps/powerpoint-vsto/ServiceConnections.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
[STAThread] public static int Main(string[] args){try{Run(args.Length==0?null:args[0]);return 0;}catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
→ try{File.WriteAllText(temp,xml,new UTF8Encoding(false,true));if(File.Exists(target))File.Replace(temp,target,null);else File.Move(temp,target);}finally{if(File.Exists(temp))File.Delete(temp);}
```

**Root cause / controls:** The path originates in CLI test arguments. ServiceConnectionTests makes a GUID-named child and constant connections.xml. Production UI uses the fixed per-user DefaultPath; Save uses a random sibling temp and atomic Replace/Move. The Delete sink cleans that owned temp. Strict schema/HTTPS/size validation already exists.

**Exploitability / business impact:** No untrusted production caller selects Save(path) in the reviewed call graph. Replace/delete/existence operations are ordinary owned persistence. A separate ownership/reparse issue or future API would need analysis. No live HTTP/auth bypass/credential storage exists here.

**Recommendation:** Keep nonsecret-only strict metadata, atomic saves and cancellation. Encapsulate fixed production store separately from test-only path injection or use an explicitly owned directory context. Handle ancestry under approved policy; do not invent connectivity/authentication success.

**Regression risk:** Atomic replacement, environment-independent fixtures and preservation after invalid save/cancel must remain; do not delete user config or send new traffic.

**Required validation:** Existing strict service cases, roundtrip, invalid save preserves bytes, temp cleanup/write failure, fixed UI destination, unsafe parent/junction and no plaintext credentials.


### CX-R029 — Path_Traversal — Result 21 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1477818700**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=QbtrbW0spE1yBq55tWcf43HPSD8%3D); report line 1517 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Nonsecret service configuration |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `apps/powerpoint-vsto/ServiceConnections.cs:95` — Save, `Exists` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `SERVICE-CONFIG` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Save at line 95 of /powerpoint-intelligence-main/apps/powerpoint-vsto/ServiceConnections.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ try{File.WriteAllText(temp,xml,new UTF8Encoding(false,true));if(File.Exists(target))File.Replace(temp,target,null);else File.Move(temp,target);}finally{if(File.Exists(temp))File.Delete(temp);}
```

**Root cause / controls:** The path originates in CLI test arguments. ServiceConnectionTests makes a GUID-named child and constant connections.xml. Production UI uses the fixed per-user DefaultPath; Save uses a random sibling temp and atomic Replace/Move. The Delete sink cleans that owned temp. Strict schema/HTTPS/size validation already exists.

**Exploitability / business impact:** No untrusted production caller selects Save(path) in the reviewed call graph. Replace/delete/existence operations are ordinary owned persistence. A separate ownership/reparse issue or future API would need analysis. No live HTTP/auth bypass/credential storage exists here.

**Recommendation:** Keep nonsecret-only strict metadata, atomic saves and cancellation. Encapsulate fixed production store separately from test-only path injection or use an explicitly owned directory context. Handle ancestry under approved policy; do not invent connectivity/authentication success.

**Regression risk:** Atomic replacement, environment-independent fixtures and preservation after invalid save/cancel must remain; do not delete user config or send new traffic.

**Required validation:** Existing strict service cases, roundtrip, invalid save preserves bytes, temp cleanup/write failure, fixed UI destination, unsafe parent/junction and no plaintext credentials.


### CX-R030 — Path_Traversal — Result 22 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-56861040**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=VPYjr0UowoKtD3UfEJbdKGktSvQ%3D); report line 1558 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Nonsecret service configuration |
| Source | `tests/s0/ServiceConnectionTests.cs:9` — Main, `args` |
| Sink | `apps/powerpoint-vsto/ServiceConnections.cs:95` — Save, `Move` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `SERVICE-CONFIG` |

**Original description:** Method Main at line 9 of /powerpoint-intelligence-main/tests/s0/ServiceConnectionTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Save at line 95 of /powerpoint-intelligence-main/apps/powerpoint-vsto/ServiceConnections.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
[STAThread] public static int Main(string[] args){try{Run(args.Length==0?null:args[0]);return 0;}catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
→ try{File.WriteAllText(temp,xml,new UTF8Encoding(false,true));if(File.Exists(target))File.Replace(temp,target,null);else File.Move(temp,target);}finally{if(File.Exists(temp))File.Delete(temp);}
```

**Root cause / controls:** The path originates in CLI test arguments. ServiceConnectionTests makes a GUID-named child and constant connections.xml. Production UI uses the fixed per-user DefaultPath; Save uses a random sibling temp and atomic Replace/Move. The Delete sink cleans that owned temp. Strict schema/HTTPS/size validation already exists.

**Exploitability / business impact:** No untrusted production caller selects Save(path) in the reviewed call graph. Replace/delete/existence operations are ordinary owned persistence. A separate ownership/reparse issue or future API would need analysis. No live HTTP/auth bypass/credential storage exists here.

**Recommendation:** Keep nonsecret-only strict metadata, atomic saves and cancellation. Encapsulate fixed production store separately from test-only path injection or use an explicitly owned directory context. Handle ancestry under approved policy; do not invent connectivity/authentication success.

**Regression risk:** Atomic replacement, environment-independent fixtures and preservation after invalid save/cancel must remain; do not delete user config or send new traffic.

**Required validation:** Existing strict service cases, roundtrip, invalid save preserves bytes, temp cleanup/write failure, fixed UI destination, unsafe parent/junction and no plaintext credentials.


### CX-R031 — Path_Traversal — Result 23 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1477818700**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=w4q6ja6SJYsr1jKzW2moh43nj4g%3D); report line 1599 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Nonsecret service configuration |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `apps/powerpoint-vsto/ServiceConnections.cs:95` — Save, `Replace` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `SERVICE-CONFIG` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Save at line 95 of /powerpoint-intelligence-main/apps/powerpoint-vsto/ServiceConnections.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ try{File.WriteAllText(temp,xml,new UTF8Encoding(false,true));if(File.Exists(target))File.Replace(temp,target,null);else File.Move(temp,target);}finally{if(File.Exists(temp))File.Delete(temp);}
```

**Root cause / controls:** The path originates in CLI test arguments. ServiceConnectionTests makes a GUID-named child and constant connections.xml. Production UI uses the fixed per-user DefaultPath; Save uses a random sibling temp and atomic Replace/Move. The Delete sink cleans that owned temp. Strict schema/HTTPS/size validation already exists.

**Exploitability / business impact:** No untrusted production caller selects Save(path) in the reviewed call graph. Replace/delete/existence operations are ordinary owned persistence. A separate ownership/reparse issue or future API would need analysis. No live HTTP/auth bypass/credential storage exists here.

**Recommendation:** Keep nonsecret-only strict metadata, atomic saves and cancellation. Encapsulate fixed production store separately from test-only path injection or use an explicitly owned directory context. Handle ancestry under approved policy; do not invent connectivity/authentication success.

**Regression risk:** Atomic replacement, environment-independent fixtures and preservation after invalid save/cancel must remain; do not delete user config or send new traffic.

**Required validation:** Existing strict service cases, roundtrip, invalid save preserves bytes, temp cleanup/write failure, fixed UI destination, unsafe parent/junction and no plaintext credentials.


### CX-R032 — Path_Traversal — Result 24 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-47860084**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=mze2hs7RNMWIUThkc1ecoDXuPIw%3D); report line 1640 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Nonsecret service configuration |
| Source | `tests/s0/ServiceConnectionTests.cs:9` — Main, `args` |
| Sink | `apps/powerpoint-vsto/ServiceConnections.cs:95` — Save, `Exists` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `SERVICE-CONFIG` |

**Original description:** Method Main at line 9 of /powerpoint-intelligence-main/tests/s0/ServiceConnectionTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Save at line 95 of /powerpoint-intelligence-main/apps/powerpoint-vsto/ServiceConnections.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
[STAThread] public static int Main(string[] args){try{Run(args.Length==0?null:args[0]);return 0;}catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
→ try{File.WriteAllText(temp,xml,new UTF8Encoding(false,true));if(File.Exists(target))File.Replace(temp,target,null);else File.Move(temp,target);}finally{if(File.Exists(temp))File.Delete(temp);}
```

**Root cause / controls:** The path originates in CLI test arguments. ServiceConnectionTests makes a GUID-named child and constant connections.xml. Production UI uses the fixed per-user DefaultPath; Save uses a random sibling temp and atomic Replace/Move. The Delete sink cleans that owned temp. Strict schema/HTTPS/size validation already exists.

**Exploitability / business impact:** No untrusted production caller selects Save(path) in the reviewed call graph. Replace/delete/existence operations are ordinary owned persistence. A separate ownership/reparse issue or future API would need analysis. No live HTTP/auth bypass/credential storage exists here.

**Recommendation:** Keep nonsecret-only strict metadata, atomic saves and cancellation. Encapsulate fixed production store separately from test-only path injection or use an explicitly owned directory context. Handle ancestry under approved policy; do not invent connectivity/authentication success.

**Regression risk:** Atomic replacement, environment-independent fixtures and preservation after invalid save/cancel must remain; do not delete user config or send new traffic.

**Required validation:** Existing strict service cases, roundtrip, invalid save preserves bytes, temp cleanup/write failure, fixed UI destination, unsafe parent/junction and no plaintext credentials.


### CX-R033 — Path_Traversal — Result 25 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **1027859451**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=prf%2FG20fkbrV1cVnAg2XKyRmomI%3D); report line 1681 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/DemoScenario.cs:100` — Prepare, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Prepare at line 100 of /powerpoint-intelligence-main/tests/s0/DemoScenario.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ File.WriteAllText(linkedSource,id=="waterfall"?"Category,Change\nA,137.5\nB,-8\nC,20":"Category,Actual,Plan\n1,137.5,15\n2,22,18\n3,30,35");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R034 — Path_Traversal — Result 26 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-419043739**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=UCUGxd7c%2FuWuMYeFlEEkmB85BAo%3D); report line 1722 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/DemoScenario.cs:180` — ExecuteAndAssert, `OpenRead` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in ExecuteAndAssert at line 180 of /powerpoint-intelligence-main/tests/s0/DemoScenario.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ case "export-json":using(var stream=File.OpenRead(output)){var model=(ChartDocument)new DataContractJsonSerializer(typeof(ChartDocument)).ReadObject(stream);MetadataCodec.Validate(model);Check(model.ChartId==managedId,"Exported chart identity");}break;
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R035 — Path_Traversal — Result 27 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **601440768**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=ZlHvfegC62a7%2FttRbYAlyE8dtI0%3D); report line 1763 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/DemoScenario.cs:56` — Assets, `Exists` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Assets at line 56 of /powerpoint-intelligence-main/tests/s0/DemoScenario.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ if(!File.Exists(path))using(var image=new Bitmap(30,20)){using(var graphics=Graphics.FromImage(image))graphics.Clear(Color.Navy);image.Save(path,System.Drawing.Imaging.ImageFormat.Png);}
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R036 — Path_Traversal — Result 28 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1372051600**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=b5YYBMN6vBAATo8xvqFDnGwD0Zc%3D); report line 1804 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Nonsecret service configuration |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `apps/powerpoint-vsto/ServiceConnections.cs:90` — Load, `Exists` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `SERVICE-CONFIG` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Load at line 90 of /powerpoint-intelligence-main/apps/powerpoint-vsto/ServiceConnections.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ if(!File.Exists(path))return ServiceConnectionSettings.Empty();if(new FileInfo(path).Length>MaxFileBytes)throw new InvalidDataException("Saved configuration exceeds byte limit.");return Decode(File.ReadAllText(path,new UTF8Encoding(false,true)));
```

**Root cause / controls:** The path originates in CLI test arguments. ServiceConnectionTests makes a GUID-named child and constant connections.xml. Production UI uses the fixed per-user DefaultPath; Save uses a random sibling temp and atomic Replace/Move. The Delete sink cleans that owned temp. Strict schema/HTTPS/size validation already exists.

**Exploitability / business impact:** No untrusted production caller selects Save(path) in the reviewed call graph. Replace/delete/existence operations are ordinary owned persistence. A separate ownership/reparse issue or future API would need analysis. No live HTTP/auth bypass/credential storage exists here.

**Recommendation:** Keep nonsecret-only strict metadata, atomic saves and cancellation. Encapsulate fixed production store separately from test-only path injection or use an explicitly owned directory context. Handle ancestry under approved policy; do not invent connectivity/authentication success.

**Regression risk:** Atomic replacement, environment-independent fixtures and preservation after invalid save/cancel must remain; do not delete user config or send new traffic.

**Required validation:** Existing strict service cases, roundtrip, invalid save preserves bytes, temp cleanup/write failure, fixed UI destination, unsafe parent/junction and no plaintext credentials.


### CX-R037 — Path_Traversal — Result 29 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-315100207**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=KDJQcnDCZ5BPb7l2uW%2FO5IZcI2M%3D); report line 1845 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Nonsecret service configuration |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `apps/powerpoint-vsto/ServiceConnections.cs:90` — Load, `ReadAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `SERVICE-CONFIG` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Load at line 90 of /powerpoint-intelligence-main/apps/powerpoint-vsto/ServiceConnections.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ if(!File.Exists(path))return ServiceConnectionSettings.Empty();if(new FileInfo(path).Length>MaxFileBytes)throw new InvalidDataException("Saved configuration exceeds byte limit.");return Decode(File.ReadAllText(path,new UTF8Encoding(false,true)));
```

**Root cause / controls:** The path originates in CLI test arguments. ServiceConnectionTests makes a GUID-named child and constant connections.xml. Production UI uses the fixed per-user DefaultPath; Save uses a random sibling temp and atomic Replace/Move. The Delete sink cleans that owned temp. Strict schema/HTTPS/size validation already exists.

**Exploitability / business impact:** No untrusted production caller selects Save(path) in the reviewed call graph. Replace/delete/existence operations are ordinary owned persistence. A separate ownership/reparse issue or future API would need analysis. No live HTTP/auth bypass/credential storage exists here.

**Recommendation:** Keep nonsecret-only strict metadata, atomic saves and cancellation. Encapsulate fixed production store separately from test-only path injection or use an explicitly owned directory context. Handle ancestry under approved policy; do not invent connectivity/authentication success.

**Regression risk:** Atomic replacement, environment-independent fixtures and preservation after invalid save/cancel must remain; do not delete user config or send new traffic.

**Required validation:** Existing strict service cases, roundtrip, invalid save preserves bytes, temp cleanup/write failure, fixed UI destination, unsafe parent/junction and no plaintext credentials.


### CX-R038 — Path_Traversal — Result 30 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1654445831**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=m3cXpZsIBSYcabBVEynwtU3Hn7U%3D); report line 1886 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Nonsecret service configuration |
| Source | `tests/s0/ServiceConnectionTests.cs:9` — Main, `args` |
| Sink | `apps/powerpoint-vsto/ServiceConnections.cs:90` — Load, `ReadAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `SERVICE-CONFIG` |

**Original description:** Method Main at line 9 of /powerpoint-intelligence-main/tests/s0/ServiceConnectionTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Load at line 90 of /powerpoint-intelligence-main/apps/powerpoint-vsto/ServiceConnections.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
[STAThread] public static int Main(string[] args){try{Run(args.Length==0?null:args[0]);return 0;}catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
→ if(!File.Exists(path))return ServiceConnectionSettings.Empty();if(new FileInfo(path).Length>MaxFileBytes)throw new InvalidDataException("Saved configuration exceeds byte limit.");return Decode(File.ReadAllText(path,new UTF8Encoding(false,true)));
```

**Root cause / controls:** The path originates in CLI test arguments. ServiceConnectionTests makes a GUID-named child and constant connections.xml. Production UI uses the fixed per-user DefaultPath; Save uses a random sibling temp and atomic Replace/Move. The Delete sink cleans that owned temp. Strict schema/HTTPS/size validation already exists.

**Exploitability / business impact:** No untrusted production caller selects Save(path) in the reviewed call graph. Replace/delete/existence operations are ordinary owned persistence. A separate ownership/reparse issue or future API would need analysis. No live HTTP/auth bypass/credential storage exists here.

**Recommendation:** Keep nonsecret-only strict metadata, atomic saves and cancellation. Encapsulate fixed production store separately from test-only path injection or use an explicitly owned directory context. Handle ancestry under approved policy; do not invent connectivity/authentication success.

**Regression risk:** Atomic replacement, environment-independent fixtures and preservation after invalid save/cancel must remain; do not delete user config or send new traffic.

**Required validation:** Existing strict service cases, roundtrip, invalid save preserves bytes, temp cleanup/write failure, fixed UI destination, unsafe parent/junction and no plaintext credentials.


### CX-R039 — Path_Traversal — Result 31 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **1760295004**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=OPKG4AmecTGOraHvljYF4pyiQeE%3D); report line 1927 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/library/TemplateLibraryTests.cs:21` — Main, `args` |
| Sink | `tests/library/TemplateFidelity.cs:107` — PersistedGeometry, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 21 of /powerpoint-intelligence-main/tests/library/TemplateLibraryTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in PersistedGeometry at line 107 of /powerpoint-intelligence-main/tests/library/TemplateFidelity.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
[STAThread] public static int Main(string[] args)
→ try{var deck=(PowerPoint.Presentation)slide.Parent;deck.SaveCopyAs(file,PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,Office.MsoTriState.msoFalse);string geometry=PackageGeometry(file,slide.SlideIndex);if(diagnosticFile!=null)File.WriteAllText(diagnosticFile,geometry);return geometry;}
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R040 — Path_Traversal — Result 32 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-90541740**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=%2F6dyTvajWsejAugWJhJLVj0IIJo%3D); report line 1968 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `tests/s0/DataSourceTests.cs:102` — Run, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 102 of /powerpoint-intelligence-main/tests/s0/DataSourceTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ string numericJson=Path.Combine(folder,"numeric.json");File.WriteAllText(numericJson,"{\"headers\":[\"Category\",\"Value\"],\"rows\":[[\"East\",137.5]]}");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R041 — Path_Traversal — Result 33 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **1947077620**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=9tshnRuRaOh%2FQ2p7zF4vVco0KNM%3D); report line 2009 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `tests/s0/DataSourceTests.cs:100` — Run, `Delete` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 100 of /powerpoint-intelligence-main/tests/s0/DataSourceTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ File.Delete(path);Reject(()=>LinkedChartData.Load(link),format+" moved file fails instead of using sample data");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R042 — Path_Traversal — Result 34 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **46418524**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=eeBNEtYVAX501e1%2BF5fF%2FMK39m8%3D); report line 2050 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/DataSourceTests.cs:127` — OfficeRoundtrip, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in OfficeRoundtrip at line 127 of /powerpoint-intelligence-main/tests/s0/DataSourceTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ string source=Path.GetFullPath(Path.Combine(directory,"reference"+option.Extensions[0]));File.WriteAllText(source,"reference only");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R043 — Path_Traversal — Result 35 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1299226345**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=1q6p%2FVmmWFx3VlQpJNII8sHblIw%3D); report line 2091 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/DemoScenario.cs:55` — Assets, `CreateDirectory` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Assets at line 55 of /powerpoint-intelligence-main/tests/s0/DemoScenario.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ var directory=Path.Combine(root,"assets");Directory.CreateDirectory(directory);var path=Path.Combine(directory,"sample.png");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R044 — Path_Traversal — Result 36 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1761994869**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=ez9sm8xwlWdE0Sni9PLLoBo9WFw%3D); report line 2132 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `tests/s0/DataSourceTests.cs:58` — Run, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 58 of /powerpoint-intelligence-main/tests/s0/DataSourceTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ string path=Path.Combine(folder,"日本語 & source"+extension);File.WriteAllText(path,"owned reference-only fixture");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R045 — Path_Traversal — Result 37 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1720959708**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=OzG%2BGsRnvJ%2BJNdYBLSNzkLflaDI%3D); report line 2173 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `tests/s0/DataSourceTests.cs:66` — Run, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 66 of /powerpoint-intelligence-main/tests/s0/DataSourceTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ string wrong=Path.Combine(folder,"wrong.bin");File.WriteAllText(wrong,"fixture");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R046 — Path_Traversal — Result 38 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **462301936**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=QHgALfXjuFdymcCH0LkzRQOECWg%3D); report line 2214 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `tests/s0/DataSourceTests.cs:98` — Run, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 98 of /powerpoint-intelligence-main/tests/s0/DataSourceTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ File.WriteAllText(path,content);var link=DataSources.Create(format,path);var table=LinkedChartData.Load(link);
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R047 — Path_Traversal — Result 39 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **1407586478**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=0faOjG8MMMkpoEXFC7j4JoxhaDw%3D); report line 2255 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/library/TemplateLibraryTests.cs:21` — Main, `args` |
| Sink | `tests/library/TemplateLibraryTests.cs:134` — Geometry, `WriteAllLines` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 21 of /powerpoint-intelligence-main/tests/library/TemplateLibraryTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Geometry at line 134 of /powerpoint-intelligence-main/tests/library/TemplateLibraryTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
[STAThread] public static int Main(string[] args)
→ var rows=new List<string>();Action<PowerPoint.Shape,string> visit=null;visit=(shape,path)=>{string text="";if(shape.HasTextFrame==Office.MsoTriState.msoTrue&&shape.TextFrame.HasText==Office.MsoTriState.msoTrue)using(var hash=SHA256.Create())text=BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(shape.TextFrame.TextRange.Text)));rows.Add(path+"\t"+shape.Type+"\t"+String.Join("\t",new[]{shape.Left,shape.Top,shape.Width,shape.Height,shape.Rotation}.Select(x=>x.ToString("R",CultureInfo.InvariantCulture)))+"\t"+text);if(shape.Type==Office.MsoShapeType.msoGroup)for(int i=1;i<=shape.GroupItems.Count;i++)visit(shape.GroupItems[i],path+"."+i);};for(int i=1;i<=slide.Shapes.Count;i++)visit(slide.Shapes[i],i.ToString(CultureInfo.InvariantCulture));File.WriteAllLines(output,rows);return String.Join("\n",rows);
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R048 — Path_Traversal — Result 40 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **758400680**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=9IyWAKdgwRXVRQ%2Bp5zET8QlG2XA%3D); report line 2296 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/ServiceConnectionTests.cs:9` — Main, `args` |
| Sink | `tests/s0/ServiceConnectionTests.cs:33` — Run, `WriteAllBytes` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 9 of /powerpoint-intelligence-main/tests/s0/ServiceConnectionTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 33 of /powerpoint-intelligence-main/tests/s0/ServiceConnectionTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
[STAThread] public static int Main(string[] args){try{Run(args.Length==0?null:args[0]);return 0;}catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
→ string tooLarge=Path.Combine(root,"oversize.xml");File.WriteAllBytes(tooLarge,new byte[ServiceConnections.MaxFileBytes+1]);Reject(()=>ServiceConnections.Load(tooLarge),"saved byte bound");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R049 — Path_Traversal — Result 41 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1517092728**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=RvaNuTvm%2BNj6rJJ%2FDK5FV8TiQj0%3D); report line 2337 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/library/TemplateLibraryTests.cs:21` — Main, `args` |
| Sink | `tests/library/TemplateFidelity.cs:129` — Formatting, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 21 of /powerpoint-intelligence-main/tests/library/TemplateLibraryTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Formatting at line 129 of /powerpoint-intelligence-main/tests/library/TemplateFidelity.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
[STAThread] public static int Main(string[] args)
→ if(diagnosticFile!=null)File.WriteAllText(diagnosticFile,text.ToString());
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R050 — Path_Traversal — Result 42 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **2097746304**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=T2ChVPRSY%2BLM4HcQYOSaqns7EFA%3D); report line 2378 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `tests/s0/ServiceConnectionTests.cs:33` — Run, `WriteAllBytes` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 33 of /powerpoint-intelligence-main/tests/s0/ServiceConnectionTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ string tooLarge=Path.Combine(root,"oversize.xml");File.WriteAllBytes(tooLarge,new byte[ServiceConnections.MaxFileBytes+1]);Reject(()=>ServiceConnections.Load(tooLarge),"saved byte bound");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R051 — Path_Traversal — Result 43 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **1237596672**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=5HKsds0NJHUQo0B1ZfwrGVrohe0%3D); report line 2419 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/ServiceConnectionTests.cs:9` — Main, `args` |
| Sink | `tests/s0/ServiceConnectionTests.cs:32` — Run, `ReadAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 9 of /powerpoint-intelligence-main/tests/s0/ServiceConnectionTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 32 of /powerpoint-intelligence-main/tests/s0/ServiceConnectionTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
[STAThread] public static int Main(string[] args){try{Run(args.Length==0?null:args[0]);return 0;}catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
→ var badSave=Valid();badSave.Profiles[0].Endpoint="https://example.invalid/?token=secret";string preserved=File.ReadAllText(file);Reject(()=>ServiceConnections.Save(file,badSave),"invalid save rejected");Check(File.ReadAllText(file)==preserved,"invalid save preserves previous config");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R052 — Path_Traversal — Result 44 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-152178112**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=eOpUzxNR8%2BTRVZ%2FYrOvmjMXFECQ%3D); report line 2460 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `tests/s0/ServiceConnectionTests.cs:31` — Run, `ReadAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 31 of /powerpoint-intelligence-main/tests/s0/ServiceConnectionTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ Check(ServiceConnections.Load(file).Profiles.All(p=>ServiceConnections.Status(p)=="Not configured")&&!Directory.Exists(root),"read/cancel path creates no file or directory");ServiceConnections.Save(file,Valid());string first=File.ReadAllText(file);Check(ServiceConnections.Encode(ServiceConnections.Load(file))==first,"file round trip");var changed=Valid();changed.Profiles[0].Deployment="second-approved-deployment";ServiceConnections.Save(file,changed);Check(ServiceConnections.Load(file).Profiles[0].Deployment=="second-approved-deployment","existing config atomically replaced");Check(Directory.GetFiles(root).Length==1,"temporary files cleaned after save");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R053 — Path_Traversal — Result 45 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **934697572**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=rCnZQhRqYmVmumlgiQW2ERETOeA%3D); report line 2501 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `tests/s0/ServiceConnectionTests.cs:32` — Run, `ReadAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 32 of /powerpoint-intelligence-main/tests/s0/ServiceConnectionTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ var badSave=Valid();badSave.Profiles[0].Endpoint="https://example.invalid/?token=secret";string preserved=File.ReadAllText(file);Reject(()=>ServiceConnections.Save(file,badSave),"invalid save rejected");Check(File.ReadAllText(file)==preserved,"invalid save preserves previous config");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R054 — Path_Traversal — Result 46 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1719932254**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=%2B1sMmiqsfzL5NOYYe3dZKY5WkJE%3D); report line 2542 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/ContextPropertiesOfficeTests.cs:46` — EnsureTestForeground, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in EnsureTestForeground at line 46 of /powerpoint-intelligence-main/tests/s0/ContextPropertiesOfficeTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ var doc=new XmlDocument();var root=doc.CreateElement("testForeground");doc.AppendChild(root);root.SetAttribute("checkpoint",checkpoint);root.SetAttribute("initiallyForeground",result.InitiallyForeground?"true":"false");root.SetAttribute("requestAccepted",result.RequestAccepted?"true":"false");root.SetAttribute("verified",result.Verified?"true":"false");root.SetAttribute("promptShown",promptShown?"true":"false");root.SetAttribute("userStarted",userStarted?"true":"false");root.SetAttribute("userRequestAccepted",userRequestAccepted?"true":"false");File.WriteAllText(Path.Combine(evidenceDirectory,"context-foreground-"+checkpoint+".xml"),doc.OuterXml);
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R055 — Path_Traversal — Result 47 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1345387132**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=8fwkln%2B35vKKqche26enAJ2akEc%3D); report line 2583 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/ContextPropertiesOfficeTests.cs:24` — Wait, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Wait at line 24 of /powerpoint-intelligence-main/tests/s0/ContextPropertiesOfficeTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ var watch=System.Diagnostics.Stopwatch.StartNew();XmlElement state=null;do{Application.DoEvents();state=State(api);if(visible&&ContextSelectionBridge.GetForegroundWindow()!=expectedOwner){File.WriteAllText(Path.Combine(evidenceDirectory,"context-properties-focus-lost.xml"),state.OuterXml);throw new ContextTestForegroundUnavailableException("Foreground focus was lost during the "+kind+" panel assertion. This assertion is BLOCKED, not a panel failure or a pass. See context-properties-focus-lost.xml.");}if(state.GetAttribute("refreshes")!="0"&&!String.IsNullOrEmpty(state.GetAttribute("refreshes"))&&state.GetAttribute("visible")== (visible?"true":"false")&&(!visible||state.GetAttribute("kind")==kind&&state.GetAttribute("selectionMatches")=="true"))return state;Thread.Sleep(75);}while(watch.ElapsedMilliseconds<8000);
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R056 — Path_Traversal — Result 48 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-86767479**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=eMioYE5YNtdWZ2msuv0r57E2IL0%3D); report line 2624 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/ContextPropertiesOfficeTests.cs:25` — Wait, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Wait at line 25 of /powerpoint-intelligence-main/tests/s0/ContextPropertiesOfficeTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ File.WriteAllText(Path.Combine(evidenceDirectory,"context-properties-failure.xml"),state.OuterXml);
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R057 — Path_Traversal — Result 49 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-695368680**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=TQjneoiBCesQ%2FxnkaoETHc7z%2B3I%3D); report line 2665 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/ServiceConnectionTests.cs:9` — Main, `args` |
| Sink | `tests/s0/ServiceConnectionTests.cs:31` — Run, `ReadAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 9 of /powerpoint-intelligence-main/tests/s0/ServiceConnectionTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 31 of /powerpoint-intelligence-main/tests/s0/ServiceConnectionTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
[STAThread] public static int Main(string[] args){try{Run(args.Length==0?null:args[0]);return 0;}catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
→ Check(ServiceConnections.Load(file).Profiles.All(p=>ServiceConnections.Status(p)=="Not configured")&&!Directory.Exists(root),"read/cancel path creates no file or directory");ServiceConnections.Save(file,Valid());string first=File.ReadAllText(file);Check(ServiceConnections.Encode(ServiceConnections.Load(file))==first,"file round trip");var changed=Valid();changed.Profiles[0].Deployment="second-approved-deployment";ServiceConnections.Save(file,changed);Check(ServiceConnections.Load(file).Profiles[0].Deployment=="second-approved-deployment","existing config atomically replaced");Check(Directory.GetFiles(root).Length==1,"temporary files cleaned after save");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R058 — Path_Traversal — Result 50 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1532911912**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=ULFg%2FaeHJRDkQKH0w5aDwnB4DO8%3D); report line 2706 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/ContextPropertiesOfficeTests.cs:55` — ProbeSelectedText, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in ProbeSelectedText at line 55 of /powerpoint-intelligence-main/tests/s0/ContextPropertiesOfficeTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ finally{File.WriteAllText(Path.Combine(evidenceDirectory,"context-selection-probe-"+checkpoint+".xml"),doc.OuterXml);}
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R059 — Path_Traversal — Result 51 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1429002613**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=037lUjvK6mixxACxKJRNe9NmIkY%3D); report line 2747 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/library/TemplateLibraryTests.cs:21` — Main, `args` |
| Sink | `tests/library/TemplateLibraryTests.cs:157` — OriginalRoundtrip, `WriteAllLines` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 21 of /powerpoint-intelligence-main/tests/library/TemplateLibraryTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in OriginalRoundtrip at line 157 of /powerpoint-intelligence-main/tests/library/TemplateLibraryTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
[STAThread] public static int Main(string[] args)
→ for(int i=0;i<assets.Length;i++){var slide=deck.Slides[i+2];bool exact=RenderHash(slide,Path.Combine(output,"original-"+i+"-reopened.png"),heights[i])==pixelHashes[i];var raster=TemplateFidelity.Files(Path.Combine(output,"original-"+i+"-source.png"),Path.Combine(output,"original-"+i+"-reopened.png"));bool geometry=Geometry(slide,Path.Combine(output,"original-"+i+"-reopened-geometry.tsv"))==nativeGeometry[i],format=TemplateFidelity.Formatting(slide,Path.Combine(output,"original-"+i+"-reopened-format.txt"))==nativeFormatting[i];fidelityRows.Add(i+",reopen,"+exact+","+geometry+","+format+",,,"+raster.Accepted+","+raster.ChangedPixels+","+raster.MaxChannelDelta+","+raster.AlphaChanges+","+raster.NonEdgeChanges+","+raster.AllowedPixels);File.WriteAllLines(Path.Combine(output,"original-fidelity.csv"),fidelityRows);if(!raster.Accepted||!geometry||!format)fidelityFailures.Add("Reopen fidelity "+i+": tolerance="+raster.Accepted+", geometry="+geometry+", formatting="+format);Check(slide.TimeLine.MainSequence.Count==sourceEffects[i]&&slide.Tags["PPI_LIBRARY_ASSET_VERSION"]==assets[i].Signature,"Original builds/provenance lost on reopen");}
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R060 — Path_Traversal — Result 52 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **2129211647**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=AdNMFSvl6r5H0QYeOh1URUqNVBE%3D); report line 2788 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/LinkedChartDiagnostic.cs:38` — Run, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 38 of /powerpoint-intelligence-main/tests/s0/LinkedChartDiagnostic.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ string source=Path.Combine(Path.GetFullPath(output),"diagnostic-fixture.csv");File.WriteAllText(source,"Category,Actual,Plan\r\nA,137.5,150\r\nB,20,25\r\nC,30,35\r\n");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R061 — Path_Traversal — Result 53 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1429002613**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=fgAH3TuhlpTVKSEyvxqfvfr4Aps%3D); report line 2829 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/library/TemplateLibraryTests.cs:21` — Main, `args` |
| Sink | `tests/library/TemplateLibraryTests.cs:152` — OriginalRoundtrip, `WriteAllLines` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 21 of /powerpoint-intelligence-main/tests/library/TemplateLibraryTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in OriginalRoundtrip at line 152 of /powerpoint-intelligence-main/tests/library/TemplateLibraryTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
[STAThread] public static int Main(string[] args)
→ bool exactPixels=RenderHash(slide,Path.Combine(output,"original-"+i+"-inserted.png"),heights[i])==pixelHashes[i];bool exactGeometry=Geometry(slide,Path.Combine(output,"original-"+i+"-inserted-geometry.tsv"))==geometry;bool exactFormat=TemplateFidelity.Formatting(slide,Path.Combine(output,"original-"+i+"-inserted-format.txt"))==nativeFormatting[i];bool repeatSource=RenderHash(session.Source(item),Path.Combine(output,"original-"+i+"-source-repeat.png"),heights[i])==pixelHashes[i];var raster=TemplateFidelity.Files(Path.Combine(output,"original-"+i+"-source.png"),Path.Combine(output,"original-"+i+"-inserted.png"));var repeatRaster=TemplateFidelity.Files(Path.Combine(output,"original-"+i+"-source.png"),Path.Combine(output,"original-"+i+"-source-repeat.png"));fidelityRows.Add(i+",insert,"+exactPixels+","+exactGeometry+","+exactFormat+","+repeatSource+","+repeatRaster.Accepted+","+raster.Accepted+","+raster.ChangedPixels+","+raster.MaxChannelDelta+","+raster.AlphaChanges+","+raster.NonEdgeChanges+","+raster.AllowedPixels);File.WriteAllLines(Path.Combine(output,"original-fidelity.csv"),fidelityRows);if(!raster.Accepted||!repeatRaster.Accepted||!exactGeometry||!exactFormat)fidelityFailures.Add("Insertion fidelity "+i+": exactPixels="+exactPixels+", tolerance="+raster.Accepted+", geometry="+exactGeometry+", formatting="+exactFormat+", repeated source="+repeatSource);Check(slide.Tags["PPI_LIBRARY_FORMATTING"]=="KEEP_SOURCE","Source provenance missing");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R062 — Path_Traversal — Result 54 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **1703303725**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=M5Tu3xfYZ9wb1tL4qqlq9bgLZtY%3D); report line 2870 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/LinkedChartDiagnostic.cs:45` — Run, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 45 of /powerpoint-intelligence-main/tests/s0/LinkedChartDiagnostic.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ File.WriteAllText(source,"Category,Actual,Plan\r\nA,42.5,150\r\nB,20,25\r\nC,30,35\r\n");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R063 — Path_Traversal — Result 55 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **1552682661**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=4RxSjuv9iBFXPjHMWrJ2J9cnlO8%3D); report line 2911 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `tests/s0/DataSourceTests.cs:18` — Run, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 18 of /powerpoint-intelligence-main/tests/s0/DataSourceTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ Directory.CreateDirectory(output);var registryPath=Path.Combine(Path.GetFullPath(output),"registry-fixture.csv");File.WriteAllText(registryPath,"Category,Value\r\nFixture,1\r\n");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R064 — Path_Traversal — Result 56 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1579543158**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=6HPoj9WWxywzdkc7fXCXqpwJ180%3D); report line 2952 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/DataSourceTests.cs:121` — OfficeRoundtrip, `CreateDirectory` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in OfficeRoundtrip at line 121 of /powerpoint-intelligence-main/tests/s0/DataSourceTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ string directory=Path.Combine(output,"data-source-reference-fixtures");Directory.CreateDirectory(directory);
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R065 — Path_Traversal — Result 57 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-473184483**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=d%2FVgcGQpX%2BWZw4jgGfJcZI%2FIdWg%3D); report line 2993 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `tests/s0/DataSourceTests.cs:49` — Run, `CreateDirectory` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 49 of /powerpoint-intelligence-main/tests/s0/DataSourceTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ string folder=Path.Combine(Path.GetFullPath(output),"data-source-fixtures");Directory.CreateDirectory(folder);
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R066 — Path_Traversal — Result 58 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-82668503**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=T5ABTrrHNyDPfR%2FM%2FeXesGCmyoM%3D); report line 3034 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/LinkedChartDiagnostic.cs:16` — Lambda, `AppendAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Lambda at line 16 of /powerpoint-intelligence-main/tests/s0/LinkedChartDiagnostic.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ Action<string,string,string> record=(stage,key,value)=>File.AppendAllText(log,stage+"\t"+key+"\t"+value.Replace('\r',' ').Replace('\n',' ')+"\r\n");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R067 — Path_Traversal — Result 59 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **1966682696**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=8T3l2deh8eqcwfxDme8Sa6b7Etc%3D); report line 3075 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/ServiceConnectionTests.cs:9` — Main, `args` |
| Sink | `tests/s0/ServiceConnectionTests.cs:31` — Run, `GetFiles` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 9 of /powerpoint-intelligence-main/tests/s0/ServiceConnectionTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 31 of /powerpoint-intelligence-main/tests/s0/ServiceConnectionTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
[STAThread] public static int Main(string[] args){try{Run(args.Length==0?null:args[0]);return 0;}catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
→ Check(ServiceConnections.Load(file).Profiles.All(p=>ServiceConnections.Status(p)=="Not configured")&&!Directory.Exists(root),"read/cancel path creates no file or directory");ServiceConnections.Save(file,Valid());string first=File.ReadAllText(file);Check(ServiceConnections.Encode(ServiceConnections.Load(file))==first,"file round trip");var changed=Valid();changed.Profiles[0].Deployment="second-approved-deployment";ServiceConnections.Save(file,changed);Check(ServiceConnections.Load(file).Profiles[0].Deployment=="second-approved-deployment","existing config atomically replaced");Check(Directory.GetFiles(root).Length==1,"temporary files cleaned after save");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R068 — Path_Traversal — Result 60 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-485459168**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=BYhOIdbH%2F8vHuhcAivhI21DyQnM%3D); report line 3116 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `tests/s0/ServiceConnectionTests.cs:31` — Run, `GetFiles` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 31 of /powerpoint-intelligence-main/tests/s0/ServiceConnectionTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ Check(ServiceConnections.Load(file).Profiles.All(p=>ServiceConnections.Status(p)=="Not configured")&&!Directory.Exists(root),"read/cancel path creates no file or directory");ServiceConnections.Save(file,Valid());string first=File.ReadAllText(file);Check(ServiceConnections.Encode(ServiceConnections.Load(file))==first,"file round trip");var changed=Valid();changed.Profiles[0].Deployment="second-approved-deployment";ServiceConnections.Save(file,changed);Check(ServiceConnections.Load(file).Profiles[0].Deployment=="second-approved-deployment","existing config atomically replaced");Check(Directory.GetFiles(root).Length==1,"temporary files cleaned after save");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R069 — Path_Traversal — Result 61 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **361596616**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=jMolourFcoVPzKEwVrK%2F7oi%2F7ag%3D); report line 3157 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/LinkedChartDiagnostic.cs:15` — Run, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 15 of /powerpoint-intelligence-main/tests/s0/LinkedChartDiagnostic.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ File.WriteAllText(log,"Stage\tField\tValue\r\n");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R070 — Path_Traversal — Result 62 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-756336288**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=AwSvr3czt2E%2Bu5%2Bc0uq25vj8hHI%3D); report line 3198 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/ServiceConnectionTests.cs:9` — Main, `args` |
| Sink | `tests/s0/ServiceConnectionTests.cs:31` — Run, `Exists` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 9 of /powerpoint-intelligence-main/tests/s0/ServiceConnectionTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 31 of /powerpoint-intelligence-main/tests/s0/ServiceConnectionTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
[STAThread] public static int Main(string[] args){try{Run(args.Length==0?null:args[0]);return 0;}catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
→ Check(ServiceConnections.Load(file).Profiles.All(p=>ServiceConnections.Status(p)=="Not configured")&&!Directory.Exists(root),"read/cancel path creates no file or directory");ServiceConnections.Save(file,Valid());string first=File.ReadAllText(file);Check(ServiceConnections.Encode(ServiceConnections.Load(file))==first,"file round trip");var changed=Valid();changed.Profiles[0].Deployment="second-approved-deployment";ServiceConnections.Save(file,changed);Check(ServiceConnections.Load(file).Profiles[0].Deployment=="second-approved-deployment","existing config atomically replaced");Check(Directory.GetFiles(root).Length==1,"temporary files cleaned after save");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R071 — Path_Traversal — Result 63 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1176590899**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=b8Z5%2Fzl17pnVf%2FpRjGcYwA28xQc%3D); report line 3239 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/library/TemplateLibraryTests.cs:21` — Main, `args` |
| Sink | `tests/library/TemplateLibraryTests.cs:51` — Main, `WriteAllLines` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 21 of /powerpoint-intelligence-main/tests/library/TemplateLibraryTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 51 of /powerpoint-intelligence-main/tests/library/TemplateLibraryTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
[STAThread] public static int Main(string[] args)
→ File.WriteAllLines(Path.Combine(output,"catalog.tsv"),rows);File.WriteAllLines(Path.Combine(output,"duplicate-structures.txt"),duplicates);
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R072 — Path_Traversal — Result 64 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1657118664**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=DDkLbu9wv23TSfTCvCB9zpAqpxU%3D); report line 3280 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `tests/s0/ServiceConnectionTests.cs:31` — Run, `Exists` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 31 of /powerpoint-intelligence-main/tests/s0/ServiceConnectionTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ Check(ServiceConnections.Load(file).Profiles.All(p=>ServiceConnections.Status(p)=="Not configured")&&!Directory.Exists(root),"read/cancel path creates no file or directory");ServiceConnections.Save(file,Valid());string first=File.ReadAllText(file);Check(ServiceConnections.Encode(ServiceConnections.Load(file))==first,"file round trip");var changed=Valid();changed.Profiles[0].Deployment="second-approved-deployment";ServiceConnections.Save(file,changed);Check(ServiceConnections.Load(file).Profiles[0].Deployment=="second-approved-deployment","existing config atomically replaced");Check(Directory.GetFiles(root).Length==1,"temporary files cleaned after save");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R073 — Path_Traversal — Result 65 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-753889073**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=zZAcIDQ8GfRxOBXHC0eRigaopNU%3D); report line 3321 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/library/TemplateLibraryTests.cs:21` — Main, `args` |
| Sink | `tests/library/TemplateLibraryTests.cs:51` — Main, `WriteAllLines` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 21 of /powerpoint-intelligence-main/tests/library/TemplateLibraryTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 51 of /powerpoint-intelligence-main/tests/library/TemplateLibraryTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
[STAThread] public static int Main(string[] args)
→ File.WriteAllLines(Path.Combine(output,"catalog.tsv"),rows);File.WriteAllLines(Path.Combine(output,"duplicate-structures.txt"),duplicates);
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R074 — Path_Traversal — Result 66 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **711231073**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=3dEtsmftRL1f%2BHrVLf3wbSqvQRw%3D); report line 3362 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/HostIntegrationTests.cs:50` — Main, `args` |
| Sink | `tests/s0/HostIntegrationTests.cs:95` — Main, `CreateDirectory` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 50 of /powerpoint-intelligence-main/tests/s0/HostIntegrationTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 95 of /powerpoint-intelligence-main/tests/s0/HostIntegrationTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ Directory.CreateDirectory(directory);
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R075 — Path_Traversal — Result 67 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-2135577675**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=DKXgiQ%2BLRJyqu61OG6SQ0%2FOg1qw%3D); report line 3403 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/HostIntegrationTests.cs:50` — Main, `args` |
| Sink | `tests/s0/HostIntegrationTests.cs:94` — Main, `Exists` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 50 of /powerpoint-intelligence-main/tests/s0/HostIntegrationTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 94 of /powerpoint-intelligence-main/tests/s0/HostIntegrationTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ if (Directory.Exists(directory)) throw new IOException("Use a new fixture directory.");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R076 — Path_Traversal — Result 68 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **1335748697**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=ikxSTZLx%2BghS1uqVecAS41spa%2F4%3D); report line 3444 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/DataSourceTests.cs:136` — OfficeRoundtrip, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in OfficeRoundtrip at line 136 of /powerpoint-intelligence-main/tests/s0/DataSourceTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ File.WriteAllText(Path.Combine(output,"data-sources-runtime.txt"),"PASS four source references, duplicate suppression, unchanged slide and native save/reopen. Native picker UI not exercised.");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R077 — Path_Traversal — Result 69 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **745660673**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=wEU8ETJxc7jXC53bKlfnQO0VaQE%3D); report line 3485 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/ContextPropertiesOfficeTests.cs:74` — Run, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 74 of /powerpoint-intelligence-main/tests/s0/ContextPropertiesOfficeTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ File.WriteAllText(Path.Combine(output,"context-properties-state.xml"),State(api).OuterXml);Console.WriteLine("CONTEXT "+count+" native assertions passed; Office="+app.Version+"; process="+(Environment.Is64BitProcess?"x64":"x86")+". Color-picker visuals, every chart family and per-monitor DPI still need manual coverage.");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R078 — Path_Traversal — Result 70 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **920427747**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=3FGlCeSwfv0F4SsQZF9GaRmWFHg%3D); report line 3526 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/ContextPropertiesOfficeTests.cs:70` — Run, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 70 of /powerpoint-intelligence-main/tests/s0/ContextPropertiesOfficeTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ var connector=UniversalConnectors.Create(app,new ConnectorSpec {Begin=new ConnectorEndpoint {ShapeId=text.Id,Site=1},End=new ConnectorEndpoint {ShapeId=rounded.Id,Site=1}});connector.Select();Wait(api,app,"Connector",true);Check(true,"universal connector receives its own controls");var connectorState=UniversalConnectors.Read(connector);Check(connectorState.Begin.ShapeId==text.Id&&connectorState.End.ShapeId==rounded.Id,"context connector retains native endpoint bindings");File.WriteAllText(Path.Combine(output,"context-connector-geometry.xml"),new XElement("connectorGeometry",new XAttribute("rotation",connector.Rotation),new XAttribute("width",connector.Width),new XAttribute("height",connector.Height),new XAttribute("flipH",connector.HorizontalFlip),new XAttribute("flipV",connector.VerticalFlip)).ToString());
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R079 — Path_Traversal — Result 71 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **2145587874**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=471cONyEolbOYAbpv56QuiieU3c%3D); report line 3567 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s3/SceneContractTests.cs:17` — Main, `args` |
| Sink | `tests/s3/SceneContractTests.cs:77` — Main, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 17 of /powerpoint-intelligence-main/tests/s3/SceneContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 77 of /powerpoint-intelligence-main/tests/s3/SceneContractTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ File.WriteAllText(Path.Combine(output, "results.json"), JsonConvert.SerializeObject(rows, Formatting.Indented));
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R080 — Path_Traversal — Result 72 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-2077706329**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=aekmxu6ShX26P9jtPjXB4GQ6nZg%3D); report line 3608 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Scene corpus artifact filename |
| Source | `tests/s3/SceneContractTests.cs:17` — Main, `args` |
| Sink | `tests/s3/SceneContractTests.cs:72` — Lambda, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Confirmed weakness**, open / not remediated |
| Plan mapping | **H1**, profile `CORPUS-NAME` |

**Original description:** Method Main at line 17 of /powerpoint-intelligence-main/tests/s3/SceneContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Lambda at line 72 of /powerpoint-intelligence-main/tests/s3/SceneContractTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ if (expected) { model = SceneCodec.Decode(json); File.WriteAllText(Path.Combine(output, name + ".v1.json"), SceneCodec.Encode(model)); }
```

**Root cause / controls:** SceneContractTests.cs:67-72 reads item["name"] from cases.json and combines output with name+".v1.json" without a basename/containment check. This data-derived child path differs from an intentional CLI output root.

**Exploitability / business impact:** A malicious corpus name such as ../escaped can write outside the chosen output directory when its scene JSON is valid. Side-effect-free Windows path evaluation yields C:\escaped.v1.json from C:\owned-output plus ..\escaped.v1.json. This is a verified static write primitive, not an executed exploit. Impact is runner-writable file corruption; production exposure/elevation are not demonstrated. Two reported records share this sink.

**Recommendation:** Reject invalid bounded basenames, rooted/drive/UNC/extended paths, separators, dot segments, ADS/device and trailing-dot/space forms. Enforce separator-aware canonical containment in an owned output root and no unowned-file overwrite. Reject rather than sanitize collisions; retain valid current fixture names.

**Regression risk:** C#/Node fixture naming and roundtrip validators must remain synchronized. Silent renaming may hide/collide cases.

**Required validation:** All current valid corpus names plus parent traversal, absolute/drive-relative/UNC, sibling-prefix, ADS/devices, duplicates, existing targets and junction parents. Verify no outside creation/overwrite and all valid cross-language roundtrips.


### CX-R081 — Path_Traversal — Result 73 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1426692475**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=tc3bhX5P8vsrUTLG5zn1BCrA2no%3D); report line 3649 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/library/TemplateLibraryTests.cs:21` — Main, `args` |
| Sink | `tests/library/TemplateLibraryTests.cs:26` — Main, `CreateDirectory` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 21 of /powerpoint-intelligence-main/tests/library/TemplateLibraryTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 26 of /powerpoint-intelligence-main/tests/library/TemplateLibraryTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
[STAThread] public static int Main(string[] args)
→ string output=Path.GetFullPath(args[0]);Directory.CreateDirectory(output);
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R082 — Path_Traversal — Result 74 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1599453999**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=4K6uhHw56IlTn8a2snod6w7RZqw%3D); report line 3690 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `tests/library/TemplateFidelityTests.cs:83` — Replay, `WriteAllLines` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Replay at line 83 of /powerpoint-intelligence-main/tests/library/TemplateFidelityTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ File.WriteAllLines(Path.Combine(output,"raster-replay.csv"),rows);Console.WriteLine("PASS "+passed+" saved raster comparisons. This is an offline replay, not a new native Office/formatting pass.");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R083 — Path_Traversal — Result 75 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-498998273**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=hyLs6GPJqaJt1MgB5u0YwbmGD7E%3D); report line 3731 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s3/SceneContractTests.cs:17` — Main, `args` |
| Sink | `tests/s3/SceneContractTests.cs:31` — Lambda, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 17 of /powerpoint-intelligence-main/tests/s3/SceneContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Lambda at line 31 of /powerpoint-intelligence-main/tests/s3/SceneContractTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ File.WriteAllText(Path.Combine(output, "all-primitives.v1.json"), before);
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R084 — Path_Traversal — Result 76 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **1235437117**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=ot%2BbAuXAQP5B%2FQmLzvP9A%2FBtC4w%3D); report line 3772 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s3/SceneContractTests.cs:17` — Main, `args` |
| Sink | `tests/s3/SceneContractTests.cs:67` — Main, `ReadAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 17 of /powerpoint-intelligence-main/tests/s3/SceneContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 67 of /powerpoint-intelligence-main/tests/s3/SceneContractTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ foreach (var item in JArray.Parse(File.ReadAllText(Path.Combine(corpus, "cases.json"))))
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R085 — Path_Traversal — Result 77 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **1025575657**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=XUMx5d3nRJGAVlalseCZQ%2BoCYMU%3D); report line 3813 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s3/SceneContractTests.cs:17` — Main, `args` |
| Sink | `tests/s3/SceneContractTests.cs:64` — Lambda, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 17 of /powerpoint-intelligence-main/tests/s3/SceneContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Lambda at line 64 of /powerpoint-intelligence-main/tests/s3/SceneContractTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ File.WriteAllText(Path.Combine(output, "csharp-produced.v1.json"), SceneCodec.Encode(model));
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R086 — Path_Traversal — Result 78 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1153363084**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=8hsqbMHC84tH9lPWCiJxYJFPp6I%3D); report line 3854 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/ContextPropertiesOfficeTests.cs:59` — Run, `CreateDirectory` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 59 of /powerpoint-intelligence-main/tests/s0/ContextPropertiesOfficeTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ Directory.CreateDirectory(output);evidenceDirectory=output;focusCheckpoint=0;count=0;var connections=app.COMAddIns.Cast<Office.COMAddIn>().Where(a=>(a.ProgId+" "+a.Description).Contains("PowerPointIntelligence")||(a.ProgId+" "+a.Description).Contains("PowerPoint Intelligence")).ToArray();if(connections.Length!=1||!connections[0].Connect)throw new InvalidOperationException("Install/connect the current PowerPoint Intelligence add-in before contextual UI validation.");dynamic api=connections[0].Object;string build=(string)api.GetBuildInfo();if(!build.StartsWith(DemoRibbon.BuildMarker+";",StringComparison.Ordinal))throw new InvalidOperationException("Installed add-in is stale. Install the current source before testing its automatic property box.");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R087 — Path_Traversal — Result 79 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-620214048**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=FRi1OZgfblH%2FgS3QsTHXNAvNx5o%3D); report line 3895 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/ContextFormattingOfficeTests.cs:9` — Run, `CreateDirectory` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 9 of /powerpoint-intelligence-main/tests/s0/ContextFormattingOfficeTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ internal static void Run(PowerPoint.Application app,string outputDirectory){count=0;Directory.CreateDirectory(outputDirectory);string file=Path.Combine(Path.GetFullPath(outputDirectory),"context-font-"+Guid.NewGuid().ToString("N")+".pptx");PowerPoint.Presentation deck=null;try{deck=app.Presentations.Add(Office.MsoTriState.msoTrue);deck.PageSetup.SlideWidth=960;deck.PageSetup.SlideHeight=540;deck.Windows[1].ViewType=PowerPoint.PpViewType.ppViewNormal;var slide=deck.Slides.Add(1,PowerPoint.PpSlideLayout.ppLayoutBlank);slide.Select();var a=slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,50,50,250,70);a.TextFrame.TextRange.Text="ABCD";a.TextFrame.TextRange.Font.Name="Arial";a.TextFrame.TextRange.Font.Size=18;a.TextFrame.TextRange.Font.Bold=Office.MsoTriState.msoFalse;a.TextFrame.TextRange.Font.Italic=Office.MsoTriState.msoFalse;a.TextFrame.TextRange.Characters(2,1).Font.Bold=Office.MsoTriState.msoTrue;a.Select();var read=ContextFormatting.Read(app);Check(read.Family=="Arial"&&read.Size==18&&read.Bold==null&&read.Italic==false,"read reports per-property mixed state without mutation");a.TextFrame.TextRange.Characters(2,2).Select();ContextFormatting.ApplyFont(app,"Calibri",24,null,true);Check(a.TextFrame.TextRange.Characters(1,1).Font.Name=="Arial"&&a.TextFrame.TextRange.Characters(2,1).Font.Name=="Calibri"&&a.TextFrame.TextRange.Characters(4,1).Font.Size==18,"selected-character scope preserved");Check(a.TextFrame.TextRange.Characters(2,1).Font.Bold==Office.MsoTriState.msoTrue&&a.TextFrame.TextRange.Characters(3,1).Font.Bold==Office.MsoTriState.msoFalse,"unrequested mixed bold formatting retained");var b=slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle,350,50,200,70);b.TextFrame.TextRange.Text="Second";var group=slide.Shapes.Range(new object[]{a.Name,b.Name}).Group();int groupId=group.Id;group.Select();ContextFormatting.ApplyFont(app,"Arial",20,true,false);Check(group.GroupItems[1].TextFrame.TextRange.Font.Size==20&&group.GroupItems[2].TextFrame.TextRange.Font.Bold==Office.MsoTriState.msoTrue,"bounded native group targets formatted");var table=slide.Shapes.AddTable(2,2,50,200,300,140);for(int r=1;r<=2;r++)for(int c=1;c<=2;c++){table.Table.Cell(r,c).Shape.TextFrame.TextRange.Text="Cell";table.Table.Cell(r,c).Shape.TextFrame.TextRange.Font.Size=12;}table.Select();ContextFormatting.ApplyFont(app,null,16,null,null);Check(table.Table.Cell(1,1).Shape.TextFrame.TextRange.Font.Size==16&&table.Table.Cell(2,2).Shape.TextFrame.TextRange.Font.Size==16,"native table whole selection formatted");var decoration=slide.Shapes.AddLine(600,50,650,50);var textLeaf=slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,600,80,200,50);textLeaf.TextFrame.TextRange.Text="Diagram label";var diagram=slide.Shapes.Range(new object[]{decoration.Name,textLeaf.Name}).Group();diagram.Select();ContextFormatting.ApplyFont(app,null,22,null,null);Check(diagram.GroupItems[2].TextFrame.TextRange.Font.Size==22,"group nontext decoration skipped while editable text formatted");var line=slide.Shapes.AddLine(10,450,100,450);slide.Shapes.Range(new object[]{group.Name,line.Name}).Select();bool refused=false;try{ContextFormatting.ApplyFont(app,null,32,null,null);}catch(InvalidOperationException){refused=true;}Check(refused&&group.GroupItems[1].TextFrame.TextRange.Font.Size==20,"unsupported late target rejected before any mutation");deck.SaveAs(file,PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,Office.MsoTriState.msoFalse);deck.Close();deck=null;deck=app.Presentations.Open(file,Office.MsoTriState.msoFalse,Office.MsoTriState.msoFalse,Office.MsoTriState.msoTrue);group=deck.Slides[1].Shapes.Cast<PowerPoint.Shape>().Single(x=>x.Id==groupId);Check(group.GroupItems[1].TextFrame.TextRange.Font.Size==20&&group.GroupItems[2].TextFrame.TextRange.Font.Bold==Office.MsoTriState.msoTrue,"native font edit survives save reopen");Console.WriteLine("FONT "+count+" native assertions passed. Evidence deck: "+file);}finally{if(deck!=null){deck.Saved=Office.MsoTriState.msoTrue;deck.Close();}}}
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R088 — Path_Traversal — Result 80 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1926027703**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=jjxqKsSaULiHMq%2BPn8sSnE580Ik%3D); report line 3936 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/WorkflowOfficeTests.cs:42` — Run, `CreateDirectory` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 42 of /powerpoint-intelligence-main/tests/s0/WorkflowOfficeTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ Directory.CreateDirectory(outputDirectory);assertions=0;
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R089 — Path_Traversal — Result 81 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1549767052**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=RsfMIoGfocZU4KpXRx6wqeJWXMk%3D); report line 3977 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/ChapterAgendaOfficeTests.cs:21` — Run, `CreateDirectory` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 21 of /powerpoint-intelligence-main/tests/s0/ChapterAgendaOfficeTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ if(app==null)throw new ArgumentNullException("app");Directory.CreateDirectory(outputDirectory);count=0;
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R090 — Path_Traversal — Result 82 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1612927088**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=5VMPhyUXxS1okvs98kmvK7AY%2FUw%3D); report line 4018 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s3/SceneContractTests.cs:17` — Main, `args` |
| Sink | `tests/s3/SceneContractTests.cs:22` — Main, `CreateDirectory` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 17 of /powerpoint-intelligence-main/tests/s3/SceneContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 22 of /powerpoint-intelligence-main/tests/s3/SceneContractTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ Directory.CreateDirectory(output);
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R091 — Path_Traversal — Result 83 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **1144353721**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=JmJIwEeIREBm3srwqPPWH4pJiHQ%3D); report line 4059 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Office compatibility configuration |
| Source | `tests/compatibility/OfficeCompatibilityTests.cs:19` — Main, `args` |
| Sink | `apps/powerpoint-vsto/OfficeCompatibility.cs:33` — Load, `Exists` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `OFFICE-CONFIG` |

**Original description:** Method Main at line 19 of /powerpoint-intelligence-main/tests/compatibility/OfficeCompatibilityTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Load at line 33 of /powerpoint-intelligence-main/apps/powerpoint-vsto/OfficeCompatibility.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ if (!File.Exists(configurationPath)) return Parse(new NameValueCollection());
```

**Root cause / controls:** The source is a compatibility-test argument. Production maps the deployed add-in configuration path and never writes POWERPNT.EXE.config; only approved Office setting names/profiles/versions are accepted.

**Exploitability / business impact:** Explicitly mapped test configuration is not proven unauthorized file access. Deployed-config tampering is a package provenance question, not an access-control exploit established by File.Exists.

**Recommendation:** Keep deployed config authority distinct from selected sources. Validate caller/context if a shared path policy is introduced; preserve intended installation paths and profile parsing.

**Regression risk:** Office 2016/2019/365 and x86/x64/build gating can regress under blanket path restrictions.

**Required validation:** Compatibility matrix, absent defaults, unknown settings, malformed profiles, hostile caller path under defined policy and no POWERPNT.EXE.config writes.


### CX-R092 — Path_Traversal — Result 84 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **1398503197**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=lpqBYhbC4y96BQ9wG4eGS0y563o%3D); report line 4100 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `tests/s0/DataSourceTests.cs:18` — Run, `CreateDirectory` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 18 of /powerpoint-intelligence-main/tests/s0/DataSourceTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ Directory.CreateDirectory(output);var registryPath=Path.Combine(Path.GetFullPath(output),"registry-fixture.csv");File.WriteAllText(registryPath,"Category,Value\r\nFixture,1\r\n");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R093 — Path_Traversal — Result 85 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-457272103**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=zRMwGMxTzxZuAf8N5Pw6VnIyf0I%3D); report line 4141 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/LinkedChartDiagnostic.cs:14` — Run, `CreateDirectory` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 14 of /powerpoint-intelligence-main/tests/s0/LinkedChartDiagnostic.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ Directory.CreateDirectory(output);string log=Path.Combine(output,"linked-chart-stages.tsv");
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R094 — Path_Traversal — Result 86 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-293388255**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=GYCbZ2p3OoIRCFfYdJb3oDvGp6k%3D); report line 4182 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/DemoOfficeTests.cs:79` — Main, `WriteAllLines` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 79 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ rows.Add(tool.Id+","+outcome+","+evidence);File.WriteAllLines(Path.Combine(args[1],"features.csv"),rows);
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R095 — Path_Traversal — Result 87 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-771679709**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=q%2FtPtV5CM40JXH2X77HaUA%2FezsA%3D); report line 4223 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `tests/s0/ContextPropertiesTests.cs:17` — Run, `CreateDirectory` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Run at line 17 of /powerpoint-intelligence-main/tests/s0/ContextPropertiesTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ Directory.CreateDirectory(output);int count=0;Action<bool,string> check=(ok,name)=>{if(!ok)throw new Exception("FAIL property box: "+name);count++;};
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R096 — Path_Traversal — Result 88 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-661553133**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=TsnHpR17qMgxNio82XXd8xjBguo%3D); report line 4264 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `tests/library/TemplateFidelityTests.cs:81` — Replay, `CreateDirectory` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Replay at line 81 of /powerpoint-intelligence-main/tests/library/TemplateFidelityTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ Directory.CreateDirectory(output);var rows=new List<string>{"Index,Stage,SameDimensions,ExactPixels,WithinTolerance,ChangedPixels,MaxChannelDelta,AlphaChanges,NonEdgeChanges,AllowedPixels"};int passed=0;
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R097 — Path_Traversal — Result 89 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **1201223223**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=xZhYq8GB51IkJmCQGbAimvRUOuY%3D); report line 4305 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s3/NativeRendererTests.cs:48` — Main, `args` |
| Sink | `tests/s3/NativeRendererTests.cs:111` — Lambda, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 48 of /powerpoint-intelligence-main/tests/s3/NativeRendererTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Lambda at line 111 of /powerpoint-intelligence-main/tests/s3/NativeRendererTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ File.WriteAllText(Path.Combine(args[2], "rendered-chart.v1.xml"), xml);
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R098 — Path_Traversal — Result 90 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **596844632**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=LIF4bHZBVq9b9FFUzuqOC5hyYmQ%3D); report line 4346 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/DemoOfficeTests.cs:40` — Main, `WriteAllLines` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 40 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ if(agendaOnly){try{ChapterAgendaOfficeTests.Run(app,args[1]);rows.Add("chapter-agenda-native,PASS,Native chapter hierarchy edits moves copy split and persistence");File.WriteAllLines(Path.Combine(args[1],"features.csv"),rows);return 0;}catch(Exception error){rows.Add("chapter-agenda-native,FAIL,"+error.GetType().Name);File.WriteAllLines(Path.Combine(args[1],"features.csv"),rows);Console.Error.WriteLine(error);return 1;}}
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R099 — Path_Traversal — Result 91 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **1018546989**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=%2Bp%2FD6MQp3SgVgaalyBnq3nTo%2Feo%3D); report line 4387 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/DemoOfficeTests.cs:46` — Main, `WriteAllLines` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 46 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ File.WriteAllLines(Path.Combine(args[1],"features.csv"),rows);
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R100 — Path_Traversal — Result 92 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **596844632**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=SJadvNRDG2JZnK5sXpRQsf8cw%2FA%3D); report line 4428 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/DemoOfficeTests.cs:40` — Main, `WriteAllLines` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 40 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ if(agendaOnly){try{ChapterAgendaOfficeTests.Run(app,args[1]);rows.Add("chapter-agenda-native,PASS,Native chapter hierarchy edits moves copy split and persistence");File.WriteAllLines(Path.Combine(args[1],"features.csv"),rows);return 0;}catch(Exception error){rows.Add("chapter-agenda-native,FAIL,"+error.GetType().Name);File.WriteAllLines(Path.Combine(args[1],"features.csv"),rows);Console.Error.WriteLine(error);return 1;}}
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R101 — Path_Traversal — Result 93 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **596844632**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=wLIZfwKBNpfOhGyE4HmLr1Bb%2Bww%3D); report line 4469 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/DemoOfficeTests.cs:59` — Main, `WriteAllLines` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 59 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ if(disconnected){blockedCommands++;rows.Add(tool.Id+",BLOCKED,Prior PowerPoint RPC disconnection; command not executed");File.WriteAllLines(Path.Combine(args[1],"features.csv"),rows);continue;}
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R102 — Path_Traversal — Result 94 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **341466547**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=1qztiZPnmmUR6aEVEP00QBEeBRw%3D); report line 4510 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/DemoOfficeTests.cs:35` — Main, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 35 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ try{MappedExcelOfficeTests.Run(app,args[1]);File.WriteAllText(Path.Combine(args[1],"mapped-excel-result.txt"),"PASS");return 0;}
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R103 — Path_Traversal — Result 95 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **596844632**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=eDvX2ZRO9yiIFGdYmTbKZfe1Pjw%3D); report line 4551 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/DemoOfficeTests.cs:39` — Main, `WriteAllLines` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 39 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ if(args[0]=="--interactive-properties"){try{ContextPropertiesOfficeTests.Run(app,args[1]);ContextFormattingOfficeTests.Run(app,args[1]);rows.Add("context-properties-native,PASS,Installed automatic property box scope native edits and save reopen");File.WriteAllLines(Path.Combine(args[1],"features.csv"),rows);return 0;}catch(Exception error){rows.Add("context-properties-native,"+(error is ContextTestForegroundUnavailableException?"BLOCKED":"FAIL")+","+error.GetType().Name);File.WriteAllLines(Path.Combine(args[1],"features.csv"),rows);Console.Error.WriteLine(error);return 1;}}
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R104 — Path_Traversal — Result 96 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **596844632**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=L4rikjAUUJIu%2B%2F6pcrdmjmLVOUY%3D); report line 4592 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/DemoOfficeTests.cs:39` — Main, `WriteAllLines` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 39 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ if(args[0]=="--interactive-properties"){try{ContextPropertiesOfficeTests.Run(app,args[1]);ContextFormattingOfficeTests.Run(app,args[1]);rows.Add("context-properties-native,PASS,Installed automatic property box scope native edits and save reopen");File.WriteAllLines(Path.Combine(args[1],"features.csv"),rows);return 0;}catch(Exception error){rows.Add("context-properties-native,"+(error is ContextTestForegroundUnavailableException?"BLOCKED":"FAIL")+","+error.GetType().Name);File.WriteAllLines(Path.Combine(args[1],"features.csv"),rows);Console.Error.WriteLine(error);return 1;}}
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R105 — Path_Traversal — Result 97 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **1855789791**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=wKU1KVx51CJxOeGBkBXbOl9d2iA%3D); report line 4633 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/DemoOfficeTests.cs:36` — Main, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 36 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ catch(Exception error){File.WriteAllText(Path.Combine(args[1],"mapped-excel-result.txt"),"FAIL "+error.GetType().Name+" 0x"+error.HResult.ToString("X8"));Console.Error.WriteLine(error);return 1;}
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R106 — Path_Traversal — Result 98 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1574684448**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=94QLNu5n%2BzDGR1XQ1d6o3fniFEw%3D); report line 4674 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s3/NativeRendererTests.cs:48` — Main, `args` |
| Sink | `tests/s3/NativeRendererTests.cs:53` — Main, `ReadAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 48 of /powerpoint-intelligence-main/tests/s3/NativeRendererTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 53 of /powerpoint-intelligence-main/tests/s3/NativeRendererTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ canonical = File.ReadAllText(args[0]); scene = File.ReadAllText(args[1]);
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R107 — Path_Traversal — Result 99 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **360136018**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=EQVW2i8iB6DJE6t3nX7kXA3NhY8%3D); report line 4715 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `tests/s0/DemoContractTests.cs:72` — Main, `WriteAllLines` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 72 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ File.WriteAllLines(Path.Combine(args[1],"demo-tool-ids.txt"),tools.Select(t=>t.Id));
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R108 — Path_Traversal — Result 100 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **596844632**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=I%2B75%2BQNr3HYFnmzuonQGJJYRDew%3D); report line 4756 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/DemoOfficeTests.cs:31` — Main, `WriteAllLines` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 31 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ catch(Exception ex){if(args[0]=="--interactive-properties")rows.Add("context-properties-native,BLOCKED,"+ex.GetType().Name);else if(agendaOnly)rows.Add("chapter-agenda-native,BLOCKED,"+ex.GetType().Name);else if(args[0]=="--interactive-elements")foreach(var id in new[]{"productivity-native","text-elements-native","process-grid-native","indicators-native","universal-connectors-native"})rows.Add(id+",BLOCKED,"+ex.GetType().Name);else foreach(var tool in DemoTools.Catalog)rows.Add(tool.Id+",BLOCKED,"+ex.GetType().Name);File.WriteAllLines(Path.Combine(args[1],"features.csv"),rows);Console.Error.WriteLine(ex);return 1;}
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R109 — Path_Traversal — Result 101 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-2135358301**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=wdVhyy5Ln1tIpocgOMvWeU7PyIA%3D); report line 4797 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s3/NativeRendererTests.cs:48` — Main, `args` |
| Sink | `tests/s3/NativeRendererTests.cs:53` — Main, `ReadAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 48 of /powerpoint-intelligence-main/tests/s3/NativeRendererTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 53 of /powerpoint-intelligence-main/tests/s3/NativeRendererTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ canonical = File.ReadAllText(args[0]); scene = File.ReadAllText(args[1]);
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R110 — Path_Traversal — Result 102 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **661522298**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=CBaCfq8Qh%2FrW2QyvHfwoyS8rOSw%3D); report line 4838 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s1/DomainTests.cs:18` — Main, `args` |
| Sink | `tests/s1/DomainTests.cs:22` — Main, `ReadAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 18 of /powerpoint-intelligence-main/tests/s1/DomainTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 22 of /powerpoint-intelligence-main/tests/s1/DomainTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ golden = File.ReadAllText(args[0]);
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R111 — Path_Traversal — Result 103 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **793367280**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=kPBPyWXbDxV5KJB5krneuHhe8lY%3D); report line 4879 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s3/SceneContractTests.cs:17` — Main, `args` |
| Sink | `tests/s3/SceneContractTests.cs:21` — Main, `ReadAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 17 of /powerpoint-intelligence-main/tests/s3/SceneContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 21 of /powerpoint-intelligence-main/tests/s3/SceneContractTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ string golden = File.ReadAllText(args[0]), output = args[1], corpus = args[2];
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R112 — Path_Traversal — Result 104 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **1739815605**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=o7cgfuoPlTuwxziEBeECQFd7x1o%3D); report line 4920 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoOfficeTests.cs:24` — Main, `args` |
| Sink | `tests/s0/DemoOfficeTests.cs:28` — Main, `CreateDirectory` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 24 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 28 of /powerpoint-intelligence-main/tests/s0/DemoOfficeTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ Directory.CreateDirectory(args[1]);var rows=new List<string>{"ToolId,Status,Evidence"};
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R113 — Path_Traversal — Result 105 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1880250405**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=SD6e4EaEooYdJl1z3aINa8nXO2I%3D); report line 4961 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s1/DomainTests.cs:18` — Main, `args` |
| Sink | `tests/s1/DomainTests.cs:27` — Lambda, `WriteAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 18 of /powerpoint-intelligence-main/tests/s1/DomainTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Lambda at line 27 of /powerpoint-intelligence-main/tests/s1/DomainTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ File.WriteAllText(args[1], encoded);
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R114 — Path_Traversal — Result 106 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-169465736**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=5qK5FqnMohLhaIy3U2g%2BSl1Ujn4%3D); report line 5002 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `tests/s0/DemoContractTests.cs:67` — Main, `ReadAllText` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 67 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ string source=File.ReadAllText(args[0]);
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R115 — Path_Traversal — Result 107 of 107

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **724762689**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=FKYh5I37GKd7ysP6u0eELJvUzYQ%3D); report line 5043 |
| Original severity / CWE / CVE | **High** / **CWE-22** / no CVE supplied |
| Component | Developer/native fixtures and evidence |
| Source | `tests/s0/DemoContractTests.cs:11` — Main, `args` |
| Sink | `tests/s0/DemoContractTests.cs:70` — Main, `CreateDirectory` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Potential false positive**, open / not remediated |
| Plan mapping | **H3**, profile `TEST-ARTIFACT` |

**Original description:** Method Main at line 11 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs gets dynamic data from the args element. This element’s value then flows through the code and is eventually used in a file path for local disk access in Main at line 70 of /powerpoint-intelligence-main/tests/s0/DemoContractTests.cs. This may cause a Path Traversal vulnerability.

**Observed source → sink:**

```text
public static int Main(string[] args)
→ Directory.CreateDirectory(args[1]);document.Save(Path.Combine(args[1],"demo-ribbon.xml"));
```

**Root cause / controls:** The test executable intentionally accepts fixture input/output directories through CLI arguments and derives fixed or internally generated names. The reviewed entrypoint has no server, elevation, delegated authorization or lower-trust production request. This classification excludes the separately confirmed data-derived corpus name.

**Exploitability / business impact:** The caller already running this test can choose paths with its own token; no crossed authority boundary is demonstrated. CI consuming untrusted arguments or shared/junction directories would require additional controls. Existing-file evidence can be corrupted if workspace ownership assumptions fail.

**Recommendation:** Consider an owned per-run workspace, bounded fixture inputs, safe child names and create-new/atomic outputs. Validate before directory creation/Office startup/cleanup. Preserve positional CLI contracts and stable runner paths; do not remove tests or exclude files from scanning.

**Regression risk:** Shell wrappers, goldens, foreground handoff, fidelity evidence and cross-language validation depend on current contracts. Never close user presentations as cleanup.

**Required validation:** Current scripts, Unicode/spaces/long paths, argument counts, missing fixtures, existing files, output ownership, sibling-prefix/junction cases, nonzero failure exits and all domain/scene/template/native roundtrips.


### CX-R116 — Unchecked_Input_for_Loop_Condition — Result 1 of 2

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-900544832**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=8XU40UjF1fMUxkR7L4gn1x95Re4%3D); report line 5100 |
| Original severity / CWE / CVE | **Medium** / **CWE-606** / no CVE supplied |
| Component | Offline PPTX repair resource use |
| Source | `scripts/Repair-TemplateAlignment.py:48` — repair, `read` |
| Sink | `scripts/Repair-TemplateAlignment.py:31` — fix_slide, `bodies` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Confirmed weakness**, open / not remediated |
| Plan mapping | **M1**, profile `ZIP-REPAIR` |

**Original description:** Method repair at line 48 of /powerpoint-intelligence-main/scripts/Repair-TemplateAlignment.py obtains user input from read - the range of this value is not validated, and is eventually used in a loop condition in fix_slide at line 31 of /powerpoint-intelligence-main/scripts/Repair-TemplateAlignment.py, allowing attackers to provide a very high number of iterations.

**Observed source → sink:**

```text
data=zin.read(info.filename)
→ for body in bodies:
```

**Root cause / controls:** Repair-TemplateAlignment.py reads every ZIP entry fully and materializes XML/XPath sets without archive/member/uncompressed-byte/shape/work/time budgets. Both transformation and comparison passes depend on document content. The safe parser at line 19 blocks entities/network but not aggregate decompression or quadratic container/label matching.

**Exploitability / business impact:** Repairing a malicious/oversized PPTX can exhaust memory/CPU or leave partial output. This is an offline utility weakness, not a demonstrated remote server DoS. Two Medium records share the same missing-budget design.

**Recommendation:** Preflight unique normalized ZIP members, count, sizes and ratios; enforce actual streamed bytes and XML depth/node/text/shape/aggregate-work limits in both passes. Use explicit safe parsers, cancellation and transactional new deck/sidecar publication; calibrate limits on valid decks.

**Regression risk:** Preserve valid large decks within documented limits, source bytes, text, native shapes, timing, transitions and untouched ZIP parts; never silently truncate.

**Required validation:** Small bounded high-compression, oversized/deep XML, excess entries/shapes, duplicate/truncated ZIP, cancellation and quadratic-work fixtures. Reject without partial publication; unchanged input and accepted-deck native reopen.


### CX-R117 — Unchecked_Input_for_Loop_Condition — Result 2 of 2

| Field | Assessment |
|---|---|
| Original identity | Similarity ID **-1745516950**; [original result](https://ind-2.ast.checkmarx.net/sast-results/f28c8294-b4b8-45e7-8792-500ce1e8d393/bed87600-7c2c-45b7-8b17-15b0b268ec82?resultId=7RIaiQLJLCtZFhGXKASGDfPI%2FWw%3D); report line 5141 |
| Original severity / CWE / CVE | **Medium** / **CWE-606** / no CVE supplied |
| Component | Offline PPTX repair resource use |
| Source | `scripts/Repair-TemplateAlignment.py:56` — repair, `read` |
| Sink | `scripts/Repair-TemplateAlignment.py:62` — repair, `xpath` |
| Correspondence | Both reported snippets match the current reported lines |
| Assessment | **Confirmed weakness**, open / not remediated |
| Plan mapping | **M1**, profile `ZIP-REPAIR` |

**Original description:** Method repair at line 56 of /powerpoint-intelligence-main/scripts/Repair-TemplateAlignment.py obtains user input from read - the range of this value is not validated, and is eventually used in a loop condition in repair at line 62 of /powerpoint-intelligence-main/scripts/Repair-TemplateAlignment.py, allowing attackers to provide a very high number of iterations.

**Observed source → sink:**

```text
old=E.fromstring(before.read(name));new=E.fromstring(after.read(name))
→ assert all(p.get('algn')=='ctr' for p in new.xpath('//p:txBody/a:p/a:pPr | //a:txBody/a:p/a:pPr',namespaces=NS))
```

**Root cause / controls:** Repair-TemplateAlignment.py reads every ZIP entry fully and materializes XML/XPath sets without archive/member/uncompressed-byte/shape/work/time budgets. Both transformation and comparison passes depend on document content. The safe parser at line 19 blocks entities/network but not aggregate decompression or quadratic container/label matching.

**Exploitability / business impact:** Repairing a malicious/oversized PPTX can exhaust memory/CPU or leave partial output. This is an offline utility weakness, not a demonstrated remote server DoS. Two Medium records share the same missing-budget design.

**Recommendation:** Preflight unique normalized ZIP members, count, sizes and ratios; enforce actual streamed bytes and XML depth/node/text/shape/aggregate-work limits in both passes. Use explicit safe parsers, cancellation and transactional new deck/sidecar publication; calibrate limits on valid decks.

**Regression risk:** Preserve valid large decks within documented limits, source bytes, text, native shapes, timing, transitions and untouched ZIP parts; never silently truncate.

**Required validation:** Small bounded high-compression, oversized/deep XML, excess entries/shapes, duplicate/truncated ZIP, cancellation and quadratic-work fixtures. Reject without partial publication; unchanged input and accepted-deck native reopen.


## Summary-only original SCS finding — blocked

This is a coverage slot for the **one original High SCS result**, engine `2ms`, not a fabricated finding ID. Original ID, CWE/CVE, source file/line, rule, result URL, evidence and trace: **not supplied**. Exploitability, root cause and false-positive status: **unknown**. Exact fix, regression risk and validation cannot responsibly be specified before identifying the affected asset. Assessment: **Blocked—missing report details**.

Request a redacted machine-readable SCS result with original ID, scanner/rule, location and evidence, plus SAST rule/result identity to reconcile the overview mismatch. Do not paste raw secrets into source, logs or these files. If a credential is actually identified, approve appropriate incident/rotation handling rather than assuming this aggregate means a leaked token. This report is located unambiguously; assessment of this individual finding remains blocked. Keep it High/open. No blanket exclusions or suppression.

## Additional review findings — separate from Checkmarx counts

These are **not** original scanner IDs and do not increase the original 118. Severities/CWEs here are proposed reviewer priorities, not Checkmarx classifications. No CVE is asserted.

### AR-01 — Dependency download integrity — proposed High, CWE-494 / CWE-353

Evidence: `scripts/Test-S1.ps1:16-24` downloads Newtonsoft.Json **13.0.3** over HTTPS, expands it without an independently pinned package/DLL digest or signature verification, then uses it for compilation/tests and the production project's HintPath. Source-wide network search identified this direct network implementation. **Confirmed missing independent integrity control; exploitability conditional on supply-chain/artifact compromise.** HTTPS and a fixed version are positive controls; no malicious package, interception or executed compromise is demonstrated.

Fix proposal H4: independently establish approved package identity/hash/signature data, verify before extraction/copy/load and on cache reuse; stage atomically and reject unsafe archive members. Keep 13.0.3 unless a separate dependency change is approved. Never invent an expected hash or trust a checksum obtained from the same untrusted response. Regression risk: offline restore/test behavior and downstream DLL references. Tests: bad ZIP/DLL, cache corruption, partial download, archive traversal, exact existing-version restoration and approved offline cache. No dependency was upgraded or downloaded in Run 1.

### AR-02 — Physical filesystem identity / races — proposed Medium, CWE-59 / CWE-367 candidate

Evidence: `src/Library.Core/PiidxCatalog.cs:33-44,59-60`, `PresentationSource.cs:58-70,126-135`, `apps/powerpoint-vsto/ServiceConnections.cs:94-95`, `easySetup/Bootstrapper.cs:15-18,59-67`. Lexical prefix/normalization and leaf reparse checks precede separate opens/writes/extraction. These do not prove the actual target's ancestry or prevent a directory swap. **Conditional defense gap**: a hostile actor must be able to redirect an ancestor or win a race across a meaningful ownership boundary; that attacker/ACL condition is not established.

Fix proposal M2: define source versus owned-output authority, approved redirected/UNC paths and physical identity requirements. Use suitable Windows final-path/handle identity, controlled directory ownership, create-new/atomic operations and bounded reparse policy where needed. Do not reject every legitimate redirected profile. Regression risk: shared libraries, roaming profiles, index journals and per-user setup. Tests: disposable Windows junction ancestors, aliases, races, existing targets and ordinary approved paths; pure lexical tests do not establish junction safety. No race exploit was executed.

### AR-03 — Repair publication and validation — proposed Low, CWE-754 defense-in-depth

Evidence: `scripts/Repair-TemplateAlignment.py:43,46,53-66` checks output existence then creates with write mode, uses `assert` for preservation, and writes the `.validation.json` sidecar without checking its prior existence. **Confirmed defensive correctness gaps; security exploitability not established.** Optimized Python removes assert checks; a pre-existing sidecar can be overwritten and exceptions can leave partial output.

Fix proposal L1 with M1: explicit checked exceptions, create-new temporary deck/sidecar and transactional publication only after successful validation. No unrelated sidecar overwrite. Preserve original bytes/text/timing/transitions. Tests: `python -O`, induced preservation mismatch, sidecar collision, output race/cancellation and clean failure without a success report. This is not a claim of remote code execution.

## Repository-wide impact and positive controls

The complete current source/deployment/test/configuration tree was searched, not only scanner-listed files: 217 code/config/dependency files were fingerprinted, common filesystem operations appear in 45 source files, and all 32 distinct reported source/sink files were inspected at the relevant paths. Generated snapshots/third-party installations are not independent production modules.

| Module | Evidence / controls reviewed | Implication and residual scope |
|---|---|---|
| Ribbon / initialization | `DemoRibbon.cs:11-17`, `ThisAddIn.cs:37-75,133-158`, `OfficeCompatibility.cs:30-45` | Deployed add-in config, UI dispatch/mutation guards and bounded Office profile handling. Preserve all commands, panes and STA behavior; no extra exposed server endpoint was identified. |
| Native rendering / COM | `src/Chart.Renderers/NativeRenderingCore.cs`, `PowerPointSceneSurface.cs`, `PowerPointHost.cs`, `Context*`, `XmlPartUpdates.cs` | Typed scene/model validation, capture/transaction and bounded metadata controls. Same-user COM automation is intentional, not proof of a network auth bypass. Native editability/foreground/save-reopen still need Windows regression. |
| Charts / data / shapes | `DataSources.cs:94-130,283-294`, `ExcelRangeMapping.cs:51-64`, `DemoLinks.cs:49-70` | Explicit mapping, size/schema validation and no moved-file fallback. Document-carried source consent, UNC policy and source identity changes remain distinct questions. Legacy DemoLinks lacks the explicit AutomationSecurity assignment used by modern readers; review/test negative macro/link behavior if exposed, without asserting demonstrated macro execution. |
| Template library | `TemplateLibrary.cs:345-354,444-485`, `LocalLibraryWorkspace.cs:73-83` | Bounded XML, hidden readonly source access and source snapshot checks. Preserve current presentation, native originals, thumbnail rendering, palettes/alignment and optional timing/builds. Authorized external roots are legitimate. |
| PIIDX/indexing | `LibraryConfiguration.cs:26-87`, `PiidxCatalog.cs:16-70,107-114`, `PresentationSource.cs:56-92,126-158`, `PiidxPreviewCache.cs:18-39,64-72`, `NativeSqlite.cs:27-48,75-95` | Parameterized SQL, trusted_schema=OFF, System32-only winsqlite3, PNG validation, limits/cancellation and logical digests are positive controls. Unkeyed digests detect corruption, not authorship. Preserve concurrent/readonly indexing and migrations; AR-02 covers conditional physical ownership. |
| Azure AI / external services / auth | `ServiceConnections.cs:28-29,44-55,74-95`; no production HttpClient/WebClient/fetch implementation in searched source | Nonsecret configuration only; no implemented authenticated backend/live Azure traffic was found. Do not invent SSRF or authorization bypass from an HTTPS configuration field. Q-003/Q-004/Q-005 remain enterprise integration/identity/data-policy decisions. |
| Parsing / filesystem | Test workspaces, source registries, ZIP/XML/JSON, diagnostics and setup | Distinguish caller-selected source files from owned generated output. DTD prohibition / null XmlResolver and JSON TypeNameHandling.None/MaxDepth are controls; no unsafe BinaryFormatter/polymorphic-deserializer signature was found in searched source. This is not a guarantee for dependencies. |
| Dependencies | Newtonsoft.Json 13.0.3 fallback restore; npm package locks in chart-domain/chart-engine; AJV 8.20.0 / TypeScript 5.9.3 in chart-engine | AR-01 is provenance, not an invented CVE. Zero filtered SCA findings is insufficient assurance. No dependency upgrade/install/new SCA scan was performed. |
| Installation / deployment | `easySetup/Build-EasySetup.ps1:9-35,75-87`, `Bootstrapper.cs:15-18,34-78`, `Install-Local.ps1`, `Test-DemoPackage.ps1`, runner cache scripts | Existing signature/hash, per-user, versioning and scanner-review controls remain required. Do not remove signing gates, trust a new root automatically, bypass execution controls or add antivirus exclusions to obtain a pass. |

## Limitations / stop gate

This is source assessment, not a fresh Checkmarx scan or executed exploit. The complete scanner ZIP/commit binding, raw graph intermediates, query/rules version, individual SCS details, CI/cache ACL ownership model and target security policy are unavailable. The report's general server-folder advice cannot be applied blindly to a desktop add-in that intentionally reads selected external files. Functional historical Office passes are not security-remediation proof. **All findings remain open for Run 2 approval; no implementation starts in Run 1.**

## Assessment integrity verification

Documentation checks verified 117 unique assessment aliases, all 117 original distinct result URLs, every repeated Similarity ID, all 234 source/sink references, complete plan/handoff mapping, balanced code fences and the 4 + 3 + 110 detailed-record classification totals. The separate SCS summary slot accounts for the remaining original reported result without inventing its ID. All **217 baseline application/tool/configuration/dependency files remain byte-identical**. The source report's raw SHA-256 remains unchanged. These are assessment-integrity checks, not application security tests or proof that any finding was fixed. Only the three requested security Markdown documents were written in the repository.
