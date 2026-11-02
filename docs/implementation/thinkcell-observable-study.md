# think-cell observable library study and independent implementation

Date: 26 October 2026. Owner priority: **slide library, layout and template controls**.

## Scope and boundaries

The owner explicitly invoked the REA skill. Repository AGENTS.md rule 4 prohibits reverse-engineering or copying think-cell binaries, proprietary assets or protected implementation details. This study therefore uses **observable UI and published documentation**. No think-cell DLL was decompiled or instrumented, and no vendor templates/assets were imported into our library. The installed product version was read from file version metadata only. Source analysis of our repository uses ordinary code tools, not decompilation.

REA's native-UI observation workflow in this tool version is macOS-only. On this Windows machine, native UI evidence was captured using the computer-use skill's supported Windows accessibility/screenshot API. REA was used for execution-free inspection of our shipped managed add-in, establishing an exact binary baseline. These are different evidence authorities and must not be conflated.

## Evidence you can review

| ID | Evidence | Observation / limitation |
|---|---|---|
| UI-01 | `artifacts/thinkcell-study-20261026/thinkcell-library-0.png` | PowerPoint with the think-cell Library open. Captured from the installed application, not a mockup. |
| UI-02 | `artifacts/thinkcell-study-20261026/thinkcell-library-2.png` | Library detail showing Slide Templates (497), hierarchical folders, search box, grouped thumbnails and three hover icons. Icon meanings were not confirmed through tooltips in this session. |
| UI-03 | `artifacts/thinkcell-study-20261026/thinkcell-library-uia.txt` | Accessibility tree confirms the Library window, search edit with control ID 102, Layouts pane and scroll controls. |
| VERSION-01 | Installed file version metadata | think-cell 15.0.102.172. The count 497 describes this observed installation, not every edition/version. |
| REA-01 | `artifacts/thinkcell-study-20261026/intelligence-shipped-rea.json` and `rea-evidence.json` | Our shipped PowerPointIntelligence.dll is a managed x86_64 PE targeting .NET Framework 4.8. SHA256 `78801837c2de9d3836850b5097fb1e9f9817262dc01d69784c0bb619a29b2f7f`; module MVID `07b56bc8-b9c8-4185-89a5-2393834687a1`. No target code was loaded or executed by this REA inspection. |
| DOC-01 | https://www.thinkcell.com/en/resources/manual/slide-templates | Published behavior: slide-text/file-name search, category result counts, thumbnail resizing, source/destination formatting and insertion after the current slide. |
| DOC-02 | https://www.thinkcell.com/en/resources/manual/library-slide-image-assets | Published behavior: source filters, sorting, object-only insertion, and refreshing locally changed assets. |
| DOC-03 | https://www.thinkcell.com/en/resources/manual/automatic-layout | Published automatic layout resizes to content and maintains dynamic alignment relationships. Static grid placement is not equivalent. |

REA summary Evidence ID: `ev_92b4f008409316bd5636b1cc2ddcfd72112e8902cd558efa2ac09dc11b7c17bc`.
Parent managed-artifact Evidence ID: `ev_46a72dee3da206bf62d6825e6f096c10cb0d8a0cf91d9345e68e09774b2f0e44`.
The exported REA bundle's zero registered unknown-ledger records does not mean there are no unknown product behaviors. Static metadata does not prove runtime feature behavior.

The first native activation and search-control actions failed. After the owner brought the Library forward, screenshots and accessibility were captured successfully. A subsequent search action failed with an unavailable cached element, and another app obscured the window on re-observation. **Search execution, insertion, source/destination formatting and automatic layout were not runtime-tested in think-cell.** They remain documentation-backed observations only. No user presentation content was edited and no proprietary template was inserted.

## Implementation changes in PowerPoint Intelligence

| Capability | Before | Current increment | Evidence level |
|---|---|---|---|
| Template discovery | Literal substring in name/category/purpose | Whitespace-separated terms match across names, purpose, IDs and default editable field text, case-insensitively | Source + local regression |
| Counts | Page-wide count only | Query-aware counts beside all 26 categories and total | Source + partition-count regression |
| Preview size | Fixed 192 px | 128–256 px slider and Ctrl+mouse-wheel controls | Source; native interaction pending |
| Sort / metadata | Catalog order, minimal tooltip | Catalog or name sorting; tooltip with category, purpose and stable ID | Source; native interaction pending |
| Field editing | Drafts lost on navigation; manual preview refresh | Per-template in-memory drafts survive selection/filter/page/palette changes; committed edits refresh after 250 ms; 2,000-character UI limit | Pure draft/query regression; event wiring compiled |
| Grid layout | Fixed three-column coordinates, regardless of dimensions | Preflight regular-grid placement using actual object widths/heights, 16 pt spacing and slide margins; impossible fit rejected before COM edits | Bounds/overlap tests on 16:9 and 4:3; native mutation pending |

This is an **original implementation of a limited feature increment**, not a full think-cell clone or complete parity claim. Existing Executive Slate default, twenty palette choices, centered text defaults, alignment controls and optional native builds/transitions are retained. No new SceneContractTests.exe or TemplateLibraryTests.exe was created to get around their provenance gates.

Drafts are session-local memory, not saved personal assets or a filesystem index. Search covers catalog/default field text, not arbitrary PowerPoint files, network drives, OneDrive or custom draft values. Preview sizing does not alter inserted slide geometry. Grid layout tries regular grids of up to 12 columns without resizing selected shapes; it is not a dynamic constraint solver or an optimal packing algorithm. If a native grid move fails, restoration of prior positions is attempted and incomplete rollback is reported.

## Remaining gaps — not claimed implemented

- Destination slide-master font/color adaptation and source/destination formatting choices.
- Inserting only template objects onto the current slide, rather than a new slide.
- User-added folder/network libraries, persistent indexing, file/date/author/aspect/master filters and duplicate-slide recognition.
- Search-term highlighting in thumbnails.
- Dynamic edge locking, snap relationships and content-driven reflow across objects.
- Runtime validation of our new browser interactions and bounds-aware grid in installed PowerPoint.

The owner-selected library/layout scope does not close unrelated chart, data-link, S0 or antivirus/provenance issues. We have not inferred proprietary rendering algorithms or internal architecture from a screenshot.

## Tests and reproduction

`tests/library/LibraryBrowserTests.cs` adds 62 assertions for field-text/multi-term search, query normalization, category count consistency, isolated/copy-safe draft values, empty and invalid drafts, geometry bounds, non-overlap and impossible layouts. They run through the existing DemoContractTests runner via `scripts/Test-Ribbon.ps1`, without rebuilding the blocked template test executable. The Office grid scenario now checks bounds and overlap instead of mirroring the old fixed-coordinate formula.

After installing the updated build, open Library and search `accountable leader`; results should include organization layouts through their editable field text. Change a title, browse another category/page, then return; the draft should remain. Change the palette, zoom thumbnails and sort by name. Edit a field and observe a debounced preview refresh. Select heterogeneous native shapes, use Grid layout, and check margins/no overlap on 16:9 and 4:3. Oversized selections should fail before moving any shape.

### Validation outcome — 26 October 2026

`artifacts/validation-20261026-131030-566920d8/results.json` records **18 PASS / 8 FAIL** for source snapshot `E709F44F24C6929ECE8BE5BF52953C737262FA5F6AB2F851CD826FD6F50F8C93`. The ribbon regression check includes all 62 new assertions. AnyCPU, x86 and x64 compilation passed, as did native renderer core, domain, schema, calculations, compatibility and cross-language checks.

Failures are retained as failures: scene-contracts AnyCPU/x64 require reviewed provenance; template-library and template-library-office stop on unreceipted existing runners. Installed-ribbon, demo-feature-runtime, native-shapes-office and office-1 stop at PowerPoint COM activation with `0x80070005`. No security settings or runner receipts were changed to suppress these results. Further repeated Office runs were not attempted after the activation failure.

The source ribbon marker is `PPI-DEMO-RIBBON-20261026-v10`; the previously installed package is not evidence that these new interactions work. Desktop installation and native interaction/save/reopen validation remain required. **Acceptance: partial source increment validated locally; installed Office acceptance incomplete.**

## Library completion increment — 26 October 2026

The subsequent owner request authorizes source fixes and implementation of the remaining library/layout controls. The following original controls are implemented; native Office claims remain gated on the tests below.

- **Objects on current slide:** inserts a uniquely tagged editable native group without creating a slide, replacing existing objects, changing the existing slide background or setting its transition. Color/alignment operations can target the selected template instance. Per-stage builds are explicitly rejected in this mode; use new-slide insertion for staged builds. Cleanup only removes this insertion's objects.
- **Destination theme:** new-slide insertion uses the active slide's design and native theme color/font references. The selected palette remains the default. The preview clearly states it shows palette colors; destination appearance is applied by PowerPoint. This is theme adaptation, not a guarantee that arbitrary master placeholders or layouts are reconstructed. Microsoft documents [theme color roles](https://learn.microsoft.com/en-us/office/vba/api/office.msothemecolorindex) and [major/minor font roles](https://learn.microsoft.com/en-us/office/vba/api/office.themefontscheme).
- **My PowerPoint files:** choose a local/network folder of owned `.pptx`/`.potx` files. The bounded OOXML index follows presentation relationship order, reads slide text and author/layout/master metadata without opening a source presentation, and filters by words, author, layout/master, aspect ratio and modified date. Sort by file/slide, title or newest. Refresh rebuilds the index. Macro-enabled formats are excluded. Native slide insertion uses [Slides.InsertFromFile](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slides.insertfromfile); destination-design adaptation is optional. Source changes are checked before insertion. Exact duplicate recognition is conservative and includes package identity; semantically equivalent slides in different packages are not necessarily recognized.
- **Optional persistent index:** the explicit Remember checkbox stores searchable slide text and metadata on this computer under `%LOCALAPPDATA%\PowerPointIntelligence\library-index-v1.xml`. It is off for new users. Unchecking deletes the index. No index or presentation content is uploaded or written to diagnostic logs. The versioned XML cache rejects unknown versions, DTDs, escaped folder paths and oversized content. Cached results require Refresh after source changes.
- **Search highlights:** matching text fields in procedural template previews receive a highlight. These are preview annotations, not shapes added to the resulting slide. The local-file browser exposes searchable text/metadata rather than rendered native slide thumbnails.
- **Layout Links ribbon command:** opt-in, one-hop native alignment, horizontal/vertical gap preservation, slide-edge locks and shape-to-text fit. Tags persist the relationships. A UI-thread timer updates only an explicitly enabled active Normal-view slide; it stops on an invalid/missing/ambiguous anchor or COM failure, with status in the task pane. Opening Layout Links resumes it. Relationships are also evaluated before saving the active presentation. Cycles, chains and out-of-slide results are rejected. Clear links releases selected relationships and dependents of selected anchors. No dynamic layout is claimed while the add-in is unloaded, in a slide show, or for arbitrary unlinked objects; this is not a general layout solver.

The source ribbon marker is now `PPI-DEMO-RIBBON-20261026-v11`. The prior Remaining gaps section describes the earlier increment; the capabilities above supersede those entries only to the specific extent described. Native slide thumbnail rendering for user files, arbitrary master-placeholder mapping, multi-hop layout solving, semantic duplicate matching and installed runtime acceptance remain open.

### Failure remediation evidence

- `artifacts/activation-agent-current/activation.json`: x64 Windows process, AppContainer=1, restricted token=1; registered PowerPoint class and LocalServer32 found in both registry views; COM instance creation returns `0x80070005`. This establishes the failing stage and execution context, not which specific Windows policy/ACL denied it. A normal desktop-token comparison remains necessary. Registry, DCOM, trust and security settings were not weakened.
- `artifacts/scanner-agent-current/scene-threat-evidence.json`: eight historical McAfee quarantine records for scene-contracts-x64 paths, with exact hashes/threat labels and no collector read errors. These are not false-positive clearance records. The owner temporarily disabled protection, then confirmed it active again; later tests run after that confirmation. No protection settings were changed by the agent.
- Added an explicit reviewed stable-path runner rebuild (`-RebuildReviewedRunner -ScannerEvidence <report>`). It records prior hashes/receipt and scanner-report hash; the successful actual compilation creates the new receipt. It refuses missing files, matching detected paths/hashes and incomplete diagnostic reads. It never changes antivirus settings or copies/renames executable files. Default reuse still verifies build inputs and file hashes. Five new inert-fixture assertions cover review behavior and scanner refusal.
- AnyCPU scene and template runner receipt failures were repaired by actual source compilation into their existing paths. x64 scene remains held because McAfee's records identify that path; a source rebuild or exclusion would not itself establish vendor clearance.

### Acceptance scenarios added

The pure regression runner now covers 96 assertions, including theme roles, local OOXML ordering/filters/date/aspect/duplicates, index round-trip/version/path validation and layout-rule wire/geometry failures. The existing Office template test adds object-only insertion/background isolation, selected-instance recoloring, destination theme roles, anchor-resize reflow and object/link save/reopen assertions. These Office assertions must execute; compiling them is not a pass.

### Desktop activation comparison received — 26 October 2026, 13:55 Dubai

The owner ran the same diagnostic from normal x64 Windows PowerShell and saved `artifacts/activation-desktop-current/activation.json`. It reports **IsAppContainer=0, IsElevated=0, Activation=PASS, Application.Version=16.0**, with the same CLSID and LocalServer32 registration as the agent diagnostic. Desktop activation is therefore verified without elevation, registry edits or weakened DCOM/security settings. The agent's failure is specific to its execution context; the exact enforcing ACL/policy is not identified. Do not report this agent denial as a broken desktop Office installation. The comparison does not verify the installed version, ribbon interaction, layout callbacks, or save/reopen acceptance.

The reviewed template runner was rebuilt and executed again after the owner confirmed McAfee protection active. It passed 520 procedural layouts and 10,400 palette/layout combinations. Pre-existing exclusions were not inspected or changed, so this result is not evidence that McAfee scans and clears every executable. x64 scene scanner clearance remains unresolved.

### Final local source validation

`artifacts/validation-20261026-135649-7c072e46/results.json`, source snapshot `B28F6656EBFF2D8CDC8525B21B41296537E49D0EAA94AC5FE76F8078E5DABE70`: **20 PASS / 1 FAIL** across all 21 local checks. AnyCPU/x86/x64 compilation passes; native renderer core 41/41, ribbon preparation 9/9, library/layout regression 96 assertions, cache policy 15 assertions, scene AnyCPU/x86 148 assertions and 73 round trips each, and 520-template/10,400-palette-combination tests pass. The x64 scene check fails before execution on its legacy provenance receipt. A compile-only reviewed rebuild request separately refuses the detected path, recorded in `artifacts/scene-x64-review-current.log`. No x64 scene execution or scanner-clearance pass is claimed.

The earlier 13:50 integration run caught a timer callback parameter naming conflict in all three VSTO compile targets. That source error was fixed; the final local snapshot above compiles all three successfully. Earlier agent Office failures remain context-bound diagnostic results. Desktop activation PASS is independently established by the owner's 13:55 report; installed-v11 integration, object editability, layout links and save/reopen still require the requested normal-desktop install/test run. Feature acceptance is not complete.

Imported slides also receive `PPI_LIBRARY_ASSET_SCHEMA=1`, `PPI_LIBRARY_ASSET_VERSION=<indexed signature>` and `PPI_LIBRARY_SOURCE_SLIDE=<source position>` tags. These persist provenance without embedding the source file path. Their save/reopen persistence remains a native acceptance assertion to execute.

Read-only file-header/version evidence at `artifacts/activation-desktop-current/powerpoint-file-metadata.json` identifies the registered POWERPNT.EXE as machine `0x8664` (x64), file/product build `16.0.20527.20088`. This does not infer the licensed Office edition from the shared version 16.0 number. Combined with the owner's desktop activation PASS, it establishes the target executable architecture; the x64 scene runner remains unexecuted.

### Desktop validation stall — investigation in progress

The owner's desktop run `artifacts/validation-20261026-140258-f9755fa9` passed S0 unit tests and then stopped producing output during `Test-SceneRunnerCache.ps1`, before ribbon compilation or Office activation. Its last recorded output is the original ten inert cache-policy assertions. The exact blocked operation is not yet established; no antivirus or COM cause is inferred for this stall.

The suite now uses `scripts/Check-Process.ps1`: file-backed stdout/stderr, bounded waits, explicit failed timeout results, and stopping later checks after a timeout. It stops only the owned check process; it does not terminate user Office applications. Stream files remain available as evidence. Default local timeout is 300 seconds; Office timeout is 3600 seconds, configurable at invocation. New process-harness tests verify simultaneous large stdout/stderr output, nonzero exit propagation, embedded-quote/space handling and an actual one-second timeout against a synthetic sleep process. These passed locally; Windows PowerShell desktop compatibility still requires execution there.

Cache review now emits `REVIEW` stage markers before reading scanner JSON, hashing each inert fixture, and writing the audit. The owner has been asked to stop the stuck run and execute the isolated policy script to identify the last completed stage. Do not mark the desktop stall resolved merely because a timeout was added.

Deployment 1.0.0.15 successfully installed v11 at 14:02 Dubai. Its source receipt precedes the final imported-asset version-tag edit, so its hash does not match the final source. A final rebuild/install remains needed. The agent acknowledged requesting installation before that last source edit was finished.

### Latest harness-inclusive local run

`artifacts/validation-20261026-141325-e867efc9/results.json`, snapshot `E81E4B9BFA8F06706A4DA66FD0433CD37C0B7B4490F8E5EA6D07D9C4AB3A20E1`: **20 PASS / 1 FAIL**, with no timeout. This run includes final imported-asset tags and the bounded process harness. All three VSTO compile targets pass. The sole local failure remains the held x64 scene legacy receipt. The desktop stall remains under investigation awaiting the isolated REVIEW-stage output; it is not marked fixed by this agent-context run. The new normal-desktop install/test request remains pending.

### Cache audit serialization fix — 26 October 2026

The owner's isolated Windows PowerShell run reached all three `REVIEW hash` stages and then stalled after the original `REVIEW write audit receipt` marker. No audit file was present in that fixture directory. Inspection confirmed that `Get-Content -Raw` on its prior receipt returns a string decorated with `PSPath`, `PSParentPath`, `PSChildName`, `PSDrive`, `PSProvider` and `ReadCount`; `[IO.File]::ReadAllText` returns plain string content without those properties. Microsoft documents that ConvertTo-Json stopped serializing extended string properties only in PowerShell 7.2: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/convertto-json . This supplies a concrete Windows PowerShell/PowerShell 7 behavioral difference consistent with the desktop-only stall; the original marker covered several operations, so desktop verification remains necessary before claiming the exact stall resolved.

The review audit now reads prior text through ReadAllText, applies a 1 MB prior-receipt limit, constructs explicitly flat audit fields, serializes separately at depth 5, and writes UTF-8 through WriteAllText. Additional stage markers distinguish prior-text read, serialization, write and completion. Scanner/path/hash gates and the original receipt remain unchanged. No antivirus or PowerPoint settings were changed.

Executed locally: `artifacts/cache-policy-plain-audit.log` contains **19 passing policy assertions**, including exact plain prior text, bounded/provider-free JSON, original file hashes and no false execution claim. `artifacts/ribbon-plain-audit.log` passes the full compiled ribbon regression. These targeted checks cover this script-only change; there was no need to rebuild add-in binaries. The owner has been asked to confirm audit completion and all 19 assertions in normal Windows PowerShell. The prior final local suite's x64 scene gate and installed Office acceptance remain open.

### Desktop cache-stall resolution confirmed

The owner reran `scripts/Test-SceneRunnerCache.ps1` in normal Windows PowerShell after the plain-text serialization change. Their complete output reaches `REVIEW read plain prior receipt`, `REVIEW serialize plain audit fields`, `REVIEW write audit receipt`, `REVIEW audit complete`, then **19 total cache policy assertions passed**. Fixture directory: `C:\Users\srini\AppData\Local\Temp\ppi-cache-policy-38a31bf1b7404452a288fd61c7af2060`. This confirms the reported desktop audit stall is resolved by the changed serialization path. It does not establish x64 scene runner clearance or installed Office integration acceptance.

## Original-source library mode — owner correction on 26 October 2026

The owner explicitly rejected the 520 procedural layouts as a substitute and requested unchanged think-cell designs. That catalog is not a reconstruction of the vendor collection. Read-only discovery located the actual `.potx` files at `C:\Program Files (x86)\think-cell\templates`. The owner subsequently approved temporary hidden, read-only template access for native previews and source formatting; chart generation still targets the active presentation.

### Evidence

- Independent ZIP/XML inventory: `artifacts/thinkcell-study-20261026/original-template-inventory.json`, **62 source files, 497 slides**, per-file SHA256 and structural counts. No source slide XML or assets were copied into product sources.
- Product index execution against those files: `artifacts/thinkcell-study-20261026/original-index-validation.json`, **62 PASS, 497 indexed slides, zero failures**. Maps and hand-drawn callout slides required a bounded 40 MB streaming slide reader; metadata XML remains capped at 2 MB. Package/file/text/count limits and disabled DTD/entity resolution remain enforced.
- REA opened `Diagrams\Cubes\Cubes.potx` as a ZIP archive, SHA256 `e21afb8cca7cd82eb39af4079fcdc96d953316452ed6d7d38c57e8efc1a9eeb7`, 62 central-directory file occurrences. Summary evidence `ev_3373c5860f06b87cc8bcfaf78cdd23a5ba966cd8dd598b8273e8bc353434d1f6`, inspection parent `ev_e8b2183d262b7426352ede5ce0620cda5d3c69d5b31f9cfc8b405fd41df240b9`. This is container metadata, not binary decompilation or proof of native editability. The session was closed. The later attempted export contained zero records after close and is not used as evidence.

### Implemented source behavior (v12)

Library opens the installed original-file collection when available; procedural layouts are a fallback. Recursive discovery preserves the folder taxonomy and slide counts. Native previews use PowerPoint's slide export from an owned hidden read-only source. The browser retains one selected source session, closes it on collection changes or form close, and refuses to take ownership of an already-open source presentation. Closing the library is cancelled with a status message if its owned source cannot be closed.

Original insertion validates identical slide dimensions, checks source identity for changes, uses native slide import, clones the source design, restores its custom-layout association and master-background setting, and records source version tags. It does not apply our palette, text alignment, animations or transitions. It inserts into the captured destination presentation, never into the hidden source. On failure it attempts to delete only its inserted slide and cloned design. Exact rendered fidelity remains unverified until the native assertions execute; installed fonts and think-cell-specific behaviors may affect it.

This is **use of the locally installed original template files**, not an independently recreated or redistributed think-cell library or editing engine. The external collection must remain available on the machine. think-cell-specific chart editing behavior is not reimplemented by importing a slide; no standalone parity claim is made.

### Validation and outstanding gates

The existing contract runner passes 102 library/layout assertions after six new recursive-discovery/root-boundary/size-preservation cases, plus ribbon/icon regressions. Added a focused `--source-originals` native scenario to the existing stable TemplateLibraryTests.exe: Cubes and Pyramids source/inserted/reopened pixel equality, native object/build/transition counts, editable text persistence, source provenance and unchanged original-file hashes. Interactive template validation selects this scenario when the original collection exists; the old 520-procedural native scenario is available explicitly through `-ProceduralOffice`.

The prior desktop run `validation-20261026-142809-bc7afb64` exposed a process-handle/exit-code problem: large stdout/stderr completed but normal Windows PowerShell returned null exit codes for completed checks. The harness now retains the process handle before waiting, does not discard it through Refresh, and refuses to infer success from a missing code. Its three local regressions pass; desktop confirmation is pending. That run's procedural Office test timed out after 3600 seconds, leaving a descendant that locks the stable TemplateLibraryTests.exe. The agent did not create another executable to avoid that lock. The owner was asked to stop only that runner and close its test deck normally. Native original fidelity compilation/execution resumes after the stable file is released.

### Source-mode completion checks and stable-runner recovery

After the owner stopped the prior runner, its stable executable was absent. The preceding failed compiler attempt had reported CS2012 (file in use); the disappearance's exact cause was not established. Expanded scanner collection at `artifacts/scanner-originals-current/scene-threat-evidence.json` found six historical TemplateLibraryTests detections in old 23 October snapshot paths, as well as the eight scene detections. None of those template records matches the stable root path or its most recent receipted hash. This corrects the earlier limited collector's statement that no template detections had been found; the original collector searched scene names only. No 26 October quarantine cause is inferred.

Added explicit missing-file recovery for the reviewed template build. It requires a prior build key and this file's valid SHA256, an error-free scanner report, and refuses any detected prior hash or exact path. Default review still refuses missing files; recovery review creates only an audit, never restores a binary. The actual compiler then rebuilt **the same TemplateLibraryTests.exe path** from current sources. No alternative executable or antivirus exclusion was created. Four new inert recovery assertions bring cache policy coverage to 23. `artifacts/original-fidelity-compile.log` records the successful same-path compilation; `artifacts/originals-static-final.log` records successful static template validation afterward. Original native fidelity assertions compile but remain unexecuted until a normal desktop run.

The product compiles for AnyCPU, x86 and x64. The original-source index runs successfully across all installed files, and pure/ribbon checks pass with 102 library/layout assertions. The source-file browser defaults to installed originals, preserves original formatting controls, and does not write into the source template files. Source-exact output acceptance is pending native pixel/save-reopen checks, rather than inferred from source availability or compilation.

### v12 final local source snapshot

`artifacts/validation-20261026-155254-b81f1beb/results.json`, source SHA256 `8C8458044A1474D384D3C4F4319CD935F1426C37BBCE2873C80CA1DA07DA7EB5`: **20 PASS / 1 FAIL**, all AnyCPU/x86/x64 compile checks pass. The sole failure remains the x64 scene runner's legacy receipt/scanner-review gate, unrelated to original-template discovery. The suite includes 102 library/layout assertions, 23 cache-policy assertions, three process-harness tests, 41 renderer core tests and the static procedural-catalog regressions. Native original fidelity is compiled, not executed in this agent run.

The owner was asked to verify the process-handle fix with Test-CheckProcess.ps1 in normal PowerShell, install final v12 source, then run Test-TemplateLibrary.ps1 -InteractiveOffice. This selects the actual original-source pixel/edit/save-reopen scenario on this installation. The previously installed v11 package does not contain original-source mode. Source updates are frozen while this desktop validation is requested.

## Desktop run 26 October 2026, 15:59 — failures isolated

`artifacts/validation-20261026-155936-44d22bed/results.json`: **25 PASS / 3 FAIL**. Installed-ribbon identity/hash/OnLoad, native-shape Office persistence, and all three S0 adapter integration runs PASS. These are real desktop x64 Office results and supersede the agent-context activation denial for that execution context. The process-handle fix also propagates exit codes correctly in this run. Full process restart/task-pane visual acceptance is not implied by adapter assertions.

### Template fidelity false failure

The first Cubes slide reached native source export and insertion, then the original comparator failed. Independent Pillow comparison of its saved source and inserted PNGs found **zero different RGBA pixels**, identical 720×405 dimensions, maximum difference (0,0,0,0). Reproduction showed source-preview PNG DPI about 95.99 with 32-bit storage, versus imported PNG about 53.98 with 24-bit storage. Drawing those bitmaps through Graphics.DrawImageUnscaled before hashing produced different buffers even though decoded image pixels match. This is a comparator failure, not evidence of a visually changed Cube.

Pixels() now hashes exact decoded RGBA values and dimensions directly, without a Graphics drawing operation. No tolerance or acceptance threshold was relaxed. Two new regressions cover differing DPI/storage with identical colors, and detecting a one-pixel change. The revised comparator also yields identical SHA256 values on the captured native pair; see `artifacts/thinkcell-study-20261026/cube-pixel-failure-analysis.json`. The other fifteen original slides and save/reopen were not reached and remain unverified.

### Demo report workbook-close disconnection

The sole executed demo command, demo-tour, failed at DataSources.cs workbook.Close(true) with 0x80010108. The CSV shows **1 FAIL / 89 BLOCKED**; those 89 commands were not executed. An early progress message incorrectly called them passes and was explicitly corrected after inspecting the CSV. The log establishes the failing call, not whether Excel's close had completed before its proxy disconnected.

The proposed fix never replays a disconnected Close. It accepts that specific close result only if the full write body completed and independent reads of the live native chart show expected series/point counts and populated numeric values; verification does not call Refresh in that ambiguous-close path. An uncommitted chart, unavailable PowerPoint server, or unrelated COM error still fails. Four pure fault-injection cases pass for normal close, verified disconnected close, failed verification and unrelated server loss. The linked-report integration assertion now checks chart values across all seven generated slides before save and again after reopen. Desktop execution of this proposed fix remains pending.

### x64 scene runner hold is byte-specific

Current SceneContractTests.exe SHA256 is `EA5C34B16761FD30AC9ADF0F8CB4DE362B1AA81A93DD9D170D06EA0A087FA7A3`, exactly matching McAfee's recorded `ti!EA5C34B16761` / infection quarantined / 22 October 12:03:34 UTC. It also differs from its receipt's `E02CDBF2FEFB13AD09CB45D78750391FDC49A38AE46541A29D94A4E6BD96B17E`. Do not treat this as only a historical path warning. The existing quarantine/hash/provenance guard remains intact; no execution, replacement, renaming or security change was used to turn it green.

### v13 source validation

`artifacts/validation-20261026-160844-c4733ca0/results.json`, snapshot `95AAC0C666FFF5A48AE91AE736AB5AAF27F0F38C25D84444BB15817DE45D456D`: **20 PASS / 1 FAIL** locally. All three architecture builds, ribbon/contracts, 104 library/layout assertions, renderer and domain checks pass; x64 scene remains held. Native original assertions compile into the same stable template executable. The product marker advances to v13 for the workbook-close fix; installation and desktop revalidation are required before accepting that change.

## Thumbnail gallery request — v15

The owner supplied a screenshot of the original-source browser and asked to replace the long text list with sample images. `apps/powerpoint-vsto/LocalLibraryGallery.cs` now provides a native large-icon thumbnail grid with short collection/slide captions, folder-tree navigation, existing search/filter controls, a selected-slide preview, optional details panel, Small/Medium/Large thumbnail sizes and double-click/Enter insertion. It renders one page of 24 slides through the approved source-preview provider on the Office UI thread, and retains at most 96 thumbnail bitmaps in memory. Loading and unavailable-preview cards are explicitly labelled. No sample slide/image is substituted for a missing source render.

Ten new gallery tests cover page size, source identity/captions, image-list binding, actual queued image-slot population with an injected test renderer, next/previous boundaries and cache reuse. They exposed an initial native-handle creation issue; the control now creates its native list handle before populating cards so parent/tree selection events cannot invalidate the initial items. Tests pass. The gallery is split into its own product source file; all production/manual compile source lists and stable-runner cache inputs include it.

`artifacts/thinkcell-study-20261026/gallery-ui-preview.png` is an actual WinForms layout capture using seven previously captured native PowerPoint source PNGs through an injected preview provider. It demonstrates the grid layout, not a deployed PowerPoint session or all-497 preview execution. The library's normal runtime provider remains PowerPoint export from owned hidden read-only sources. Native deployment/interaction validation is pending installation.

### Continuing investigation of the preceding desktop run

`validation-20261026-161730-bc82f220` passed exact native insertion pixels/object counts/builds/transitions/hidden-source cleanup for all three Cubes and the first three Pyramids. The next Pyramid (source slide 4) differs in eight grayscale edge pixels within (668,154)–(675,158), max channel delta 10. That exact-pixel failure remains a failure. The product now retains the validated preview source instance instead of reopening it unconditionally before import; file-stamp changes trigger reload and the full source signature is still checked. This is a proposed stabilization, not a verified diagnosis of those eight pixels.

The original native test now collects insertion and reopen fidelity across the full set rather than stopping at the first pixel mismatch. It exports repeated-source PNGs, source/imported geometry/text hashes, a per-stage `original-fidelity.csv`, and the native fixture deck. Pixel/geometry mismatches still fail final acceptance; no tolerance was introduced. These diagnostics distinguish source-render variation from an import change on the next desktop run.

The demo CSV for that desktop run contains 1 PASS, 2 FAIL and 87 BLOCKED. The new error 0x800706B5 occurred during DLR GetITypeInfo at workbook close. Close now uses explicit reflection dispatch with optional parameters rather than dynamic metadata scanning; Microsoft documents IDispatch support in Type.InvokeMember (https://learn.microsoft.com/en-us/dotnet/api/system.type.invokemember). Wrapped COM errors retain their original HRESULT/stack. A disconnected or unknown workbook interface is accepted only after the completed write's live native chart view is refreshed and expected values/counts verified; Close/Save are not replayed. Unverified results and server loss still fail. Three additional local cases cover unknown-interface verification and reflection close binding. These production fixes remain pending desktop execution.

### Final v15 local validation

`artifacts/validation-20261026-165627-20879fae/results.json`, snapshot `C9C7F7534EAA1F4756D4FDB967B7FB8594065BBE253E30E1D2121A403150AABE`: **20 PASS / 1 FAIL**. AnyCPU/x86/x64 compilation, template/ribbon contracts, ten gallery tests, 104 library/layout assertions, 23 cache-policy assertions, renderer/domain/schema checks pass. The x64 scene runner remains held for the previously documented exact quarantine-hash and receipt mismatch. The stable template runner was rebuilt at the same path; no alternate executable was created. v15 is ready for normal-desktop installation; earlier v13 native failures are not marked resolved by local checks.
