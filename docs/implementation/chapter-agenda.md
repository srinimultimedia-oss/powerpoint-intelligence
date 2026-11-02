# Chapter-driven agenda repair

Build marker: PPI-DEMO-RIBBON-20261027-v17-agenda1. Status: implementation and local checks complete; native acceptance pending.

## Investigation

The REA-guided public-behavior and Windows UI study is in [agenda-reference-study.md](agenda-reference-study.md). Native reference observation confirmed the Agenda menu and chapter insertion in an isolated unsaved Presentation2. Rename/reorder/hierarchy/ToC behavior was documentation-backed rather than observed. REA native-window tools in this session target Mach-O; Windows observation used the computer-use skill. No vendor binary, protected internals, template or artwork was copied into product code. The pre-existing Presentation1 was not edited.

The principal defect was architectural: our agenda was a one-time list of slide titles or PowerPoint sections, with two additional unmanaged aliases. It was not a chapter-driven agenda. It also overwrote custom overview headings on refresh.

## Implemented behavior

Structure > Agenda now exposes Chapter, Table of contents, Manage Agenda, Split agenda here and Update agenda. Both legacy catalog aliases route through the managed chapter/contents workflow; they no longer add unmanaged text snapshots. The former current-slide agenda alias is now explicitly labelled a table of contents and inserts its own managed slide.

- Chapter inserts a native editable agenda slide after the active slide and selects its heading. Native slide IDs are canonical identities. It does not silently treat every slide title or PowerPoint section as a chapter.
- Table of contents inserts before the first chapter in the current scope, or after the active slide when no chapter exists. An empty contents slide is valid and populates after chapter creation. Top-level chapters are shown without a current-chapter highlight.
- Manage Agenda renames chapters, changes levels 1–4, optionally shows all levels, moves chapter blocks up/down and selects the native layout frame for positioning/resizing. It refuses application to a different presentation than the one that opened the dialog.
- Numbering is hierarchical chapter numbering, not slide page numbering. Invalid hierarchy jumps and overflow are rejected before mutation. Default chapter views show roots, the current path, siblings and immediate children; scopes and Show all control visibility.
- Editing a visible chapter caption on any managed view propagates its plain text to related views. Distinct concurrent captions for the same chapter are treated as a conflict, preserving edits and stopping synchronization rather than silently selecting a winner.
- A debounced selection-change update runs after leaving text editing, with a save-time check and an explicit Update command. The unchanged-state cache prevents repeated updates of an unchanged agenda. This is not an interception of every native keystroke/drag action.
- Chapter moves use the manager and include descendants and intervening content as a contiguous native SlideRange. Moves do not cross splitter boundaries. Native Slide Sorter chapter-only drags are not intercepted to move content automatically.
- Split agenda here marks a boundary on the selected non-agenda divider slide. Managed chapter/contents slides are refused before mutation, preventing contents-placement and chapter-movement ambiguity. Subsequent chapters belong to an independent scope. There is no dedicated merge/remove-split UI yet.
- Copied chapter slides receive their actual new anchor; their captions can change independently of the source. Versioned metadata survives via a bounded custom XML registry. Deleted chapter slides are removed from the registry on update; content slides are not automatically deleted.
- New views inherit a preceding agenda role model, falling back to the first existing agenda. Native master/layout/background, unrelated objects, custom heading and available text formatting are retained. Existing headings are not reset on refresh. Default headings are adjusted when crossing between contents/chapter kinds.
- Empty parents immediately followed by a subchapter are automatically hidden; parents hidden by this rule are unhidden when content is inserted. A manually hidden chapter is otherwise retained. Runtime behavior requires native validation.
- Native hyperlinks target stable slide IDs with refreshed index hints. Slideshow click-through playback has not been executed here.

Version-1 section overviews remain readable and explicitly refreshable. They are not silently migrated to chapters. Their custom title/subtitle text, positions and formatting are now preserved by refresh.

## Metadata and limits

ChapterAgendaModel.cs provides typed chapter/view/planning data, conflict resolution, scope-protected block movement and strict urn:powerpoint-intelligence:agenda:v2 XML. Unknown schema/fields, DTDs/entities, invalid dimensions/identities and excessive input are rejected.

The native adapter bounds scans to 3000 slides, 40 managed views, 4000 total field operations and 450 retained fields per view. The pure model bounds chapters to 200 and levels to four. Retired fields are hidden, not deleted, and count toward the history limit. Grouped/damaged/ambiguous managed fields stop updates. No unrelated object is automatically removed. Partial COM failure is surfaced with Undo guidance; property restoration does not establish a general rich-text transaction guarantee.

Remaining gaps include full shared style classes, character-level rich-text synchronization, drag/keyboard chapter editing parity, exact automatic resizing/snap semantics, chapter deletion UI, hidden-parent removal/reinsertion, page-number controls and full runtime comparison. Native font/header preservation and failure behavior need execution evidence; this is not complete reference-product parity.

## Executed checks

Final full local run: artifacts/validation-20261027-124428-83288e2b/results.json, source snapshot 9F9AD3581C99C08390E5577114DB68281F8796A9E96657834FAA1213C4FEA390. Result: 20 PASS / 1 FAIL. The only FAIL is the existing scene-contracts-x64 provenance/quarantine hold; no runner guard or antivirus setting was changed.

The run includes 65 new chapter planner/codec/movement assertions, 654 existing overview assertions, all 90 legacy tool bindings, 29 workflow bindings, all ten collapsed-group icons, 154 control images and all AnyCPU/x86/x64 compile targets. The final production and test/script source match that full-suite snapshot, including the focused agenda CLI and the divider-only splitter guard.

Final forced compilation independently checks the actual marker in all three intermediate DLLs. Logs: artifacts/agenda1-final-compile-*.log. Focused runner compile: artifacts/agenda1-focused-runner-compile.log. No new standalone EXE was created; tests reuse DemoContractTests.exe and DemoOfficeTests.exe. git diff --check passed.

ChapterAgendaOfficeTests.cs compiles native create/edit/conflict/duplicate/hierarchy/block-move/split/hyperlink/format/auto-hide/save-reopen assertions into the existing runner. Compilation is not execution; none of these new native assertions is marked passed here.

## Desktop acceptance

The agent's confirmed AppContainer COM boundary requires normal desktop execution. Save/close personal presentations first. The reference study used a separate unsaved Presentation2; no reference study was saved as a product asset.

```powershell
.\scripts\Install-Local.ps1
.\scripts\Test-DemoFeatures.ps1 -InteractiveOffice -AgendaOnly
```

Keep antivirus active and complete the visible publisher prompt. Send the reported artifacts folder. Also verify the Agenda dropdown and inline heading editing visually, and follow hyperlinks in Slide Show. Full-suite Office checks remain available through Test-All.ps1 -InteractiveOffice after the focused scenarios pass. The source is frozen while desktop validation is pending.
