# Agenda reference study

Study date: 2026-10-27. Independent public-behavior analysis under REA skill; AGENTS.md prohibits vendor binary/protected implementation inspection. The initial documentation review used our source and official documentation. The parent subsequently performed the bounded Windows UI observation below. No runtime parity claim is made.

## Public reference behavior

The [official agenda manual](https://www.think-cell.com/en/resources/manual/agenda) documents:

- Chapter insertion creates a slide after the current slide and selects its editable name. Reordering a chapter moves its content slides; rename, addition, removal and reorder synchronize agendas.
- Text levels establish hierarchy. Subchapters appear within their documented parent/same-level scope. Empty parent chapter slides automatically hide and unhide as content changes.
- Splitters isolate independent agenda regions.
- A table of contents precedes the first chapter, lists top-level chapters without highlighting, and may initially be empty.
- Heading appearance propagates by semantic type, including current chapter, ancestors and other hierarchy relationships. Partial text formatting propagates that chapter's text without restyling other chapters.
- Navigation uses double-click in editing and single-click in slide show.
- New agenda slides inherit shapes, master and layout from the nearest preceding agenda, using the first agenda initially.
- Position inheritance and optional content-driven resizing are documented.

This synopsis deliberately stays below the source's 200-word reuse allowance. Direct page opening redirected to an inaccessible domain; the complete official page was available through the search tool on the study date. No protected training PDF or vendor assets were used.

## Current implementation findings

Source baseline: PresentationAgenda.cs and DemoTools.cs as read before the parent's new agenda increment. These observations may become historical as fixes land.

| Priority | Source evidence | Actual implementation problem / proposed acceptance |
|---|---|---|
| P0 | AgendaPlanner.Build; Capture | Canonical inputs are SectionProperties or every visible title, not explicit managed chapter identities. Add versioned independent chapter records and hierarchy; keep section/title import an explicitly named optional conversion. |
| P0 | Insert | Creates a blank managed overview after the active slide, using that active content slide as context; has no chapter-name input. Separate chapter creation from overview insertion and validate empty initial overview. |
| P0 | RefreshSelected; Capture | Updates only selected agenda, and replaces chapter captions from current section/title sources. Native edits of generated captions are not canonical. Add coherent multi-agenda refresh with stable chapter identity; test rename propagation after save/reopen. |
| P0 | DemoTools.Build entries agenda / agenda-layout; Agenda / AgendaCurrent | Two legacy aliases still produce unmanaged text snapshots. Route every exposed agenda command through one managed workflow; avoid duplicate legacy behavior. |
| P1 | Upsert | Refresh hardcodes Agenda title/subtitle, resets bounds and forces whole-field bold. Edited heading text, paragraph placement and mixed emphasis can be overwritten. Preserve intentional edits; test distinct per-role styling and geometry across repeated refresh. |
| P1 | Insert uses active.Design and blank layout | Does not retain predecessor agenda shapes/layout or title edits. Add deterministic role-model selection and copy tests without cloning arbitrary owner identifiers. |
| P1 | AgendaSlideInfo / AgendaEntry | No level, parent, region, overview/chapter kind or style-class fields exist. Flat boolean Highlight cannot model a hierarchy. Extend typed plan without silently interpreting PowerPoint sections as chapter ancestry. |
| P1 | Build filters Hidden and IsAgenda | Hidden content is omitted as an input policy; no parent-without-content state exists. Do not confuse user-hidden content with managed chapter visibility. Explicitly test preservation of user intent. |
| P1 | Apply | Obsolete fields are hidden, not removed, and count toward a 256-field history bound. Repeated structural edits can permanently refuse refresh. Define a user-visible retained-history policy and test no unrelated-object deletion. |
| P1 | Existing item keys item:targetSlideId | A section's identity changes when its first visible content slide changes. Preserve chapter IDs independently of first target slide. |
| P2 | SubAddress and Upsert hyperlink | Links target first visible section/content slide, not a dedicated managed chapter slide. Verify navigation policy, editing view and slide-show behavior explicitly. |

## Evidence limits and open questions

- The retrieved agenda page does not establish a dedicated automatic page-number column, numbering options or exact chapter-deletion dialog/shortcut. Our current 01/02 prefix is an ordinal, not a slide page number. Do not present it as validated reference behavior. Native observation or another official source is required.
- General automatic synchronization is documented; event timing, transient dialogs, conflict behavior, Undo grouping and save-time recovery are not specified here. Independently designed event handling needs adversarial tests.
- No claim is made about copying internal vendor metadata or rendering. Source formatting preservation must use independently implemented native PowerPoint operations and user-owned records.
- Hidden-ancestor removal/reinsertion, exact scope of nested highlighting and splitter movement need dedicated native comparison cases before parity acceptance. No implementation should delete content slides automatically merely because its chapter was removed.
- Native end-to-end tests must exercise create, rename, reorder, hierarchy changes, splitter boundaries, multiple overviews, style propagation, user shapes, navigation, duplicate IDs and save/reopen.

## Existing tests and missing acceptance evidence

PresentationAgendaTests.cs exercises flat section/title selection, simple bold highlight, stable target IDs/index hints, layout bounds and invalid-input rejection. WorkflowOfficeTests.cs calls Insert and RefreshSelected, checks actual section captions and hyperlinks, renames a section, preserves an unrelated note and permits native title editing. Those assertions support the existing overview feature only; they do not exercise managed chapters, hierarchy, shared type styles, synchronized independent views, splitters, role-model inheritance or content movement. This documentation task did not execute those tests.

The user's requested agenda repair should be accepted against named behavior scenarios, not a ribbon label or source compilation alone.


## Parent Windows runtime observations

REA native-window observation tools advertised in this session target Mach-O applications, so Windows observation used the computer-use skill and its Sky API. Vendor binaries were not decompiled/instrumented. The owner presentation (Presentation1) was left unchanged; a separate unsaved Presentation2 was created for the study.

- UI-AGENDA-01: The installed Agenda dropdown visibly contained Chapter, a disabled Agenda Placeholder, Table of Contents and Split Agenda.
- UI-AGENDA-02: Choosing Chapter inserted a new Agenda slide after the original blank slide. Its Chapter heading was selected for editing. This directly corroborates chapter insertion, not the rest of the published workflow.
- UI-AGENDA-03: Subsequent accessibility reported selected chapter text but an inconsistent focused search field. No text was typed into that uncertain focus. A later window activation failed; synchronization/hierarchy/ToC behavior was not inferred from that failure.

No reference presentation was saved or vendor template/artwork copied into the product. Native reference rename/reorder/split/navigation and the repaired add-in still need execution evidence. The source fixes below are independently implemented from typed contracts and native PowerPoint operations.

## Implemented repair

Current product details and acceptance evidence are recorded in `docs/implementation/chapter-agenda.md`. The source bug ledger above describes the earlier baseline; it is retained as investigation history.
