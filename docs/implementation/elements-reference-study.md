# Elements public-reference study

Date: 2026-10-28. Documentation-only independent investigation requested under REA skill. AGENTS.md was read. No vendor binaries, protected internals/assets or user desktop were inspected. The parent owns native observation; source observations below precede concurrent implementation changes. Official search results exposed the manual text; direct website requests sometimes redirect to an inaccessible domain. No runtime or complete-parity acceptance is asserted.

## Public behavior synopses

Each synopsis is below 150 words from its cited source; these describe reference behavior, not implemented capabilities.

### Text box

[Text boxes](https://www.think-cell.com/en/resources/manual/textbox): Click inserts a default-size box; drag chooses custom bounds, then type. Content-driven growth depends on edge locks: left-click locks a handle, right-click unlocks it. Locking both opposite edges fixes that dimension. Top locked/bottom unlocked grows down; right locked/left unlocked grows left. Drag an edge handle to another element's edge to establish dynamic alignment; participating unlocked edges can move or resize as content changes. Moving an element, locking/dragging an aligned edge, or dragging its opposite edge breaks alignment. Comparable resizing/alignment applies to other eligible text-bearing or point-sized elements, including indicators and process shapes. A one-time Align command does not establish this continuing relationship.

### Pentagon / chevron

[Process flows](https://www.think-cell.com/en/resources/manual/processflow): Insert starts with an editable pentagon fitted to its contents. Side duplication copies text and automatically aligns; Ctrl+Alt+arrow duplicates directionally. Pentagon-to-chevron conversion occurs when needed. Multi-selection supports duplication; duplicating the header row downward creates aligned text-box rows. Alt+arrows navigates between elements. Font, fill, outline, horizontal/vertical alignment and alternating row/column fills are configurable. This is an interacting process/table layout, not simply four unrelated arrows. The page does not specify the numeric tip/notch angle, corner radius or exact conversion threshold.

### Rounded rectangle

[Rounded rectangles](https://www.think-cell.com/en/resources/manual/rounded-rectangles): Insert the element and type directly into it. The rectangle can automatically fit its text, using the linked text-box resizing behavior and general element formatting. This page does not establish a special corner-radius control, a required radius or a particular adjustment formula. Native rounded geometry alone is not proof of equivalent layout behavior.

### Harvey ball / checkbox

[Harvey balls and checkboxes](https://www.think-cell.com/en/resources/manual/harvey-balls-and-checkboxes): Types can be switched. Harvey balls default to four segments, support custom segment counts, filled/unfilled styling, outline, alignment and background. Completion increases with + or double-click, decreases with -. Keys 0–8 choose counts for up to eight segments; larger segment counts use eighth fractions. Checkbox states are empty, check and X: 0/Space, 1/v, 2/x; double-click cycles. Symbol/color, fill, outline and background are configurable; style files can supply custom symbols. Size accepts points, mm, cm, inches or slide-width/height proportions; unitless values mean points. Suite supports external links. The page does not specify completion wrap/clamp behavior or rounding of fractional segment counts.

### Universal connector

[Chart annotations: universal connectors](https://www.think-cell.com/en/resources/manual/chartdecorations): A universal connector is independently inserted, rather than owned by one chart feature. It joins available connectable points within one element or across different elements/charts. End handles snap to connection points. Dragging its middle handle changes a straight path to a right-angle path. Double-clicking the ribbon command enables repeated insertion; clicking again or Esc exits. It creates a visual relationship only and does not participate in waterfall calculations. The page does not define an obstacle-avoidance algorithm, every eligible anchor, unrestricted waypoint editing or automatic behavior for arbitrary native shape types.

### Linked data and editions

[External data links](https://www.think-cell.com/en/resources/manual/external-data-links): Current Suite/AI Suite documentation supports external links. Excel can drive charts, tables, images, Harvey balls and checkboxes; Tableau targets charts/tables. Updates, reverting the last update and switching sources are managed link operations. [Tables with datasheets](https://www.think-cell.com/en/resources/manual/tables-with-datasheets) separately specifies Harvey values as filled-segment numbers and checkbox source mappings v/o/1, x/2, Space/0; invalid cell content produces an error indicator. Keyboard mappings and linked-cell mappings are distinct contracts. [Element datasheets](https://www.think-cell.com/en/resources/manual/element-datasheets) describes associated datasheets in Suite/AI Suite. Installed edition/version must be observed; current website terminology does not establish compatibility with every historic release.

### Tools ambiguity

[Presentation basics](https://www.think-cell.com/en/resources/manual/presentation-basics) groups productivity tasks as align/match size, swap, scale shapes, replace/resize fonts, insert symbols, switch decimal symbols, cleanup/sanitize and send/save slides. This is not an exhaustive definition of the legacy Tools dropdown. [Organization resource links](https://www.think-cell.com/en/resources/manual/tools-menu-links) adds policy-configured buttons opening URLs/documents or executing programs/scripts. These are administrative integrations and must not be copied as arbitrary execution behavior. Other manual pages place Help/About, Change Style, Slide Workbook and Enable Data Tables for New Charts under Tools. Exact installed menu inventory, labels, enabled states and edition boundaries require the parent's native observation; a generic Tools button cannot substantiate all of them.

## Acceptance and source-gap ledger

Baseline read: DemoTools.cs, FeatureInventory.cs. Concurrent branches may subsequently resolve rows; preserve this as the discovery baseline, not current pass status.

| Area | Source baseline / implementation distinction | Required acceptance scenarios |
|---|---|---|
| Text box | text command inserts fixed native box with Your text; global formatting/alignment exists separately. No per-edge constraints or continuing relationship evidenced by this source. | Custom/default insertion; direct text edit; each growth direction; both edges locked; long/wrapped text; dynamic partner resize; explicit relation break; save/reopen; duplicate-ID handling. |
| Pentagon/chevron | process calls ProcessFlow: four grouped native chevrons, Step 1–4. No distinct starting right-pentagon command or directional semantic duplicate established. | Standalone pentagon; editable tip/notch control only if independently designed; side duplicate; convert where needed; navigate; text-row duplicate; mixed selection; row/column fills; native editing/save/reopen. |
| Rounded rectangle | rounded adds native rounded rectangle at fixed bounds. | Direct edit and layout growth; formatting preserved after refresh; text margins/centering; shape adjustment round-trip; disclose unsupported edge-lock relation rather than claiming it. |
| Harvey ball | harvey inserts a half-circle Unicode glyph; catalog explicitly calls static demo. | Numeric canonical segment/completion state; zero/full/custom counts; reject invalid/nonfinite/out-of-range input; state edit without glyph-font dependence; fills/outline/size/type conversion; native persistence. |
| Checkbox | checkbox inserts a fixed checked Unicode glyph. | Three explicit states; symbol color vs box fill; keyboard/double-click only if actually hooked and scoped; type switch; per-user custom symbol scope; preserve native editability and formatting after state changes. |
| Universal connector | No universal-connector command or durable endpoint model was found in the inspected DemoTools baseline. Existing graph/diagram lines are not proof of dynamic attachment. | Two explicit endpoint choices; native attachment; move/resize both objects; straight/elbow route; reconnect/detach; same-object connection if supported; cross-chart anchors explicitly scoped; deleted endpoint handling; save/reopen; Undo; no waterfall data changes. |
| Linked indicators | Chart data-source/link infrastructure does not establish indicator linkage. | Linked single-cell mapping, missing/invalid values, retained last-good display plus diagnosis, rebind/update/revert, moved-file policy, no fabricated live results. |
| Tools | Multiple scattered native utility commands and read-only quality review are subsets. | Inventory each observed menu entry; map implemented UI/API/acceptance independently. A read-only inspector must not claim destructive cleanup/sanitization. |

## Priority implementation contracts and unknowns

1. Replace static indicator glyphs with versioned, bounded state owned by independent code. Define transition behavior explicitly; public evidence does not settle overflow/wrap semantics. Do not install global keyboard hooks merely to imitate shortcuts.
2. Separate text-bearing element creation, per-shape settings, layout relations and format commands. Explicit refresh is a legitimate interim implementation but must be labeled; native AutoSize alone cannot prove dynamic multi-object layout.
3. For connectors, persist stable slide/shape identities and native connection-site choices, not object names or collection positions. Reconcile deleted/copied endpoints and reopen; retain visuals on unsupported metadata. These are implementation recommendations, not claims about vendor internals.
4. Define Tools scope using the installed menu observation. No request to change OS policy, execute organization scripts, strip document metadata or remove comments/notes is inferred from the word Tools. Offer concrete selected operations with reviewable scope.
5. Validate Office 2016/2019/365 x64 separately from vendor edition/version. Chart-internal connectable feature points may require our own renderer anchor contract; arbitrary native Office connector sites are not equivalent evidence.

Open reference questions: exact custom-segment bound and eighth rounding, cycle endpoints, custom-symbol persistence, automatic duplicate/conversion rules, all chart anchor kinds, connector routing after obstruction or rotation, and the installed Tools command inventory. Keep these as unknowns until observed/documented. No source or native tests were executed in this documentation task; implementation and end-to-end evidence remain acceptance gates.

## Installed Tools inventory — parent-supplied runtime evidence

Observed 2026-10-28 in PowerPoint Presentation1, window handle 1708252. The parent supplied this exact menu inventory from native UI observation. This agent did not operate the desktop; presence of a command establishes its label/group only, not its behavior or successful execution. This observation supersedes the earlier unknown inventory while retaining unknown command semantics.

| Observed group | Exact observed item | Repository mapping / remaining gap |
|---|---|---|
| Editing | Insert Symbol | DemoTools symbol prompts for arbitrary Unicode/text and creates a new textbox. No searchable symbol picker or insertion at the existing text caret is established. |
| Editing | Switch Decimal Symbols | DemoTools decimal swaps digit-adjacent separators in selected shapes only. It assigns the full TextRange.Text, risking loss of mixed run formatting. Scope preview, table/chart policy and run-preserving edits remain local gaps. |
| Editing | Convert to thinkcell Image | No matching conversion command identified. Export slide image saves an entire slide PNG; inserting a file is also different. A local selected-object-to-native-picture conversion can be independently implemented, with explicit replacement/copy policy. |
| Editing | Import Mekko Graphics to thinkcell | No compatible importer established. Our signed/percent Mekko renderers consume our data model, not Mekko Graphics objects. Requires a public format/authorized source contract; do not inspect proprietary binary internals or label ordinary Mekko rendering an import. |
| Editing | Clean Up | Confirmed legacy-comment removal exists; PresentationQuality is a read-only inspector. Neither supplies a comprehensive scope/options cleanup operation. Any destructive local cleanup needs explicit per-option review and preserved source/Undo evidence. |
| Data | Slide Workbook | DemoTools slide-workbook lists chart names only. Native chart-specific embedded workbooks exist, but a single shared slide-level workbook and editable cell-to-element mapping do not. |
| Data | Data Links | ThisAddIn data-manager opens DataLinkManager with actual mapped-link refresh callback. DemoTools manage-links still describes/counts legacy local links. Unify entry points and test each supported mapping/update/rebind rather than claiming every source supported. |
| Collaboration | Send Slides | No send workflow identified. Independent implementation may prepare a user-reviewed file/draft; no permission to send email is inferred from this menu study. |
| Collaboration | Save Slides | DemoTools SaveSlides copies selected native slides to a new PPTX through PowerPoint; the tooltip's word thumbnails is inaccurate. Verify source formatting, dimensions, selection order, hidden-slide policy and reopened native objects. PDF/export options are separate existing functions, not automatically integrated. |
| Collaboration | Invite to thinkcell Trial | Vendor-specific commercial invitation. No local equivalent required for functional slide editing; do not reproduce vendor invitations or impersonate the vendor. |
| Current Style | generic / Change Style submenu | DemoTools style-file reads an original JSON subset (FontName, FontSize, FillRgb, TextRgb) and applies selected shapes. It is not a presentation-level current-style service or compatible vendor style loader. Scope, theme references and style persistence need independent contracts. |
| Settings | Enable Data Tables for New Charts | Existing data-table toggles the selected native chart. It does not persist a default for newly inserted charts. An independently named default setting can be implemented and tested by chart family. |
| Settings | Restore thinkcell Tab | Vendor ribbon administration. Our ribbon could expose an independent supported settings/help action; changing the vendor ribbon is outside our feature implementation. |
| Settings | License Key | Vendor licensing, excluded from replication. Do not copy, circumvent or alter vendor licensing. |
| Settings | Help submenu | Exact submenu items were not supplied. Our own help/About/diagnostics links are appropriate; compatibility with unspecified vendor help actions is not asserted. |
| Settings | Deactivate thinkcell | Vendor add-in lifecycle administration. Do not deactivate another installed add-in as part of local feature parity. |

### Fillable local priorities

1. **Low-risk local utilities:** searchable original symbol picker; selected-object image conversion with retain-original default; run-preserving decimal edits with preview and explicit text/table scope; fix misleading Save Slides tooltip and verify native export formatting.
2. **Independent user settings:** supported new-chart data-table default and original style settings with declared scope, persistence and validation. Do not claim vendor XML/style compatibility.
3. **Larger contracts:** shared slide workbook requires its own mapping/persistence/error-recovery design; a single-chart workbook launcher is an honest interim subset. Consolidate mapped Data Links UI and eliminate label-only legacy routes.
4. **Controlled cleanup:** a review dialog can offer only implemented, separately tested options. Current read-only inspection should remain distinct from cleanup. Metadata/notes removal is not implicitly authorized merely by observing a vendor label.
5. **Evidence-dependent import:** Mekko Graphics conversion remains unsupported until a public/approved interchange contract and fixture establish inputs. Vendor administration and invitations are excluded from the implementation gap count.

Suggested regression cases: mixed-font decimal text preserves runs; punctuation and dates require explicit policy; Unicode surrogate pairs and caret position remain intact; image conversion preserves bounds/rotation/visual appearance and source policy; slide export preserves order/size/source formatting and editability on reopen; new-chart defaults do not modify existing charts; Cancel on any review writes nothing; unsupported input is retained with a clear diagnosis.
