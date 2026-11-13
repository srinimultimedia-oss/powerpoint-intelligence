# Shape, text, table and slide operation design

## Operation traceability

| Operation | Actual implementation | Native result and constraints |
|---|---|---|
| Text box/rounded/pentagon/chevron | `apps/powerpoint-vsto/TextShapeElements.cs:103` | AutoShape with editable TextFrame, center/middle default, manual geometry and XML. Unsupported native rotation/flip/kind changes fail closed. |
| Fit/convert/adjacent duplicate | `apps/powerpoint-vsto/TextShapeElements.cs:117` | Plan bounds before edit; unlocked edges may grow; all-locked unsupported overflow stops. |
| Process grid | `apps/powerpoint-vsto/ProcessFlowTools.cs:95` | Managed cells plus deck registry; hide keeps inactive history; manual independent cell geometry disallows structural edits. |
| Harvey/checkbox | `apps/powerpoint-vsto/IndicatorElements.cs:39` | Native grouped vector sector/strokes/text; state/colors/rotation XML; update replacement retains original if removal fails. |
| Universal connector | `apps/powerpoint-vsto/UniversalConnectors.cs:114` | Native straight/elbow/curve, free/site/chart-point endpoint modes; rotation/flip captured in slide coordinates. |
| Reference/difference annotation | `apps/powerpoint-vsto/ChartAnnotations.cs:41` | Editable native group, chart identity tags and explicit refresh; no generic obstacle-avoidance engine. |
| Native linked table | `apps/powerpoint-vsto/DemoTools.cs:252` | Native Table <=400 cells, centered headers/body, source and table snapshot persisted. Datasheet edits are local native edits. |
| Generated template | `apps/powerpoint-vsto/TemplateLibrary.cs:63` | Native shapes/text/freeforms, optional staged builds and tagged editable fields. |
| Original slide/template | `apps/powerpoint-vsto/TemplateLibrary.cs:489` | InsertFromFile into existing deck; clone source Design/CustomLayout; same-size requirement. |
| Agenda | `apps/powerpoint-vsto/ChapterAgenda.cs:84` | Native editable chapter/contents slides, hierarchy/page links/captions/scopes. |
| Arrange selected objects | `apps/powerpoint-vsto/DemoTools.cs:222` | Native coordinates/size/group/order/flip/rotate/duplicate; first selected shape is alignment/size reference. |
| Layout links | `apps/powerpoint-vsto/TemplateLibrary.cs:411` | Leader/follower edge/offset rules reflow through timer; dependencies are validated, not freeform full-slide constraint solving. |
| Raster copy | `apps/powerpoint-vsto/ProductivityTools.cs:79` | New embedded picture; original editable object is retained. |
| Read-only quality/navigation | `apps/powerpoint-vsto/PresentationQuality.cs:80` | Typed snapshots and issue locate; excludes masters/notes/SmartArt/tablecells/chart-internal text. |

## Text and table scoping

NativeFormatting honors selected text scope or selected table cells; no selected cells means whole table. It preflights nested groups and budgets before constructing apply/restore changes. Context property capture separately limits selected characters/parts for responsive inspection. Font-only updates leave unrequested bold/italic attributes untouched and mixed values are not coerced into an invented default.

## Slide ownership and document selection

Generated linked report appends seven new chart slides to the same active presentation (`apps/powerpoint-vsto/DemoTools.cs:139`). Ordinary insertion operates on current active slide. Hidden read-only template sources are explicitly owned session objects; user-open sources are rejected and never closed by the library. Keep-source insertion does not resize the destination deck implicitly. Save selected slides creates a user-chosen output presentation; that export path is distinct from generating a linked chart in the current deck.

## Transitions and animations

Designed-template staged builds support Fade/Wipe/Fly-in with click trigger and.35s duration; new-slide transitions None/Fade/Wipe. They are ordinary editable native timeline effects. No Morph/advanced effect-editor equivalence is claimed. Source-original slide animation/fidelity is preserved by native file/design transfer where verified, not reconstructed from bitmap previews. Static object-only insertion rejects animated builds explicitly.

## Failure and extension policy

Typed planners/codecs own finite/bounded geometry, valid IDs, shape ownership and supported modes. Native operations can fail independently during cleanup; surfaces report that state and offer Undo guidance rather than destroying foreign objects. Adding any new native object mode requires source/pure-plan tests and desktop edit/save/reopen tests, with copied/rotated/foreign-object negative scenarios.
