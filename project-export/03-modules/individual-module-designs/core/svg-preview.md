# Internal SVG geometry inspection

**VERIFIED source review.** This design describes inspected code, not current runtime/Office execution.

## 1. Purpose

Deterministic accessible SVG preview strings for six numerical families.

## 2. Implementation status

Inspection aid, not persisted scene or native editable insertion.

## 3. Source locations

- `packages/chart-engine/src/svg-preview.ts:1`
- `packages/chart-engine/README.md:1`

## 4. Internal components

Pure basic-charts only; no DOM/rendering package. Emit fixed SVG geometry, escape title and semantic IDs, preserve missing breaks, circle forfull pie; explicit point viewBox/dimensions.

## 5. Classes and interfaces

- `packages/chart-engine/src/svg-preview.ts:4` — `export type PreviewKind = 'clustered' | 'stacked' | 'positive-100' | 'line' | 'area' | 'pie';`

## 6. Important methods and functions

Declaration locator below is lexical and may include overload/call-like lines; semantic workflow is described separately.

- `packages/chart-engine/src/svg-preview.ts:14` — `export function svgPreview(dataset: Dataset, plot: Plot, kind: PreviewKind, title = 'Chart preview'): string`

## 7. Inputs and outputs

Dataset + Plot + kind + title -> SVG string; input unchanged.

## 8. Dependencies

Pure basic-charts only; no DOM/rendering package.

## 9. Workflow

Emit fixed SVG geometry, escape title and semantic IDs, preserve missing breaks, circle forfull pie; explicit point viewBox/dimensions.

## 10. Integration points

Dataset + Plot + kind + title -> SVG string; input unchanged. Inspection aid, not persisted scene or native editable insertion.

## 11. Exceptions and recovery

CalculationError on invalid title/kind/plot/nonfinite coordinates.

## 12. Configuration dependencies

Fixedsix-color palette; title<=1024 code points; line/area break policy.

## 13. Security considerations

No script/foreignObject/href/images; user text escaped.

## 14. Unit and integration test definitions

- `packages/chart-engine/test/svg-preview.test.mjs:1`

Inspected definitions only. See the parent current test report for execution; historical passes do not establish current acceptance.

## 15. Known limitations

No axes, legend or labels/fonts/themes/editing handles or Office fidelity guarantee.

## 16. Technical debt

Preview and scenebuilder parallel entry points need policy-parity regression.

## 17. Extension points

Extend display only without substitutingSVG for editable native objects.
