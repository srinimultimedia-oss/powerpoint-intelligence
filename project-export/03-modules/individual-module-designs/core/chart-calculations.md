# Pure JavaScript and TypeScript numerical engines

**VERIFIED source review.** This design describes inspected code, not current runtime/Office execution.

## 1. Purpose

Pure totals/waterfall/CAGR and chart geometry preserving raw values and semantic identity without hidden edits.

## 2. Implementation status

Tooling packages and testable pure functions; not embedded Node service or complete native renderer.

## 3. Source locations

- `packages/chart-core/src/calculations.mjs:1`
- `packages/chart-core/package.json:1`
- `packages/chart-engine/src/basic-charts.ts:1`
- `packages/chart-engine/package.json:1`
- `packages/chart-engine/tsconfig.json:1`

## 4. Internal components

Node >= 20 for chart-core; Node >= 24 for chart-engine; TypeScript 5.9.3 strict noEmit; no Office/DOM/network state. Validate closed normalized Dataset; clustered orientation and gap; mixed stacks accumulate positive/negative separately; positive100 normalizes nonnegative values but retains raw values; line break/connect/zero with an imputation tag; area closes paths to zero; single nonnegative pie final sector closes 360. TypeScript intermediate overflow fails. The smaller JavaScript helpers validate inputs but do not consistently recheck accumulated/result overflow. Waterfall totals/subtotals show the running value without adding twice.

## 5. Classes and interfaces

- `packages/chart-engine/src/basic-charts.ts:2` — `export interface Category`
- `packages/chart-engine/src/basic-charts.ts:3` — `export interface Series`
- `packages/chart-engine/src/basic-charts.ts:4` — `export interface Dataset`
- `packages/chart-engine/src/basic-charts.ts:5` — `export interface Plot`
- `packages/chart-engine/src/basic-charts.ts:6` — `export interface Domain`
- `packages/chart-engine/src/basic-charts.ts:7` — `export interface Bar`
- `packages/chart-engine/src/basic-charts.ts:12` — `export interface MissingPoint`
- `packages/chart-engine/src/basic-charts.ts:13` — `export interface Diagnostic`
- `packages/chart-engine/src/basic-charts.ts:14` — `export class CalculationError extends Error`
- `packages/chart-engine/src/basic-charts.ts:135` — `export interface LinePoint`

## 6. Important methods and functions

Declaration locator below is lexical and may include overload/call-like lines; semantic workflow is described separately.

- `packages/chart-core/src/calculations.mjs:2` — `export function stackedTotals(series)`
- `packages/chart-core/src/calculations.mjs:14` — `export function normalizePositive100Stack(series)`
- `packages/chart-core/src/calculations.mjs:21` — `export function waterfallSegments(points, initial = 0)`
- `packages/chart-core/src/calculations.mjs:38` — `export function cagr(start, end, periods)`
- `packages/chart-engine/src/basic-charts.ts:33` — `export function validateDataset(value: unknown): asserts value is Dataset`
- `packages/chart-engine/src/basic-charts.ts:51` — `export function validatePlot(value: unknown): asserts value is Plot`
- `packages/chart-engine/src/basic-charts.ts:82` — `export function clustered(dataset: Dataset, plot: Plot, orientation: 'vertical' | 'horizontal' = 'vertical', gapRatio = 0.2)`
- `packages/chart-engine/src/basic-charts.ts:98` — `export function stacked(dataset: Dataset, plot: Plot, orientation: 'vertical' | 'horizontal' = 'vertical', gapRatio = 0.2)`
- `packages/chart-engine/src/basic-charts.ts:102` — `export function positive100(dataset: Dataset, plot: Plot, orientation: 'vertical' | 'horizontal' = 'vertical', gapRatio = 0.2)`
- `packages/chart-engine/src/basic-charts.ts:159` — `export function pie(dataset: Dataset, plot: Plot, startAngle = -90)`

## 7. Inputs and outputs

Identified categories/finite number or null series plus a point-based Plot and explicit options -> totals or local geometry/missing/diagnostics.

## 8. Dependencies

Node >= 20 for chart-core; Node >= 24 for chart-engine; TypeScript 5.9.3 strict noEmit; no Office/DOM/network state.

## 9. Workflow

Validate closed normalized Dataset; clustered orientation and gap; mixed stacks accumulate positive/negative separately; positive100 normalizes nonnegative values but retains raw values; line break/connect/zero with an imputation tag; area closes paths to zero; single nonnegative pie final sector closes 360. TypeScript intermediate overflow fails. The smaller JavaScript helpers validate inputs but do not consistently recheck accumulated/result overflow. Waterfall totals/subtotals show the running value without adding twice.

## 10. Integration points

Identified categories/finite number or null series plus a point-based Plot and explicit options -> totals or local geometry/missing/diagnostics. Tooling packages and testable pure functions; not embedded Node service or complete native renderer.

## 11. Exceptions and recovery

CalculationError coded failures in TypeScript, Error in JavaScript. Nonfinite/intermediate overflow, malformed dimensions/options and signed 100% values reject.

## 12. Configuration dependencies

Categories 1–1,000 / series 1–100; point plot<=1 M; bar 0 <= gap < 1; area baseline 0; pie angle [-360,360].

## 13. Security considerations

Pure data transforms; no COM, file, network or DOM or hidden raw value replacement.

## 14. Unit and integration test definitions

- `packages/chart-core/test/calculations.test.mjs:1`
- `packages/chart-engine/test/basic-charts.test.mjs:1`

Inspected definitions only. See the parent current test report for execution; historical passes do not establish current acceptance.

## 15. Known limitations

No nice ticks, logarithmic/time/break axes, font metrics or financial decimal rounding. Signed 100% data is unsupported; pie accepts one series. JavaScript stackedTotals/waterfallSegments/cagr can return nonfinite accumulated/results after finite inputs because they do not consistently recheck outputs; the TypeScript engine does.

## 16. Technical debt

Independent MJS/TS APIs and canonical model must not be conflated with demonstration native commands.

## 17. Extension points

Add explicit policies plus numerical reference/property tests and separate host adapter acceptance.
