/** S2 pure calculations. No Office, DOM, network, state mutation or hidden data edits. */
export interface Category { id: string; label: string }
export interface Series { id: string; name: string; values: Array<number | null> }
export interface Dataset { categories: Category[]; series: Series[] }
export interface Plot { x: number; y: number; width: number; height: number; unit: 'pt' }
export interface Domain { minimum: number; maximum: number }
export interface Bar {
  id: string; categoryId: string; seriesId: string; categoryIndex: number; seriesIndex: number;
  rawValue: number; renderValue: number; base: number; end: number;
  x: number; y: number; width: number; height: number;
}
export interface MissingPoint { categoryId: string; seriesId: string }
export interface Diagnostic { code: string; categoryId?: string }
export class CalculationError extends Error {
  readonly code: string;
  constructor(code: string) { super(code); this.name = 'CalculationError'; this.code = code; }
}
function requireValue(condition: unknown, code: string): asserts condition {
  if (!condition) throw new CalculationError(code);
}
function finite(value: number, code = 'NUMERIC_OVERFLOW'): number {
  requireValue(Number.isFinite(value), code); return value;
}
function item<T>(items: T[], index: number): T {
  const value = items[index]; requireValue(value !== undefined, 'INVALID_INDEX'); return value;
}
function object(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}
function keys(value: Record<string, unknown>, expected: string[]): void {
  requireValue(Object.keys(value).length === expected.length && expected.every(k => Object.hasOwn(value, k)), 'INVALID_FIELDS');
}
export function validateDataset(value: unknown): asserts value is Dataset {
  requireValue(object(value), 'INVALID_DATASET'); keys(value, ['categories', 'series']);
  requireValue(Array.isArray(value.categories) && value.categories.length > 0 && value.categories.length <= 1000, 'INVALID_CATEGORIES');
  requireValue(Array.isArray(value.series) && value.series.length > 0 && value.series.length <= 100, 'INVALID_SERIES');
  const ids = new Set<string>();
  const id = (v: unknown) => {
    requireValue(typeof v === 'string' && /^[A-Za-z0-9][A-Za-z0-9:._-]{0,127}$/.test(v) && !ids.has(v), 'INVALID_OR_DUPLICATE_ID'); ids.add(v);
  };
  const text = (v: unknown) => requireValue(typeof v === 'string' && /\S/u.test(v) && [...v].length <= 1024, 'INVALID_TEXT');
  for (const category of value.categories) {
    requireValue(object(category), 'INVALID_CATEGORY'); keys(category, ['id', 'label']); id(category.id); text(category.label);
  }
  for (const series of value.series) {
    requireValue(object(series), 'INVALID_SERIES'); keys(series, ['id', 'name', 'values']); id(series.id); text(series.name);
    requireValue(Array.isArray(series.values) && series.values.length === value.categories.length, 'DIMENSION_MISMATCH');
    for (const number of series.values) requireValue(number === null || (typeof number === 'number' && Number.isFinite(number)), 'INVALID_NUMBER');
  }
}
export function validatePlot(value: unknown): asserts value is Plot {
  requireValue(object(value), 'INVALID_PLOT'); keys(value, ['x', 'y', 'width', 'height', 'unit']);
  requireValue(value.unit === 'pt', 'INVALID_UNIT');
  for (const key of ['x', 'y', 'width', 'height']) requireValue(typeof value[key] === 'number' && Number.isFinite(value[key]) && Math.abs(value[key] as number) <= 1000000, 'INVALID_PLOT');
  requireValue((value.width as number) > 0 && (value.height as number) > 0, 'INVALID_PLOT');
}
function domain(numbers: number[], includeZero = true): Domain {
  if (numbers.length === 0) return { minimum: -1, maximum: 1 };
  let minimum = includeZero ? 0 : item(numbers, 0), maximum = minimum;
  for (const value of numbers) { minimum = Math.min(minimum, value); maximum = Math.max(maximum, value); }
  if (minimum === maximum) {
    const padding = Math.max(Math.abs(minimum) * 0.1, 1);
    minimum = finite(minimum - padding); maximum = finite(maximum + padding);
  }
  finite(maximum - minimum);
  return { minimum, maximum };
}
function valueCoordinate(value: number, axis: Domain, plot: Plot, orientation: 'vertical' | 'horizontal'): number {
  const ratio = finite((value - axis.minimum) / (axis.maximum - axis.minimum));
  return finite(orientation === 'vertical' ? plot.y + (1 - ratio) * plot.height : plot.x + ratio * plot.width);
}
function rectangle(base: number, end: number, band: number, thickness: number, axis: Domain, plot: Plot, orientation: 'vertical' | 'horizontal') {
  const a = valueCoordinate(base, axis, plot, orientation), b = valueCoordinate(end, axis, plot, orientation);
  return orientation === 'vertical'
    ? { x: plot.x + band, y: Math.min(a, b), width: thickness, height: Math.abs(a - b) }
    : { x: Math.min(a, b), y: plot.y + band, width: Math.abs(a - b), height: thickness };
}
function options(orientation: string, gapRatio: number): asserts orientation is 'vertical' | 'horizontal' {
  requireValue(orientation === 'vertical' || orientation === 'horizontal', 'INVALID_ORIENTATION');
  requireValue(Number.isFinite(gapRatio) && gapRatio >= 0 && gapRatio < 1, 'INVALID_GAP');
}
export function clustered(dataset: Dataset, plot: Plot, orientation: 'vertical' | 'horizontal' = 'vertical', gapRatio = 0.2) {
  validateDataset(dataset); validatePlot(plot); options(orientation, gapRatio);
  const axis = domain(dataset.series.flatMap(s => s.values.filter((v): v is number => v !== null)));
  const categorySpan = (orientation === 'vertical' ? plot.width : plot.height) / dataset.categories.length;
  const groupSpan = categorySpan * (1 - gapRatio), slot = groupSpan / dataset.series.length;
  const bars: Bar[] = [], missing: MissingPoint[] = [];
  dataset.categories.forEach((category, ci) => dataset.series.forEach((series, si) => {
    const value = item(series.values, ci);
    if (value === null) { missing.push({ categoryId: category.id, seriesId: series.id }); return; }
    const band = ci * categorySpan + categorySpan * gapRatio / 2 + si * slot + slot * 0.05;
    bars.push({ id: `bar:${series.id}/${category.id}`, categoryId: category.id, seriesId: series.id,
      categoryIndex: ci, seriesIndex: si, rawValue: value, renderValue: value, base: 0, end: value,
      ...rectangle(0, value, band, slot * 0.9, axis, plot, orientation) });
  }));
  return { sceneVersion: 1, kind: 'clustered', unit: 'pt', plot: { ...plot }, orientation, axis, bars, missing };
}
export function stacked(dataset: Dataset, plot: Plot, orientation: 'vertical' | 'horizontal' = 'vertical', gapRatio = 0.2) {
  return stack(dataset, plot, orientation, gapRatio, false);
}
/** Explicit positive-only policy. Signed 100% policy remains an owner decision. */
export function positive100(dataset: Dataset, plot: Plot, orientation: 'vertical' | 'horizontal' = 'vertical', gapRatio = 0.2) {
  return stack(dataset, plot, orientation, gapRatio, true);
}
function stack(dataset: Dataset, plot: Plot, orientation: 'vertical' | 'horizontal', gapRatio: number, normalized: boolean) {
  validateDataset(dataset); validatePlot(plot); options(orientation, gapRatio);
  const segments: Array<{ ci: number; si: number; raw: number; value: number; base: number; end: number }> = [];
  const totals: Array<{ categoryId: string; positive: number; negative: number; rawTotal: number }> = [];
  const diagnostics: Diagnostic[] = [], missing: MissingPoint[] = [];
  dataset.categories.forEach((category, ci) => {
    const values = dataset.series.map(s => item(s.values, ci));
    if (normalized) requireValue(values.every(v => v === null || v >= 0), 'SIGNED_100_POLICY_REQUIRED');
    let rawTotal = 0;
    for (const value of values) if (value !== null) rawTotal = finite(rawTotal + value);
    if (normalized && rawTotal === 0) diagnostics.push({ code: values.every(v => v === null) ? 'EMPTY_CATEGORY' : 'ZERO_TOTAL', categoryId: category.id });
    let positive = 0, negative = 0;
    values.forEach((raw, si) => {
      if (raw === null) { missing.push({ categoryId: category.id, seriesId: item(dataset.series, si).id }); return; }
      if (normalized && rawTotal === 0) return;
      const value = finite(normalized ? raw / rawTotal * 100 : raw);
      const base = value < 0 ? negative : positive, end = finite(base + value);
      if (value < 0) negative = end; else positive = end;
      segments.push({ ci, si, raw, value, base, end });
    });
    totals.push({ categoryId: category.id, positive, negative, rawTotal });
  });
  const axis = normalized ? { minimum: 0, maximum: 100 } : domain(totals.flatMap(t => [t.positive, t.negative]));
  const span = (orientation === 'vertical' ? plot.width : plot.height) / dataset.categories.length;
  const bars: Bar[] = segments.map(s => ({ id: `bar:${item(dataset.series, s.si).id}/${item(dataset.categories, s.ci).id}`,
    categoryId: item(dataset.categories, s.ci).id, seriesId: item(dataset.series, s.si).id, categoryIndex: s.ci, seriesIndex: s.si,
    rawValue: s.raw, renderValue: s.value, base: s.base, end: s.end,
    ...rectangle(s.base, s.end, s.ci * span + span * gapRatio / 2, span * (1 - gapRatio), axis, plot, orientation) }));
  return { sceneVersion: 1, kind: normalized ? 'positive-100' : 'stacked', unit: 'pt', plot: { ...plot }, orientation, axis, bars, totals, missing, diagnostics };
}
export interface LinePoint { id: string; categoryId: string; seriesId: string; rawValue: number | null; renderValue: number; imputed: boolean; x: number; y: number }
export function line(dataset: Dataset, plot: Plot, settings: { gapPolicy: 'break' | 'connect' | 'zero'; includeZero: boolean }) {
  validateDataset(dataset); validatePlot(plot);
  requireValue(object(settings), 'INVALID_LINE_SETTINGS'); keys(settings, ['gapPolicy', 'includeZero']);
  requireValue(['break', 'connect', 'zero'].includes(settings.gapPolicy) && typeof settings.includeZero === 'boolean', 'INVALID_LINE_SETTINGS');
  const axisValues = dataset.series.flatMap(s => s.values.filter((v): v is number => v !== null));
  if (settings.gapPolicy === 'zero' && dataset.series.some(s => s.values.includes(null))) axisValues.push(0);
  const axis = domain(axisValues, settings.includeZero), span = plot.width / dataset.categories.length;
  const paths = dataset.series.map(series => {
    const points: LinePoint[] = [], segments: LinePoint[][] = []; let current: LinePoint[] = [];
    series.values.forEach((raw, ci) => {
      if (raw === null && settings.gapPolicy !== 'zero') {
        if (settings.gapPolicy === 'break' && current.length) { segments.push(current); current = []; } return;
      }
      const value = raw ?? 0, category = item(dataset.categories, ci);
      const point = { id: `point:${series.id}/${category.id}`, categoryId: category.id, seriesId: series.id, rawValue: raw,
        renderValue: value, imputed: raw === null, x: plot.x + (ci + 0.5) * span, y: valueCoordinate(value, axis, plot, 'vertical') };
      points.push(point); current.push(point);
    });
    if (current.length) segments.push(current);
    return { seriesId: series.id, points, segments };
  });
  return { sceneVersion: 1, kind: 'line', unit: 'pt', plot: { ...plot }, axis, gapPolicy: settings.gapPolicy, paths };
}
export function pie(dataset: Dataset, plot: Plot, startAngle = -90) {
  validateDataset(dataset); validatePlot(plot); requireValue(dataset.series.length === 1, 'PIE_REQUIRES_SINGLE_SERIES'); finite(startAngle, 'INVALID_ANGLE');
  requireValue(Math.abs(startAngle) <= 360, 'INVALID_ANGLE');
  const values = item(dataset.series, 0).values;
  requireValue(values.every(v => v === null || v >= 0), 'NEGATIVE_PIE_VALUE');
  let total = 0; for (const value of values) if (value !== null) total = finite(total + value);
  let angle = startAngle; const positiveCount = values.filter(v => v !== null && v > 0).length;
  let emitted = 0; const diagnostics: Diagnostic[] = [];
  if (total === 0) diagnostics.push({ code: 'ZERO_TOTAL' });
  const sectors: Array<{ id: string; categoryId: string; rawValue: number; fraction: number; startAngle: number; endAngle: number; sweep: number }> = [];
  values.forEach((value, ci) => {
    if (value === null || value === 0 || total === 0) return;
    const fraction = value / total; emitted++;
    const end = emitted === positiveCount ? startAngle + 360 : angle + fraction * 360;
    if (end === angle) diagnostics.push({ code: 'PRECISION_LOSS', categoryId: item(dataset.categories, ci).id });
    sectors.push({ id: `sector:${item(dataset.categories, ci).id}`, categoryId: item(dataset.categories, ci).id, rawValue: value, fraction, startAngle: angle, endAngle: end, sweep: end - angle }); angle = end;
  });
  return { sceneVersion: 1, kind: 'pie', unit: 'pt', plot: { ...plot }, total, center: { x: plot.x + plot.width / 2, y: plot.y + plot.height / 2 }, radius: Math.min(plot.width, plot.height) / 2, sectors, diagnostics };
}
export function area(dataset: Dataset, plot: Plot, settings: { gapPolicy: 'break' | 'connect' | 'zero'; baseline: 0 }) {
  requireValue(object(settings), 'INVALID_AREA_SETTINGS'); keys(settings, ['gapPolicy', 'baseline']);
  requireValue(settings.baseline === 0, 'UNSUPPORTED_AREA_BASELINE');
  const lines = line(dataset, plot, { gapPolicy: settings.gapPolicy, includeZero: true });
  const baselineY = valueCoordinate(0, lines.axis, plot, 'vertical');
  const polygons = lines.paths.flatMap(path => path.segments.map((segment, index) => {
    const first = item(segment, 0), last = item(segment, segment.length - 1);
    return { id: `area:${path.seriesId}/${index}`, seriesId: path.seriesId,
      points: [{ x: first.x, y: baselineY }, ...segment.map(point => ({ x: point.x, y: point.y })), { x: last.x, y: baselineY }] };
  }));
  return { ...lines, kind: 'area', baseline: 0, polygons };
}
