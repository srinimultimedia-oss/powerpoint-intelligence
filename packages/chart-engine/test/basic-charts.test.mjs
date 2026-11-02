import test from 'node:test';
import assert from 'node:assert/strict';
import { clustered, stacked, positive100, line, area, pie, validateDataset, CalculationError } from '../src/basic-charts.ts';

const plot = { x: 40, y: 60, width: 600, height: 300, unit: 'pt' };
const data = (rows) => ({ categories: rows[0].map((_, i) => ({ id: `c${i}`, label: `Category ${i}` })), series: rows.map((values, i) => ({ id: `s${i}`, name: `Series ${i}`, values })) });
const close = (actual, expected, tolerance = 1e-9) => assert.ok(Math.abs(actual - expected) <= tolerance * Math.max(1, Math.abs(expected)), `${actual} != ${expected}`);
const rejects = (call, code) => assert.throws(call, error => error instanceof CalculationError && error.code === code);
const geometry = (bars) => { for (const bar of bars) { for (const key of ['x', 'y', 'width', 'height']) assert.ok(Number.isFinite(bar[key])); assert.ok(bar.width >= 0 && bar.height >= 0); } };

test('CHT-001: three series by four categories gives 12 ordered bars and reference coordinates', () => {
  const model = data([[1, 2, 3, 4], [5, 6, 7, 8], [9, 10, 11, 12]]);
  const result = clustered(model, plot);
  assert.equal(result.bars.length, 12);
  assert.deepEqual(result.bars.slice(0, 3).map(b => b.seriesId), ['s0', 's1', 's2']);
  assert.deepEqual(result.bars.filter(b => b.seriesId === 's0').map(b => b.categoryId), ['c0', 'c1', 'c2', 'c3']);
  assert.deepEqual(result.axis, { minimum: 0, maximum: 12 });
  const first = result.bars[0]; close(first.x, 57); close(first.y, 335); close(first.width, 36); close(first.height, 25);
  geometry(result.bars);
});
test('CHT-001: horizontal bars map category order to rows and signed values around zero', () => {
  const result = clustered(data([[-10, 20]]), plot, 'horizontal');
  close(result.bars[0].x, 40); close(result.bars[0].width, 200);
  close(result.bars[1].x, 240); close(result.bars[1].width, 400);
  assert.ok(result.bars[0].y < result.bars[1].y);
});
test('missing cells retain category slots and source identity; zero is not missing', () => {
  const result = clustered(data([[null, 0, 5], [2, 3, 4]]), plot);
  assert.equal(result.bars.length, 5); assert.deepEqual(result.missing, [{ categoryId: 'c0', seriesId: 's0' }]);
  assert.equal(result.bars.find(b => b.categoryId === 'c1' && b.seriesId === 's0').height, 0);
});
test('CHT-013: positive and negative stacks are independent, ordered, and never cross zero', () => {
  const result = stacked(data([[10], [-4], [5], [-6]]), plot);
  assert.deepEqual(result.bars.map(b => [b.base, b.end]), [[0, 10], [0, -4], [10, 15], [-4, -10]]);
  assert.deepEqual(result.axis, { minimum: -10, maximum: 15 });
  assert.deepEqual(result.totals, [{ categoryId: 'c0', positive: 15, negative: -10, rawTotal: 5 }]);
});
test('CHT-002: 20/30/50 maps to exact 20/30/50 percent without replacing raw values', () => {
  const result = positive100(data([[20], [30], [50]]), plot);
  assert.deepEqual(result.bars.map(b => b.renderValue), [20, 30, 50]);
  assert.deepEqual(result.bars.map(b => [b.base, b.end]), [[0, 20], [20, 50], [50, 100]]);
  close(result.bars[0].height, 60);
  const unequal = positive100(data([[2], [3], [5]]), plot);
  assert.equal(unequal.bars[0].rawValue, 2); assert.equal(unequal.bars[0].renderValue, 20);
});
test('100% empty and zero-total categories are explicit diagnostics, not NaN geometry', () => {
  const result = positive100(data([[null, 0], [null, 0]]), plot);
  assert.deepEqual(result.bars, []);
  assert.deepEqual(result.diagnostics.map(d => d.code), ['EMPTY_CATEGORY', 'ZERO_TOTAL']);
});
test('signed 100% data is rejected until a production policy is approved', () => rejects(() => positive100(data([[10], [-5]]), plot), 'SIGNED_100_POLICY_REQUIRED'));
test('CHT-012: break leaves separate contiguous paths around a missing value', () => {
  const result = line(data([[10, null, 30]]), plot, { gapPolicy: 'break', includeZero: true });
  assert.deepEqual(result.paths[0].segments.map(s => s.length), [1, 1]);
  close(result.paths[0].points[0].x, 140); close(result.paths[0].points[0].y, 260);
});
test('CHT-012: connect keeps source nulls absent; zero imputation is explicitly tagged', () => {
  const model = data([[10, null, 30]]);
  const connected = line(model, plot, { gapPolicy: 'connect', includeZero: true });
  assert.equal(connected.paths[0].segments[0].length, 2);
  const zero = line(model, plot, { gapPolicy: 'zero', includeZero: false });
  const point = zero.paths[0].points[1]; assert.equal(point.rawValue, null); assert.equal(point.renderValue, 0); assert.equal(point.imputed, true); close(point.y, 360);
  assert.equal(model.series[0].values[1], null);
});
test('line gap and baseline settings must be explicit and valid', () => {
  rejects(() => line(data([[1]]), plot, { gapPolicy: 'guess', includeZero: true }), 'INVALID_LINE_SETTINGS');
  rejects(() => line(data([[1]]), plot, { gapPolicy: 'break' }), 'INVALID_FIELDS');
});
test('CHT-012: area closes at zero baseline and respects break/zero gap policies', () => {
  const broken = area(data([[10, null, -10]]), plot, { gapPolicy: 'break', baseline: 0 });
  assert.equal(broken.polygons.length, 2);
  for (const polygon of broken.polygons) { close(polygon.points[0].y, 210); close(polygon.points.at(-1).y, 210); }
  const zero = area(data([[10, null, -10]]), plot, { gapPolicy: 'zero', baseline: 0 });
  assert.equal(zero.polygons.length, 1); assert.equal(zero.paths[0].points[1].imputed, true);
  rejects(() => area(data([[1]]), plot, { gapPolicy: 'break', baseline: 10 }), 'UNSUPPORTED_AREA_BASELINE');
});
test('CHT-010: 25/25/50 yields 90/90/180 degrees and exactly closes the circle', () => {
  const result = pie(data([[25, 25, 50]]), plot);
  assert.deepEqual(result.sectors.map(s => s.sweep), [90, 90, 180]);
  assert.equal(result.sectors.at(-1).endAngle, 270); assert.equal(result.radius, 150); assert.deepEqual(result.center, { x: 340, y: 210 });
});
test('pie preserves missing and zero inputs without dividing by zero', () => {
  const result = pie(data([[null, 0]]), plot); assert.deepEqual(result.sectors, []); assert.equal(result.total, 0); assert.equal(result.diagnostics[0].code, 'ZERO_TOTAL');
});
test('pie rejects negatives and multiple series rather than silently discarding data', () => {
  rejects(() => pie(data([[1, -1]]), plot), 'NEGATIVE_PIE_VALUE');
  rejects(() => pie(data([[1], [2]]), plot), 'PIE_REQUIRES_SINGLE_SERIES');
});
test('dataset validation rejects unknown fields, duplicates, empty categories, coercion and nonfinite values', () => {
  const cases = [
    [{ ...data([[1]]), injected: true }, 'INVALID_FIELDS'],
    [data([[]]), 'INVALID_CATEGORIES'],
    [data([[NaN]]), 'INVALID_NUMBER'], [data([[Infinity]]), 'INVALID_NUMBER'], [data([['1']]), 'INVALID_NUMBER']
  ];
  const duplicate = data([[1, 2]]); duplicate.categories[1].id = 'c0'; cases.push([duplicate, 'INVALID_OR_DUPLICATE_ID']);
  const mismatch = data([[1]]); mismatch.series[0].values.push(2); cases.push([mismatch, 'DIMENSION_MISMATCH']);
  for (const [model, code] of cases) rejects(() => validateDataset(model), code);
});
test('plot, orientation and gap validation enforce finite point geometry', () => {
  const model = data([[1]]);
  for (const width of [0, -1, NaN, Infinity]) rejects(() => clustered(model, { ...plot, width }), 'INVALID_PLOT');
  rejects(() => clustered(model, { ...plot, unit: 'px' }), 'INVALID_UNIT');
  rejects(() => clustered(model, plot, 'diagonal'), 'INVALID_ORIENTATION');
  rejects(() => clustered(model, plot, 'vertical', 1), 'INVALID_GAP');
});
test('overflow is rejected even when each source value is individually finite', () => {
  rejects(() => stacked(data([[Number.MAX_VALUE], [Number.MAX_VALUE]]), plot), 'NUMERIC_OVERFLOW');
  rejects(() => pie(data([[Number.MAX_VALUE, Number.MAX_VALUE]]), plot), 'NUMERIC_OVERFLOW');
  rejects(() => clustered(data([[Number.MAX_VALUE, -Number.MAX_VALUE]]), plot), 'NUMERIC_OVERFLOW');
});
test('pure results are deterministic, input-immutable, and IDs survive numeric edits', () => {
  const model = data([[1, 2], [3, 4]]), before = structuredClone(model);
  assert.deepEqual(clustered(model, plot), clustered(model, plot)); assert.deepEqual(model, before);
  const changed = structuredClone(model); changed.series[0].values[0] = 99;
  assert.deepEqual(clustered(model, plot).bars.map(b => b.id), clustered(changed, plot).bars.map(b => b.id));
});
test('all engines operate on frozen caller inputs', () => {
  const model = data([[1, 2], [3, 4]]);
  for (const category of model.categories) Object.freeze(category);
  for (const series of model.series) { Object.freeze(series.values); Object.freeze(series); }
  Object.freeze(model.categories); Object.freeze(model.series); Object.freeze(model); Object.freeze(plot);
  clustered(model, plot); stacked(model, plot); positive100(model, plot);
  line(model, plot, { gapPolicy: 'break', includeZero: true }); area(model, plot, { gapPolicy: 'break', baseline: 0 });
  const single = data([[25, 25, 50]]); Object.freeze(single.series[0].values); pie(single, plot);
});
test('seeded property fixtures keep 500 signed-stack geometries finite, ordered and bounded', () => {
  let seed = 20261022;
  const random = () => { seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0; return seed / 2 ** 32; };
  for (let iteration = 0; iteration < 500; iteration++) {
    const columns = 1 + Math.floor(random() * 10), series = 1 + Math.floor(random() * 6);
    const model = data(Array.from({ length: series }, () => Array.from({ length: columns }, () => random() < 0.1 ? null : Math.round((random() - 0.5) * 20000) / 100)));
    const result = stacked(model, plot, iteration % 2 ? 'horizontal' : 'vertical'); geometry(result.bars);
    assert.equal(new Set(result.bars.map(b => b.id)).size, result.bars.length);
    for (const bar of result.bars) {
      assert.ok(bar.x >= plot.x - 1e-8 && bar.y >= plot.y - 1e-8);
      assert.ok(bar.x + bar.width <= plot.x + plot.width + 1e-8 && bar.y + bar.height <= plot.y + plot.height + 1e-8);
      assert.ok(bar.renderValue < 0 ? bar.end <= bar.base && bar.base <= 0 : bar.end >= bar.base && bar.base >= 0);
    }
    for (let ci = 0; ci < columns; ci++) close(result.totals[ci].rawTotal, model.series.reduce((sum, s) => sum + (s.values[ci] ?? 0), 0));
  }
});
