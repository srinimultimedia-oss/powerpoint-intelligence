import { test } from 'node:test';
import assert from 'node:assert/strict';
import { svgPreview } from '../src/svg-preview.ts';
const data = { categories: [{ id: 'a', label: 'A' }, { id: 'b', label: 'B' }, { id: 'c', label: 'C' }], series: [{ id: 's', name: 'Series', values: [1, null, 3] }] };
const plot = { x: 40, y: 60, width: 600, height: 300, unit: 'pt' };
test('previews all six engines without changing input', () => {
  const before = JSON.stringify(data);
  for (const kind of ['clustered', 'stacked', 'positive-100', 'line', 'area', 'pie']) {
    const svg = svgPreview(data, plot, kind);
    assert.equal(svg, svgPreview(data, plot, kind));
    assert.match(svg, /^<svg xmlns="http:\/\/www.w3.org\/2000\/svg"/);
    assert.ok(!/NaN|Infinity|undefined/.test(svg));
    assert.ok(!/script|foreignObject|href=/.test(svg));
  }
  assert.equal(JSON.stringify(data), before);
});
test('line and area preserve missing-value breaks', () => {
  for (const kind of ['line', 'area']) {
    const svg = svgPreview(data, plot, kind);
    assert.equal((svg.match(/<polyline /g) ?? []).length, 2);
    assert.equal((svg.match(/<circle /g) ?? []).length, 2);
  }
});
test('escapes title markup and rejects unsupported or invalid inputs', () => {
  assert.match(svgPreview(data, plot, 'line', '<script>"&\''), /<title>&lt;script&gt;&quot;&amp;&apos;<\/title>/);
  assert.throws(() => svgPreview(data, plot, 'foreignObject'), /UNSUPPORTED_PREVIEW_KIND/);
  assert.throws(() => svgPreview(data, { ...plot, width: Infinity }, 'line'));
  assert.throws(() => svgPreview(data, plot, 'line', 'a'.repeat(1025)));
});
test('single-value pie uses a full circle, empty pie emits no sector', () => {
  const one = { ...data, series: [{ ...data.series[0], values: [2, null, 0] }] };
  assert.equal((svgPreview(one, plot, 'pie').match(/<circle /g) ?? []).length, 1);
  assert.ok(!svgPreview(one, plot, 'pie').includes('<path'));
  one.series[0].values = [0, null, 0];
  assert.ok(!/<circle|<path/.test(svgPreview(one, plot, 'pie')));
});
