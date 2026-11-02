import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { join } from 'node:path';
import { buildScene } from '../../packages/chart-engine/src/scene-builder.ts';
import { renderChartV2 } from '../../packages/chart-engine/tools/render-chart-v2.mjs';

// Independent test oracle: Buffer's UTF-8 and big-endian encoding, not the
// production browser encoder. Hash all fields except the outer renderHash.
export function referenceHash(scene) {
  const parts = [], length = n => { const b = Buffer.alloc(4); b.writeUInt32BE(n); return b; };
  const string = s => { const b = Buffer.from(s, 'utf8'); parts.push(length(b.length), b); };
  const encode = v => {
    if (v === null) parts.push(Buffer.from([0]));
    else if (typeof v === 'boolean') parts.push(Buffer.from([v ? 2 : 1]));
    else if (typeof v === 'number') { const b = Buffer.alloc(9); b[0] = 3; b.writeDoubleBE(v === 0 ? 0 : v, 1); parts.push(b); }
    else if (typeof v === 'string') { parts.push(Buffer.from([4])); string(v); }
    else if (Array.isArray(v)) { parts.push(Buffer.from([5]), length(v.length)); v.forEach(encode); }
    else { const keys = Object.keys(v).sort(); parts.push(Buffer.from([6]), length(keys.length)); for (const k of keys) { encode(k); encode(v[k]); } }
  };
  string('PPI-SCENE-1'); const { renderHash, ...payload } = scene; encode(payload);
  return createHash('sha256').update(Buffer.concat(parts)).digest('hex');
}
const chartId = '7bef06ad-f023-4b26-9a32-0d51b757bf59';
const paint = { fill: '#3b82f6', stroke: null, strokeWidth: 0, opacity: 1 };
export const primitives = {
  schemaVersion: 1, chartId, modelRevision: 7, unit: 'pt', plot: { x: 40, y: 60, width: 600, height: 300, unit: 'pt' },
  semanticScene: [
    { id: chartId + '/bar:actual/q1', geometry: { type: 'rectangle', x: 57, y: 335, width: 36, height: 25 }, paint },
    { id: chartId + '/marker:actual/q1', geometry: { type: 'ellipse', x: 1e-7, y: -0, width: 6, height: 6 }, paint },
    { id: chartId + '/line:actual/q1/q2', geometry: { type: 'polyline', points: [{ x: 0, y: 0 }, { x: 1.2345678901234567, y: 2.5 }] }, paint: { ...paint, fill: null, stroke: '#0D9488', strokeWidth: 2 } },
    { id: chartId + '/area:actual/q1/q2', geometry: { type: 'polygon', points: [{ x: 0, y: 0 }, { x: 50, y: 25 }, { x: 100, y: 0 }] }, paint: { ...paint, opacity: .35 } },
    { id: chartId + '/sector:q1', geometry: { type: 'sector', cx: 200, cy: 150, radius: 50, startAngle: -90, sweep: 180 }, paint },
    { id: chartId + '/label:q1', geometry: { type: 'text', x: 40, y: 25, width: 300, height: 30, text: 'Δ Revenue 😀\nLine 2', fontFamily: 'Segoe UI', fontSize: 12, alignment: 'left' }, paint }
  ], diagnostics: [{ code: 'EXAMPLE_WARNING', nodeId: chartId + '/label:q1' }]
};

export async function createCorpus() {
  const golden = JSON.parse(readFileSync(new URL('./fixtures/all-primitives.v1.json', import.meta.url), 'utf8'));
  if (referenceHash(golden) !== golden.renderHash) throw Error('Committed primitive golden hash mismatch');
  const cases = [{ name: 'all-primitives', valid: true, schemaValid: true, json: JSON.stringify(golden) }];
  const mutate = (name, update, valid = false, schemaValid = false) => {
    const model = structuredClone(golden); update(model); model.renderHash = referenceHash(model);
    cases.push({ name, valid, schemaValid, json: JSON.stringify(model) });
  };
  for (const [name, update] of [
    ['version', m => m.schemaVersion = 2], ['revision-negative', m => m.modelRevision = -1], ['revision-fraction', m => m.modelRevision = 1.5],
    ['revision-string', m => m.modelRevision = '7'], ['revision-overflow', m => m.modelRevision = 2147483648], ['root-extra', m => m.$type = 'System.IO.FileInfo'],
    ['missing-field', m => delete m.unit], ['wrong-unit', m => m.unit = 'px'], ['plot-zero', m => m.plot.width = 0], ['plot-extra', m => m.plot.dpi = 96],
    ['guid-empty', m => m.chartId = '00000000-0000-0000-0000-000000000000'], ['guid-uppercase', m => m.chartId = m.chartId.toUpperCase()], ['guid-newline', m => m.chartId += '\n'],
    ['script-node', m => m.semanticScene[0].geometry.type = 'script'], ['geometry-extra', m => m.semanticScene[0].geometry.onclick = 'run'], ['numeric-string', m => m.semanticScene[0].geometry.x = '57'],
    ['negative-width', m => m.semanticScene[0].geometry.width = -1], ['coordinate-bound', m => m.semanticScene[0].geometry.x = 10000001], ['node-extra', m => m.semanticScene[0].locked = true], ['node-newline', m => m.semanticScene[0].id += '\n'],
    ['null-paint', m => m.semanticScene[0].paint = null], ['paint-extra', m => m.semanticScene[0].paint.url = 'https://example.com'], ['color-url', m => m.semanticScene[0].paint.fill = 'url(https://example.com)'],
    ['color-newline', m => m.semanticScene[0].paint.fill += '\n'], ['opacity-range', m => m.semanticScene[0].paint.opacity = 1.01], ['stroke-range', m => m.semanticScene[0].paint.strokeWidth = -1],
    ['polygon-two-points', m => m.semanticScene[3].geometry.points.pop()], ['polyline-one-point', m => m.semanticScene[2].geometry.points.pop()], ['point-extra', m => m.semanticScene[2].geometry.points[0].z = 0],
    ['sweep-zero', m => m.semanticScene[4].geometry.sweep = 0], ['sweep-range', m => m.semanticScene[4].geometry.sweep = 361], ['angle-range', m => m.semanticScene[4].geometry.startAngle = -361],
    ['radius-negative', m => m.semanticScene[4].geometry.radius = -1], ['text-null', m => m.semanticScene[5].geometry.text = null], ['text-long', m => m.semanticScene[5].geometry.text = '😀'.repeat(4097)],
    ['text-surrogate', m => m.semanticScene[5].geometry.text = '\ud800'], ['font-missing', m => delete m.semanticScene[5].geometry.fontFamily], ['font-whitespace', m => m.semanticScene[5].geometry.fontFamily = '\ufeff'],
    ['font-size', m => m.semanticScene[5].geometry.fontSize = 0], ['alignment', m => m.semanticScene[5].geometry.alignment = 'justify'], ['diagnostic-code', m => m.diagnostics[0].code = 'bad'],
    ['diagnostic-newline', m => m.diagnostics[0].code += '\n'], ['too-many-path-points', m => m.semanticScene[2].geometry.points = Array.from({ length: 10001 }, () => ({ x: 0, y: 0 }))],
    ['too-many-diagnostics', m => m.diagnostics = Array.from({ length: 10001 }, () => ({ code: 'TEST', nodeId: null }))]
  ]) mutate(name, update);
  mutate('duplicate-node-id', m => m.semanticScene.push(structuredClone(m.semanticScene[0])), false, true);
  mutate('cross-chart-node', m => m.semanticScene[0].id = m.semanticScene[0].id.replace(chartId, '12345678-1234-1234-1234-123456789abc'), false, true);
  mutate('diagnostic-dangling', m => m.diagnostics[0].nodeId = chartId + '/missing', false, true);
  mutate('text-unicode-length-boundary', m => m.semanticScene[5].geometry.text = '😀'.repeat(4096), true, true);
  mutate('font-ecmascript-nonspace', m => m.semanticScene[5].geometry.fontFamily = '\u0085', true, true);
  mutate('sector-full-circle', m => m.semanticScene[4].geometry.sweep = 360, true, true);
  mutate('positive-subnormal', m => m.semanticScene[4].geometry.sweep = Number.MIN_VALUE, true, true);
  mutate('revision-integer-limit', m => m.modelRevision = 2147483647, true, true);
  const wire = JSON.stringify(golden);
  for (const [name, json] of [
    ['duplicate-root-field', wire.replace('"schemaVersion":1', '"schemaVersion":1,"schemaVersion":1')],
    ['duplicate-nested-field', wire.replace('"x":57', '"x":57,"x":57')],
    ['single-quotes', wire.replace('"schemaVersion"', "'schemaVersion'")], ['unquoted-key', wire.replace('"schemaVersion"', 'schemaVersion')],
    ['root-trailing-comma', wire.slice(0, -1) + ',}'], ['array-trailing-comma', wire.replace('"semanticScene":[', '"semanticScene":[,')],
    ['comment', '/*comment*/' + wire], ['trailing-document', wire + '{}'], ['unicode-whitespace', '\u00a0' + wire],
    ['leading-zero', wire.replace('"modelRevision":7', '"modelRevision":07')], ['plus-number', wire.replace('"modelRevision":7', '"modelRevision":+7')],
    ['nan-token', wire.replace('"modelRevision":7', '"modelRevision":NaN')], ['infinity-token', wire.replace('"modelRevision":7', '"modelRevision":Infinity')],
    ['overflow-token', wire.replace('"modelRevision":7', '"modelRevision":1e999')], ['invalid-escape', wire.replace('Revenue', 'Revenue\\q')],
    ['lone-surrogate-escape', wire.replace('Revenue', '\\ud800')], ['unescaped-control', wire.replace('Revenue', 'Rev\u0001enue')]
  ]) cases.push({ name, valid: false, schemaValid: false, json });
  for (const [name, update] of [['hash-mismatch', m => m.modelRevision++], ['hash-newline', m => m.renderHash += '\n'], ['hash-uppercase', m => m.renderHash = m.renderHash.toUpperCase()]]) {
    const model = structuredClone(golden); update(model); cases.push({ name, valid: false, schemaValid: name === 'hash-mismatch', json: JSON.stringify(model) });
  }
  let seed = 0x42cafe;
  const random = () => { seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0; return seed / 2 ** 32; };
  for (let i = 0; i < 60; i++) mutate('finite-double-' + i, m => {
    for (const node of m.semanticScene.slice(0, 2)) { node.geometry.x = (random() - .5) * 2e7; node.geometry.y = (random() - .5) * 10 ** (-Math.floor(random() * 300)); }
  }, true, true);
  const base = { chartId, modelRevision: 5, dataset: { categories: [{ id: 'a', label: 'A' }, { id: 'b', label: 'B' }, { id: 'c', label: 'C' }], series: [{ id: 's', name: 'S', values: [20, null, 50] }] }, plot: { x: 40, y: 60, width: 600, height: 300, unit: 'pt' }, palette: ['#3b82f6'] };
  for (const kind of ['clustered', 'stacked', 'positive-100', 'line', 'area', 'pie']) {
    const options = kind === 'pie' ? { startAngle: 360, palette: ['#ff0000', '#00ff00', '#0000ff'] } : ['line', 'area'].includes(kind) ? { gapPolicy: 'break' } : { orientation: 'horizontal', gapRatio: .2 };
    const scene = await buildScene({ ...base, kind, ...options });
    if (referenceHash(scene) !== scene.renderHash) throw Error('Builder hash mismatch');
    cases.push({ name: 'engine-' + kind, valid: true, schemaValid: true, json: JSON.stringify(scene) });
  }
  const canonical = await renderChartV2(readFileSync(new URL('../s1/fixtures/clustered-column.v2.json', import.meta.url), 'utf8'));
  cases.push({ name: 'canonical-v2-projection', valid: true, schemaValid: true, json: JSON.stringify(canonical) });
  return cases;
}
if (process.argv[2] === '--write-golden') {
  mkdirSync(new URL('./fixtures/', import.meta.url), { recursive: true });
  writeFileSync(new URL('./fixtures/all-primitives.v1.json', import.meta.url), JSON.stringify({ ...primitives, renderHash: referenceHash(primitives) }, null, 2) + '\n');
} else if (process.argv[2] === '--write-corpus') {
  mkdirSync(process.argv[3], { recursive: true });
  const cases = await createCorpus();
  writeFileSync(join(process.argv[3], 'cases.json'), JSON.stringify(cases));
  console.log(`Prepared ${cases.length} synthetic scene contract fixtures`);
}
