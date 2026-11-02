import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { buildScene } from '../src/scene-builder.ts';
import { sealScene, verifyScene, validateScene, sceneHashBytes } from '../src/scene-contract.ts';
const chartId = '7bef06ad-f023-4b26-9a32-0d51b757bf59';
const dataset = { categories: [{ id: 'a', label: 'A' }, { id: 'b', label: 'B' }, { id: 'c', label: 'C' }], series: [{ id: 's', name: 'Series', values: [1, null, 3] }] };
const base = { chartId, modelRevision: 4, dataset, plot: { x: 0, y: 0, width: 600, height: 300, unit: 'pt' }, palette: ['#3b82f6'] };
const request = { ...base, kind: 'clustered', orientation: 'vertical', gapRatio: .2 };
test('sealed scene retains chart identity, revision and stable scoped node IDs', async () => {
  const scene = await buildScene(request);
  assert.equal(scene.schemaVersion, 1); assert.equal(scene.modelRevision, 4);
  assert.equal(scene.semanticScene.length, 2);
  assert.deepEqual(scene.semanticScene.map(n => n.id), [chartId + '/bar:s/a', chartId + '/bar:s/c']);
  assert.deepEqual(await verifyScene(scene), scene);
  const { renderHash, ...payload } = scene;
  assert.equal(createHash('sha256').update(sceneHashBytes(payload)).digest('hex'), renderHash);
  assert.deepEqual(scene, await buildScene(request));
});
test('every basic engine maps to a closed geometry type with explicit color', async () => {
  for (const kind of ['clustered', 'stacked', 'positive-100', 'line', 'area', 'pie']) {
    const options = ['line', 'area'].includes(kind) ? { gapPolicy: 'break' } : kind === 'pie' ? { startAngle: 360 } : { orientation: 'horizontal', gapRatio: .2 };
    const scene = await buildScene({ ...base, kind, ...options, palette: kind === 'pie' ? ['#ff0000', '#00ff00', '#0000ff'] : base.palette });
    assert.ok(scene.semanticScene.length > 0); await verifyScene(scene);
    assert.ok(scene.semanticScene.every(n => ['rectangle', 'ellipse', 'polyline', 'polygon', 'sector'].includes(n.geometry.type)));
  }
});
test('revision, geometry, color, draw order and diagnostics are hash-bound', async () => {
  const scene = await buildScene(request);
  for (const mutate of [s => s.modelRevision++, s => s.semanticScene[0].geometry.x++, s => s.semanticScene[0].paint.fill = '#ff0000', s => s.semanticScene.reverse(), s => s.diagnostics.push({ code: 'TEST', nodeId: null })]) {
    const changed = structuredClone(scene); mutate(changed);
    await assert.rejects(verifyScene(changed), /scene.hash-mismatch/);
  }
});
test('object insertion order and negative zero do not change the canonical hash', async () => {
  const { renderHash, ...payload } = await buildScene(request);
  const reordered = Object.fromEntries(Object.entries(payload).reverse());
  reordered.plot.x = -0;
  assert.equal((await sealScene(reordered)).renderHash, renderHash);
});
test('seal snapshots inputs before asynchronous hashing', async () => {
  const { renderHash, ...payload } = await buildScene(request);
  const pending = sealScene(payload); payload.semanticScene[0].geometry.x = 999;
  assert.equal((await pending).renderHash, renderHash);
});
test('unknown members, executable nodes, duplicate IDs and wrong chart scope reject', async () => {
  const original = await buildScene(request);
  for (const mutate of [s => s.extra = true, s => s.semanticScene[0].geometry.type = 'script', s => s.semanticScene[0].geometry.x = '0', s => s.semanticScene.push(structuredClone(s.semanticScene[0])), s => s.semanticScene[0].id = 'other/bar:s/a', s => s.diagnostics.push({ code: 'BAD', nodeId: 'missing' }), s => s.semanticScene[0].paint.fill = 'url(https://example.com)', s => s.semanticScene[0].geometry.width = Infinity]) {
    const bad = structuredClone(original); mutate(bad);
    assert.throws(() => validateScene(bad));
  }
});
test('line gaps do not create connecting segments; isolated values still have markers', async () => {
  const scene = await buildScene({ ...base, kind: 'line', gapPolicy: 'break' });
  assert.equal(scene.semanticScene.filter(n => n.geometry.type === 'polyline').length, 0);
  assert.equal(scene.semanticScene.filter(n => n.geometry.type === 'ellipse').length, 2);
  const connected = await buildScene({ ...base, kind: 'line', gapPolicy: 'connect' });
  assert.equal(connected.semanticScene.filter(n => n.geometry.type === 'polyline').length, 1);
});
test('text nodes retain Unicode but reject lone surrogates and missing font choice', async () => {
  const { renderHash, ...payload } = await buildScene(request);
  payload.semanticScene = [{ id: chartId + '/label:a', geometry: { type: 'text', x: 0, y: 0, width: 100, height: 20, text: 'Δ 😀', fontFamily: 'Segoe UI', fontSize: 12, alignment: 'left' }, paint: { fill: '#000000', stroke: null, strokeWidth: 0, opacity: 1 } }];
  await verifyScene(await sealScene(payload));
  payload.semanticScene[0].geometry.text = '\uD800';
  await assert.rejects(sealScene(payload), /scene.unicode/);
  payload.semanticScene[0].geometry.text = 'ok'; payload.semanticScene[0].geometry.fontFamily = '';
  await assert.rejects(sealScene(payload), /scene.font/);
});
test('palette count and input fields reject ambiguity', async () => {
  await assert.rejects(buildScene({ ...request, palette: [] }), /scene.palette/);
  await assert.rejects(buildScene({ ...request, arbitrary: true }), /scene.request-fields/);
});
