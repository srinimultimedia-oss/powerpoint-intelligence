import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { renderChartV2 } from '../tools/render-chart-v2.mjs';
import { verifyScene } from '../src/scene-contract.ts';
const golden = readFileSync(new URL('../../../tests/s1/fixtures/clustered-column.v2.json', import.meta.url), 'utf8');
test('canonical v2 projects to validated scene with exact revision, colors and semantic identities', async () => {
  const model = JSON.parse(golden), scene = await renderChartV2(golden);
  assert.equal(scene.modelRevision, model.revision); assert.equal(scene.chartId, model.chartId);
  assert.deepEqual(scene.plot, model.layout); assert.equal(scene.semanticScene.length, 5);
  assert.equal(scene.semanticScene[0].id, model.chartId + '/bar:series:actual/category:q1');
  assert.equal(scene.semanticScene[0].paint.fill, model.style.seriesColors[0]);
  assert.ok(scene.semanticScene.every(n => n.geometry.type === 'rectangle'));
  assert.deepEqual(scene.diagnostics.map(d => d.code), ['AXES_NOT_RENDERED', 'LEGEND_NOT_RENDERED']);
  await verifyScene(scene); assert.equal(golden, readFileSync(new URL('../../../tests/s1/fixtures/clustered-column.v2.json', import.meta.url), 'utf8'));
});
test('canonical adapter rejects unsupported versions, malformed metadata and duplicate fields', async () => {
  const model = JSON.parse(golden); model.schemaVersion = 99;
  await assert.rejects(renderChartV2(JSON.stringify(model)), /chart.invalid/);
  await assert.rejects(renderChartV2(golden.replace('"revision": 4', '"revision": 4,"revision": 4')), /scene.invalid-json/);
  model.schemaVersion = 2; model.series[0].values.pop();
  await assert.rejects(renderChartV2(JSON.stringify(model)), /chart.invalid/);
});
test('GUID case normalization is explicit in projection and does not modify canonical source', async () => {
  const model = JSON.parse(golden); model.chartId = model.chartId.toUpperCase(); model.style.showLegend = false;
  const input = JSON.stringify(model), scene = await renderChartV2(input);
  assert.equal(scene.chartId, model.chartId.toLowerCase()); assert.equal(JSON.stringify(model), input);
  assert.deepEqual(scene.diagnostics.map(d => d.code), ['AXES_NOT_RENDERED']);
});
