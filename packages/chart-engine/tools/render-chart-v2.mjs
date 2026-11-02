// Local Node worker/CLI boundary. Files and AJV stay outside the pure engine.
import { readFileSync, writeFileSync } from 'node:fs';
import { pathToFileURL } from 'node:url';
import { validateChart } from '../../chart-domain/src/validate.mjs';
import { buildScene } from '../src/scene-builder.ts';
import { sealScene, encodeScene } from '../src/scene-contract.ts';
import { parseSceneJson } from '../src/strict-scene-json.ts';

export async function renderChartV2(json) {
  if (typeof json !== 'string' || json.length > 1048576) throw Error('chart.json-size');
  const model = parseSceneJson(json);
  const validation = validateChart(model);
  if (!validation.valid) throw Error('chart.invalid');
  const scene = await buildScene({
    chartId: model.chartId.toLowerCase(), modelRevision: model.revision, kind: 'clustered', orientation: 'vertical', gapRatio: .2,
    dataset: { categories: model.dataset.categories, series: model.series }, plot: model.layout, palette: model.style.seriesColors
  });
  const { renderHash, ...payload } = scene;
  // Geometry-only preview must disclose currently omitted canonical presentation features.
  payload.diagnostics.push({ code: 'AXES_NOT_RENDERED', nodeId: null });
  if (model.style.showLegend) payload.diagnostics.push({ code: 'LEGEND_NOT_RENDERED', nodeId: null });
  return sealScene(payload);
}
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  try {
    if (process.argv.length !== 4) throw Error('usage');
    writeFileSync(process.argv[3], await encodeScene(await renderChartV2(readFileSync(process.argv[2], 'utf8'))) + '\n');
    console.log('Created validated geometry-only scene; axes and legends are pending.');
  } catch { console.error('Chart preview failed. Supply validated v2 input and an output path. No presentation was modified.'); process.exitCode = 1; }
}
