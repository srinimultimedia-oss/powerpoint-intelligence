import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { validateChart } from '../src/validate.mjs';

const golden = JSON.parse(readFileSync(new URL('../../../tests/s1/fixtures/clustered-column.v2.json', import.meta.url), 'utf8'));
test('golden document passes schema and semantic validation without mutation', () => {
  const input = structuredClone(golden);
  assert.deepEqual(validateChart(input), { valid: true, codes: [] });
  assert.deepEqual(input, golden);
  assert.deepEqual(JSON.parse(JSON.stringify(input)), golden);
});
const invalid = {
  'unknown schema version': m => { m.schemaVersion = 99; },
  'missing field': m => { delete m.revision; },
  'unknown field': m => { m.secret = 'unsupported'; },
  'unknown nested field': m => { m.layout.rotation = 90; },
  'coerced revision': m => { m.revision = '4'; },
  'negative revision': m => { m.revision = -1; },
  'overflow revision': m => { m.revision = 2147483648; },
  'bad chart ID': m => { m.chartId = 'bad'; },
  'empty chart GUID': m => { m.chartId = '00000000-0000-0000-0000-000000000000'; },
  'unsupported chart': m => { m.chartType = 'pie'; },
  'duplicate semantic ID': m => { m.series[0].id = m.dataset.categories[0].id; },
  'dimensions': m => { m.series[0].values.pop(); },
  'empty categories': m => { m.dataset.categories = []; },
  'empty series': m => { m.series = []; },
  'nonfinite data': m => { m.series[0].values[0] = Infinity; },
  'nonfinite layout': m => { m.layout.width = NaN; },
  'zero width': m => { m.layout.width = 0; },
  'non-point units': m => { m.layout.unit = 'px'; },
  'duplicate axis role': m => { m.axes[1].role = 'category'; m.axes[1].scale = 'category'; },
  'wrong axis scale': m => { m.axes[1].scale = 'category'; },
  'missing color': m => { m.style.seriesColors.pop(); },
  'invalid color': m => { m.style.seriesColors[0] = 'red'; },
  'blank label': m => { m.dataset.categories[0].label = ' '; }
};
for (const [name, mutate] of Object.entries(invalid)) test(`reject ${name}`, () => {
  const input = structuredClone(golden); mutate(input);
  assert.equal(validateChart(input).valid, false);
});
