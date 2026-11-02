import Ajv2020 from 'ajv/dist/2020.js';
import { readFileSync } from 'node:fs';

const schema = JSON.parse(readFileSync(new URL('../schema/chart-document.v2.schema.json', import.meta.url), 'utf8'));
const structural = new Ajv2020({ allErrors: true, strict: true }).compile(schema);

export function validateChart(model) {
  if (!structural(model)) return { valid: false, codes: [...new Set(structural.errors.map(e => `schema.${e.keyword}`))] };
  const codes = [];
  const ids = [model.chartId, model.dataset.id, ...model.dataset.categories.map(c => c.id), ...model.series.map(s => s.id), ...model.axes.map(a => a.id)];
  if (new Set(ids).size !== ids.length) codes.push('semantic.id.invalid-or-duplicate');
  if (model.series.some(s => s.values.length !== model.dataset.categories.length)) codes.push('series.dimensions');
  if (model.series.some(s => s.values.some(v => v !== null && !Number.isFinite(v)))) codes.push('series.nonfinite');
  if (model.axes.map(a => a.role).sort().join(',') !== 'category,value') codes.push('axes.roles');
  if (model.style.seriesColors.length !== model.series.length) codes.push('style.colors');
  return { valid: codes.length === 0, codes };
}
