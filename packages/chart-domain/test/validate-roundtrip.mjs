import { readFileSync } from 'node:fs';
import assert from 'node:assert/strict';
import { validateChart } from '../src/validate.mjs';
const result = validateChart(JSON.parse(readFileSync(process.argv[2], 'utf8')));
assert.equal(result.valid, true, result.codes.join(','));
console.log('PASS C# serialized golden document validates against JSON schema and semantic rules');
