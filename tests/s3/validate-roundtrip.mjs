import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import Ajv2020 from '../../packages/chart-engine/node_modules/ajv/dist/2020.js';
import { decodeScene } from '../../packages/chart-engine/src/scene-contract.ts';
import { referenceHash } from './scene-fixtures.mjs';
const schema = JSON.parse(readFileSync(new URL('../../packages/chart-engine/schema/render-result.v1.schema.json', import.meta.url), 'utf8'));
const validate = new Ajv2020({ strict: true, allErrors: true }).compile(schema);
const fixtures = JSON.parse(readFileSync(join(process.argv[2], 'cases.json'), 'utf8'));
const output = process.argv[3]; let passed = 0;
for (const fixture of fixtures.filter(f => f.valid)) {
  const original = await decodeScene(fixture.json), roundtrip = await decodeScene(readFileSync(join(output, fixture.name + '.v1.json'), 'utf8'));
  assert.equal(validate(roundtrip), true, fixture.name + ': schema');
  assert.deepEqual(roundtrip, original, fixture.name + ': data changed');
  assert.equal(roundtrip.renderHash, referenceHash(roundtrip)); passed++;
}
const produced = await decodeScene(readFileSync(join(output, 'csharp-produced.v1.json'), 'utf8'));
assert.equal(validate(produced), true); assert.equal(produced.modelRevision, 8);
assert.equal(produced.semanticScene[1].geometry.y, -1.2345678901234567e-200);
assert.equal(produced.semanticScene[5].geometry.text, 'مرحبا 😀 "quoted" <&>');
assert.equal(produced.semanticScene[4].geometry.sweep, Number.MIN_VALUE);
assert.equal(produced.renderHash, referenceHash(produced));
console.log(`PASS ${passed} TypeScript → .NET 4.8 → TypeScript scene round trips and independent C# producer/hash`);
