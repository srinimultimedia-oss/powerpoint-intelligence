import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import Ajv2020 from 'ajv/dist/2020.js';
import { createCorpus, referenceHash } from '../../../tests/s3/scene-fixtures.mjs';
import { decodeScene, encodeScene } from '../src/scene-contract.ts';
import { parseSceneJson, MAX_SCENE_CHARACTERS } from '../src/strict-scene-json.ts';
const schema = JSON.parse(readFileSync(new URL('../schema/render-result.v1.schema.json', import.meta.url), 'utf8'));
const validate = new Ajv2020({ strict: true, allErrors: true }).compile(schema);
const cases = await createCorpus();
for (const fixture of cases) test('scene schema/semantic/strict JSON: ' + fixture.name, async () => {
  let model, structural = false;
  try { model = parseSceneJson(fixture.json); structural = validate(model); } catch { /* Syntax rejects before schema. */ }
  assert.equal(structural, fixture.schemaValid, fixture.name + ': schema expectation');
  if (fixture.valid) {
    const decoded = await decodeScene(fixture.json);
    assert.equal(decoded.renderHash, referenceHash(decoded));
    assert.deepEqual(await decodeScene(await encodeScene(decoded)), decoded);
  } else await assert.rejects(decodeScene(fixture.json));
});
test('strict parser rejects excessive size and nesting before DTO conversion', () => {
  assert.throws(() => parseSceneJson(' '.repeat(MAX_SCENE_CHARACTERS + 1)), /scene.json-size/);
  assert.throws(() => parseSceneJson('['.repeat(33) + '0' + ']'.repeat(33)), /scene.invalid-json/);
});
test('strict parser accepts JSON escapes and keeps prototype keys as own data', () => {
  const model = parseSceneJson('{"__proto__":1,"text":"\\uD83D\\uDE00\\n\\b\\f\\r\\t\\/\\\\\\\""}');
  assert.equal(Object.getPrototypeOf(model), null);
  assert.equal(model.__proto__, 1);
  assert.equal(model.text, '😀\n\b\f\r\t/\\"');
});
