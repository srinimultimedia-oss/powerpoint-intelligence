import test from 'node:test';
import assert from 'node:assert/strict';
import { stackedTotals, normalizePositive100Stack, waterfallSegments, cagr } from '../src/calculations.mjs';

test('stack totals', () => assert.deepEqual(stackedTotals([[20, 5], [30, 15], [50, 30]]), [100, 50]));
test('100 percent stack', () => assert.deepEqual(normalizePositive100Stack([[20], [30], [50]]), [[20], [30], [50]]));
test('zero total returns null segments', () => assert.deepEqual(normalizePositive100Stack([[0], [0]]), [[null], [null]]));
test('negative 100% input is rejected', () => assert.throws(() => normalizePositive100Stack([[-1], [2]])));
test('waterfall total is not double counted', () => {
 const x = waterfallSegments([{value:100},{value:40},{value:-25},{value:0,role:'total'}]);
 assert.equal(x.at(-1).running, 115);
 assert.equal(x.at(-1).high, 115);
});
test('CAGR is ten percent', () => assert.ok(Math.abs(cagr(100,121,2)-0.1)<1e-12));
