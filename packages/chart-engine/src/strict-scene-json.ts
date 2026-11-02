import { CalculationError } from './basic-charts.ts';

export const MAX_SCENE_CHARACTERS = 8388608;
/** JSON.parse alone discards duplicate object members. Reject them before use. */
export function parseSceneJson(json: string): unknown {
  const fail = (): never => { throw new CalculationError('scene.invalid-json'); };
  if (typeof json !== 'string' || json.length === 0 || json.length > MAX_SCENE_CHARACTERS) throw new CalculationError('scene.json-size');
  let position = 0;
  const space = () => { while (/[ \t\r\n]/.test(json[position] ?? '') && position < json.length) position++; };
  const string = (): string => {
    const start = position++;
    while (position < json.length) {
      const c = json[position++]!;
      if (c === '"') {
        let value: string;
        try { value = JSON.parse(json.slice(start, position)) as string; } catch { return fail(); }
        for (const char of value) { const n = char.codePointAt(0)!; if (n >= 0xd800 && n <= 0xdfff) fail(); }
        return value;
      }
      if (c === '\\') position++; // JSON.parse checks the exact escape grammar.
    }
    return fail();
  };
  const value = (depth: number): unknown => {
    space(); if (depth > 32 || position >= json.length) fail();
    const c = json[position];
    if (c === '"') return string();
    if (c === '{' || c === '[') {
      position++; space(); const object = c === '{', end = object ? '}' : ']';
      const entries: Record<string, unknown> = Object.create(null) as Record<string, unknown>, items: unknown[] = [];
      if (json[position] === end) { position++; return object ? entries : items; }
      while (position < json.length) {
        if (object) {
          if (json[position] !== '"') fail();
          const key = string(); space(); if (json[position++] !== ':' || Object.hasOwn(entries, key)) fail();
          entries[key] = value(depth + 1);
        } else items.push(value(depth + 1));
        space(); const delimiter = json[position++];
        if (delimiter === end) return object ? entries : items;
        if (delimiter !== ',') fail();
        space(); // A closing delimiter here is not a value: no trailing commas.
      }
      return fail();
    }
    for (const [literal, result] of [['true', true], ['false', false], ['null', null]] as const) {
      if (json.startsWith(literal, position)) { position += literal.length; return result; }
    }
    const match = /^-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?/.exec(json.slice(position));
    if (!match) return fail();
    position += match[0].length; const number = Number(match[0]); if (!Number.isFinite(number)) fail();
    return number;
  };
  const result = value(1); space(); if (position !== json.length) fail();
  return result;
}
