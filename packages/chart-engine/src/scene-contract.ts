/** Versioned renderer boundary. No COM, file, network, or arbitrary executable nodes. */
import { CalculationError, validatePlot } from './basic-charts.ts';
import type { Plot } from './basic-charts.ts';
import { parseSceneJson, MAX_SCENE_CHARACTERS } from './strict-scene-json.ts';

export interface Point { x: number; y: number }
export type Geometry =
  | { type: 'rectangle' | 'ellipse'; x: number; y: number; width: number; height: number }
  | { type: 'polyline' | 'polygon'; points: Point[] }
  | { type: 'sector'; cx: number; cy: number; radius: number; startAngle: number; sweep: number }
  | { type: 'text'; x: number; y: number; width: number; height: number; text: string; fontFamily: string; fontSize: number; alignment: 'left' | 'center' | 'right' };
export interface Paint { fill: string | null; stroke: string | null; strokeWidth: number; opacity: number }
export interface SceneNode { id: string; geometry: Geometry; paint: Paint }
export interface SceneDiagnostic { code: string; nodeId: string | null }
export interface ScenePayload {
  schemaVersion: 1; chartId: string; modelRevision: number; unit: 'pt';
  plot: Plot; semanticScene: SceneNode[]; diagnostics: SceneDiagnostic[];
}
export interface RenderResult extends ScenePayload { renderHash: string }
const guid = /^(?!00000000-0000-0000-0000-000000000000$)[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$(?![\s\S])/;
const localId = /^[A-Za-z0-9][A-Za-z0-9:._-]{0,159}(?:\/[A-Za-z0-9][A-Za-z0-9:._-]{0,159}){0,3}$(?![\s\S])/;
const color = /^#[0-9a-fA-F]{6}$(?![\s\S])/;
function requireValue(condition: unknown, code: string): asserts condition {
  if (!condition) throw new CalculationError(code);
}
function record(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}
function fields(value: unknown, names: string[]): asserts value is Record<string, unknown> {
  requireValue(record(value) && Object.keys(value).length === names.length && names.every(n => Object.hasOwn(value, n)), 'scene.fields');
}
function numeric(value: unknown, minimum = -10000000, maximum = 10000000): void {
  requireValue(typeof value === 'number' && Number.isFinite(value) && value >= minimum && value <= maximum, 'scene.number');
}
function unicode(value: unknown, maximum: number): asserts value is string {
  requireValue(typeof value === 'string' && [...value].length <= maximum, 'scene.text');
  // Reject unpaired surrogates; all implementations hash the same valid UTF-8.
  for (const character of value) {
    const code = character.codePointAt(0)!;
    requireValue(code < 0xd800 || code > 0xdfff, 'scene.unicode');
  }
}
export function validateScene(value: unknown, sealed?: true): asserts value is RenderResult;
export function validateScene(value: unknown, sealed: false): asserts value is ScenePayload;
export function validateScene(value: unknown, sealed = true): asserts value is ScenePayload {
  fields(value, ['schemaVersion', 'chartId', 'modelRevision', 'unit', 'plot', 'semanticScene', 'diagnostics', ...(sealed ? ['renderHash'] : [])]);
  requireValue(value.schemaVersion === 1 && value.unit === 'pt', 'scene.version-or-unit');
  requireValue(typeof value.chartId === 'string' && guid.test(value.chartId), 'scene.chart-id');
  requireValue(typeof value.modelRevision === 'number' && Number.isInteger(value.modelRevision) && value.modelRevision >= 0 && value.modelRevision <= 2147483647, 'scene.revision');
  validatePlot(value.plot);
  requireValue(Array.isArray(value.semanticScene) && value.semanticScene.length <= 200000, 'scene.semanticScene');
  const ids = new Set<string>();
  for (const node of value.semanticScene) {
    fields(node, ['id', 'geometry', 'paint']);
    requireValue(typeof node.id === 'string' && node.id.startsWith(`${value.chartId}/`) && localId.test(node.id.slice(37)) && !ids.has(node.id), 'scene.node-id');
    ids.add(node.id);
    const g = node.geometry;
    requireValue(record(g), 'scene.geometry');
    if (g.type === 'rectangle' || g.type === 'ellipse' || g.type === 'text') {
      fields(g, ['type', 'x', 'y', 'width', 'height', ...(g.type === 'text' ? ['text', 'fontFamily', 'fontSize', 'alignment'] : [])]);
      numeric(g.x); numeric(g.y); numeric(g.width, 0); numeric(g.height, 0);
      if (g.type === 'text') {
        unicode(g.text, 4096); unicode(g.fontFamily, 128);
        requireValue(/\S/u.test(g.fontFamily as string), 'scene.font');
        numeric(g.fontSize, 1, 1000);
        requireValue(['left', 'center', 'right'].includes(g.alignment as string), 'scene.alignment');
      }
    } else if (g.type === 'polyline' || g.type === 'polygon') {
      fields(g, ['type', 'points']);
      requireValue(Array.isArray(g.points) && g.points.length >= (g.type === 'polygon' ? 3 : 2) && g.points.length <= 10000, 'scene.points');
      for (const point of g.points) { fields(point, ['x', 'y']); numeric(point.x); numeric(point.y); }
    } else if (g.type === 'sector') {
      fields(g, ['type', 'cx', 'cy', 'radius', 'startAngle', 'sweep']);
      numeric(g.cx); numeric(g.cy); numeric(g.radius, 0); numeric(g.startAngle, -360, 360); numeric(g.sweep, Number.MIN_VALUE, 360);
    } else throw new CalculationError('scene.geometry-type');
    fields(node.paint, ['fill', 'stroke', 'strokeWidth', 'opacity']);
    for (const name of ['fill', 'stroke']) requireValue(node.paint[name] === null || (typeof node.paint[name] === 'string' && color.test(node.paint[name] as string)), 'scene.color');
    numeric(node.paint.strokeWidth, 0, 1000); numeric(node.paint.opacity, 0, 1);
  }
  requireValue(Array.isArray(value.diagnostics) && value.diagnostics.length <= 10000, 'scene.diagnostics');
  for (const diagnostic of value.diagnostics) {
    fields(diagnostic, ['code', 'nodeId']);
    requireValue(typeof diagnostic.code === 'string' && /^[A-Z][A-Z0-9_]{0,63}$(?![\s\S])/.test(diagnostic.code), 'scene.diagnostic-code');
    requireValue(diagnostic.nodeId === null || (typeof diagnostic.nodeId === 'string' && ids.has(diagnostic.nodeId)), 'scene.diagnostic-node');
  }
  if (sealed) requireValue(typeof value.renderHash === 'string' && /^[0-9a-f]{64}$(?![\s\S])/.test(value.renderHash), 'scene.hash-format');
}

/** PPI-C14N-1: tagged values, big-endian lengths/doubles, ordinal object keys.
 * Hashes geometry exactly, without JSON number-format differences or rounding.
 */
export function sceneHashBytes(payload: ScenePayload): Uint8Array {
  validateScene(payload, false);
  const result: number[] = [];
  const length = (n: number) => { result.push((n >>> 24) & 255, (n >>> 16) & 255, (n >>> 8) & 255, n & 255); };
  const utf8 = (s: string) => {
    const bytes: number[] = [];
    for (const char of s) {
      const n = char.codePointAt(0)!;
      if (n < 128) bytes.push(n);
      else if (n < 2048) bytes.push(192 | (n >> 6), 128 | (n & 63));
      else if (n < 65536) bytes.push(224 | (n >> 12), 128 | ((n >> 6) & 63), 128 | (n & 63));
      else bytes.push(240 | (n >> 18), 128 | ((n >> 12) & 63), 128 | ((n >> 6) & 63), 128 | (n & 63));
    }
    length(bytes.length); for (const b of bytes) result.push(b);
  };
  const write = (value: unknown): void => {
    if (value === null) result.push(0);
    else if (typeof value === 'boolean') result.push(value ? 2 : 1);
    else if (typeof value === 'number') {
      result.push(3); const buffer = new ArrayBuffer(8); new DataView(buffer).setFloat64(0, value === 0 ? 0 : value, false);
      for (const b of new Uint8Array(buffer)) result.push(b);
    } else if (typeof value === 'string') { result.push(4); utf8(value); }
    else if (Array.isArray(value)) { result.push(5); length(value.length); for (const item of value) write(item); }
    else if (record(value)) {
      const keys = Object.keys(value).sort(); result.push(6); length(keys.length);
      for (const key of keys) { write(key); write(value[key]); }
    } else throw new CalculationError('scene.hash-value');
  };
  // Domain separation prevents the same byte stream identifying another contract.
  utf8('PPI-SCENE-1'); write(payload);
  return Uint8Array.from(result);
}
async function digest(payload: ScenePayload): Promise<string> {
  const webCrypto = (globalThis as unknown as { crypto?: { subtle?: { digest(name: string, input: Uint8Array): Promise<ArrayBuffer> } } }).crypto;
  requireValue(webCrypto?.subtle, 'scene.sha256-unavailable');
  const bytes = new Uint8Array(await webCrypto.subtle.digest('SHA-256', sceneHashBytes(payload)));
  return Array.from(bytes, b => b.toString(16).padStart(2, '0')).join('');
}
export async function sealScene(payload: ScenePayload): Promise<RenderResult> {
  // Snapshot before awaiting WebCrypto so caller mutation cannot race the hash.
  validateScene(payload, false);
  const copy = JSON.parse(JSON.stringify(payload)) as ScenePayload;
  return { ...copy, renderHash: await digest(copy) };
}
export async function verifyScene(value: unknown): Promise<RenderResult> {
  validateScene(value);
  const copy = JSON.parse(JSON.stringify(value)) as RenderResult;
  const { renderHash, ...payload } = copy;
  requireValue(await digest(payload) === renderHash, 'scene.hash-mismatch');
  return copy;
}
export async function decodeScene(json: string): Promise<RenderResult> { return verifyScene(parseSceneJson(json)); }
export async function encodeScene(scene: RenderResult): Promise<string> {
  const json = JSON.stringify(await verifyScene(scene));
  requireValue(json.length <= MAX_SCENE_CHARACTERS, 'scene.json-size');
  return json;
}
