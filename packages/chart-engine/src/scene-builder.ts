import { clustered, stacked, positive100, line, area, pie, validateDataset, CalculationError } from './basic-charts.ts';
import type { Dataset, Plot } from './basic-charts.ts';
import { sealScene, validateScene } from './scene-contract.ts';
import type { ScenePayload, SceneNode, Geometry, Paint, RenderResult, SceneDiagnostic } from './scene-contract.ts';

interface BaseRequest { chartId: string; modelRevision: number; dataset: Dataset; plot: Plot; palette: string[] }
export type SceneRequest = BaseRequest & (
  | { kind: 'clustered' | 'stacked' | 'positive-100'; orientation: 'vertical' | 'horizontal'; gapRatio: number }
  | { kind: 'line' | 'area'; gapPolicy: 'break' | 'connect' | 'zero' }
  | { kind: 'pie'; startAngle: number }
);
/** Normalized-data adapter; persisted ChartDocument v2 is not implicitly migrated. */
export async function buildScene(request: SceneRequest): Promise<RenderResult> {
  if (!request || typeof request !== 'object') throw new CalculationError('scene.request');
  const extra = ['clustered', 'stacked', 'positive-100'].includes(request.kind) ? ['orientation', 'gapRatio']
    : ['line', 'area'].includes(request.kind) ? ['gapPolicy'] : request.kind === 'pie' ? ['startAngle'] : [];
  const keys = ['chartId', 'modelRevision', 'dataset', 'plot', 'palette', 'kind', ...extra];
  if (!extra.length || Object.keys(request).length !== keys.length || !keys.every(k => Object.hasOwn(request, k))) throw new CalculationError('scene.request-fields');
  validateDataset(request.dataset);
  const expectedColors = request.kind === 'pie' ? request.dataset.categories.length : request.dataset.series.length;
  if (!Array.isArray(request.palette) || request.palette.length !== expectedColors || !request.palette.every(c => typeof c === 'string' && /^#[0-9a-fA-F]{6}$/.test(c))) throw new CalculationError('scene.palette');
  const nodes: SceneNode[] = [], diagnostics: SceneDiagnostic[] = [];
  const paint = (index: number, stroke = false): Paint => ({ fill: stroke ? null : request.palette[index]!, stroke: stroke ? request.palette[index]! : null, strokeWidth: stroke ? 2 : 0, opacity: 1 });
  const add = (id: string, geometry: Geometry, style: Paint) => nodes.push({ id: `${request.chartId}/${id}`, geometry, paint: style });
  if (request.kind === 'clustered' || request.kind === 'stacked' || request.kind === 'positive-100') {
    const scene = request.kind === 'clustered' ? { ...clustered(request.dataset, request.plot, request.orientation, request.gapRatio), diagnostics: [] }
      : request.kind === 'stacked' ? stacked(request.dataset, request.plot, request.orientation, request.gapRatio)
      : positive100(request.dataset, request.plot, request.orientation, request.gapRatio);
    for (const bar of scene.bars) add(bar.id, { type: 'rectangle', x: bar.x, y: bar.y, width: bar.width, height: bar.height }, paint(bar.seriesIndex));
    for (const d of scene.diagnostics) diagnostics.push({ code: d.code, nodeId: null });
  } else if (request.kind === 'line' || request.kind === 'area') {
    const scene = line(request.dataset, request.plot, { gapPolicy: request.gapPolicy, includeZero: true });
    const polygons = request.kind === 'area' ? area(request.dataset, request.plot, { gapPolicy: request.gapPolicy, baseline: 0 }).polygons : [];
    for (const [si, path] of scene.paths.entries()) {
      const seriesPolygons = polygons.filter(p => p.seriesId === path.seriesId);
      for (const [segmentIndex, segment] of path.segments.entries()) {
        const first = segment[0]!, last = segment[segment.length - 1]!;
        if (request.kind === 'area') {
          // Build polygons from the exact same calculation policy, preserving gaps.
          const polygon = seriesPolygons[segmentIndex];
          if (!polygon) throw new CalculationError('scene.area-mapping');
          add(`area:${path.seriesId}/${first.categoryId}/${last.categoryId}`, { type: 'polygon', points: polygon.points }, { ...paint(si), opacity: 0.35 });
        }
        if (segment.length > 1) add(`line:${path.seriesId}/${first.categoryId}/${last.categoryId}`, { type: 'polyline', points: segment.map(p => ({ x: p.x, y: p.y })) }, paint(si, true));
      }
      for (const p of path.points) add(p.id, { type: 'ellipse', x: p.x - 3, y: p.y - 3, width: 6, height: 6 }, paint(si));
    }
  } else if (request.kind === 'pie') {
    const scene = pie(request.dataset, request.plot, request.startAngle);
    for (const sector of scene.sectors) {
      const index = request.dataset.categories.findIndex(c => c.id === sector.categoryId);
      if (sector.sweep === 0) continue; // IEEE-754 precision loss is reported below.
      add(sector.id, { type: 'sector', cx: scene.center.x, cy: scene.center.y, radius: scene.radius, startAngle: ((sector.startAngle % 360) + 360) % 360, sweep: sector.sweep }, paint(index));
    }
    for (const d of scene.diagnostics) diagnostics.push({ code: d.code, nodeId: null });
  } else throw new CalculationError('scene.kind');
  const payload: ScenePayload = { schemaVersion: 1, chartId: request.chartId, modelRevision: request.modelRevision, unit: 'pt', plot: { ...request.plot }, semanticScene: nodes, diagnostics };
  validateScene(payload, false);
  return sealScene(payload);
}
