import { clustered, stacked, positive100, line, area, pie, validatePlot, CalculationError } from './basic-charts.ts';
import type { Dataset, Plot } from './basic-charts.ts';

export type PreviewKind = 'clustered' | 'stacked' | 'positive-100' | 'line' | 'area' | 'pie';
const colors = ['#2563eb', '#0d9488', '#d97706', '#9333ea', '#dc2626', '#475569'];
function escape(value: string): string {
  return value.replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[c]!);
}
function number(value: number): string {
  if (!Number.isFinite(value)) throw new CalculationError('INVALID_PREVIEW_COORDINATE');
  return Object.is(value, -0) ? '0' : String(value);
}
/** Internal S2 geometry preview; not a host command or persisted scene contract. */
export function svgPreview(dataset: Dataset, plot: Plot, kind: PreviewKind, title = 'Chart preview'): string {
  validatePlot(plot);
  if (typeof title !== 'string' || [...title].length > 1024) throw new CalculationError('INVALID_PREVIEW_TITLE');
  const color = (index: number) => colors[index % colors.length]!;
  const seriesColor = (id: string) => color(dataset.series.findIndex(s => s.id === id));
  let body: string;
  if (kind === 'clustered' || kind === 'stacked' || kind === 'positive-100') {
    const scene = kind === 'clustered' ? clustered(dataset, plot) : kind === 'stacked' ? stacked(dataset, plot) : positive100(dataset, plot);
    body = scene.bars.map(b => `<rect data-semantic-id="${escape(b.id)}" x="${number(b.x)}" y="${number(b.y)}" width="${number(b.width)}" height="${number(b.height)}" fill="${color(b.seriesIndex)}"/>`).join('');
  } else if (kind === 'line' || kind === 'area') {
    const scene = line(dataset, plot, { gapPolicy: 'break', includeZero: true });
    body = kind === 'area' ? area(dataset, plot, { gapPolicy: 'break', baseline: 0 }).polygons.map(p =>
      `<polygon data-semantic-id="${escape(p.id)}" points="${p.points.map(p => `${number(p.x)},${number(p.y)}`).join(' ')}" fill="${seriesColor(p.seriesId)}" fill-opacity="0.35"/>`).join('') : '';
    body += scene.paths.map(path => path.segments.map(points => `<polyline points="${points.map(p => `${number(p.x)},${number(p.y)}`).join(' ')}" fill="none" stroke="${seriesColor(path.seriesId)}" stroke-width="2"/>`).join('') + path.points.map(p =>
      `<circle data-semantic-id="${escape(p.id)}" cx="${number(p.x)}" cy="${number(p.y)}" r="3" fill="${seriesColor(path.seriesId)}"/>`).join('')).join('');
  } else if (kind === 'pie') {
    const scene = pie(dataset, plot), { x, y } = scene.center, r = scene.radius;
    const point = (angle: number) => `${number(x + r * Math.cos(angle * Math.PI / 180))},${number(y + r * Math.sin(angle * Math.PI / 180))}`;
    body = scene.sectors.map((s, i) => s.sweep === 360
      ? `<circle data-semantic-id="${escape(s.id)}" cx="${number(x)}" cy="${number(y)}" r="${number(r)}" fill="${color(i)}"/>`
      : `<path data-semantic-id="${escape(s.id)}" d="M ${number(x)},${number(y)} L ${point(s.startAngle)} A ${number(r)},${number(r)} 0 ${s.sweep > 180 ? 1 : 0},1 ${point(s.endAngle)} Z" fill="${color(i)}"/>`).join('');
  } else throw new CalculationError('UNSUPPORTED_PREVIEW_KIND');
  return `<svg xmlns="http://www.w3.org/2000/svg" role="img" viewBox="${number(plot.x)} ${number(plot.y)} ${number(plot.width)} ${number(plot.height)}" width="${number(plot.width)}pt" height="${number(plot.height)}pt"><title>${escape(title)}</title>${body}</svg>`;
}
