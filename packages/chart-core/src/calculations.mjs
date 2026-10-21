/** Pure deterministic calculations. No PowerPoint or model dependencies. */
export function stackedTotals(series) {
  if (!Array.isArray(series) || series.length === 0) return [];
  const width = series[0].length;
  if (!series.every(row => Array.isArray(row) && row.length === width)) throw new Error('Series lengths differ');
  return Array.from({ length: width }, (_, i) => series.reduce((sum, row) => {
    const value = row[i];
    if (value === null) return sum;
    if (typeof value !== 'number' || !Number.isFinite(value)) throw new Error('Invalid numeric value');
    return sum + value;
  }, 0));
}

export function normalizePositive100Stack(series) {
  if (!Array.isArray(series) || series.length === 0) return [];
  const totals = stackedTotals(series);
  if (series.some(row => row.some(v => v !== null && v < 0))) throw new Error('Negative values not supported by positive-only 100% stack');
  return series.map(row => row.map((value, i) => totals[i] === 0 ? null : (value ?? 0) / totals[i] * 100));
}

export function waterfallSegments(points, initial = 0) {
  if (!Number.isFinite(initial)) throw new Error('Invalid initial value');
  let running = initial;
  return points.map((p) => {
    if (!p || !Number.isFinite(p.value)) throw new Error('Invalid waterfall point');
    const role = p.role ?? 'delta';
    if (!['delta', 'subtotal', 'total'].includes(role)) throw new Error('Invalid point role');
    if (role === 'delta') {
      const start = running;
      running += p.value;
      return { role, start, end: running, low: Math.min(start, running), high: Math.max(start, running), running };
    }
    // Explicit total/subtotal shows accumulated value; it must not add again.
    return { role, start: 0, end: running, low: Math.min(0, running), high: Math.max(0, running), running };
  });
}

export function cagr(start, end, periods) {
  if (![start, end, periods].every(Number.isFinite) || periods <= 0 || start <= 0 || end < 0) throw new Error('Unsupported CAGR domain');
  return (end / start) ** (1 / periods) - 1;
}
