import type { CategoryChartData, ChartGrouping, ChartKind, ChartSpec } from './types.js';

const finite = (value: number, fallback: number): number =>
  Number.isFinite(value) ? value : fallback;

const clamp = (value: number, min: number, max: number): number =>
  Math.min(max, Math.max(min, value));

export const isCircularKind = (kind: ChartKind): boolean => kind === 'pie' || kind === 'doughnut';

export const valueFor = (value: number | null): number | null =>
  value !== null && Number.isFinite(value) ? value : null;

export const boundedOption = (
  value: number | undefined,
  min: number,
  max: number,
  fallback: number,
): number => (value !== undefined && Number.isFinite(value) ? clamp(value, min, max) : fallback);

const saturatedAdd = (left: number, right: number): number => {
  const sum = left + right;
  if (Number.isFinite(sum)) return sum;
  return left >= 0 && right >= 0 ? Number.MAX_VALUE : -Number.MAX_VALUE;
};

/**
 * Sum ordinary stack totals in the finite spreadsheet number range. If the
 * mathematical sum exceeds that range, the representable axis is clipped at
 * ±MAX_VALUE; percent stacks use stableStackTotal below so their ratios do
 * not depend on that clipping.
 */
export const effectiveStackTotals = (
  data: CategoryChartData,
): { positive: number[]; negative: number[] } => {
  const positive = Array.from({ length: data.categories.length }, () => 0);
  const negative = Array.from({ length: data.categories.length }, () => 0);
  for (const series of data.series) {
    for (let index = 0; index < data.categories.length; index += 1) {
      const value = valueFor(series.values[index] ?? null);
      if (value === null) continue;
      if (value >= 0) {
        positive[index] = saturatedAdd(positive[index] ?? 0, value);
      } else {
        negative[index] = saturatedAdd(negative[index] ?? 0, value);
      }
    }
  }
  return { positive, negative };
};

interface StableStackTotal {
  readonly scale: number;
  readonly total: number;
}

const stableStackTotal = (
  data: CategoryChartData,
  pointIndex: number,
  positive: boolean,
): StableStackTotal => {
  let scale = 0;
  for (const series of data.series) {
    const value = valueFor(series.values[pointIndex] ?? null);
    if (value === null || value >= 0 !== positive) continue;
    scale = Math.max(scale, Math.abs(value));
  }
  if (!(scale > 0)) return { scale: 1, total: 0 };
  let total = 0;
  for (const series of data.series) {
    const value = valueFor(series.values[pointIndex] ?? null);
    if (value === null || value >= 0 !== positive) continue;
    total = saturatedAdd(total, Math.abs(value) / scale);
  }
  return { scale, total };
};

interface StackSegment {
  readonly start: number;
  readonly end: number;
  readonly value: number;
}

export const segmentsFor = (
  data: CategoryChartData,
  grouping: ChartGrouping,
): readonly (readonly (StackSegment | null)[])[] => {
  const stableTotals =
    grouping === 'percent-stacked'
      ? data.categories.map((_category, pointIndex) => ({
          positive: stableStackTotal(data, pointIndex, true),
          negative: stableStackTotal(data, pointIndex, false),
        }))
      : [];
  const result: (StackSegment | null)[][] = data.series.map(() =>
    Array.from({ length: data.categories.length }, () => null),
  );
  for (let seriesIndex = 0; seriesIndex < data.series.length; seriesIndex += 1) {
    const series = data.series[seriesIndex];
    const target = result[seriesIndex];
    if (!series || !target) continue;
    for (let pointIndex = 0; pointIndex < data.categories.length; pointIndex += 1) {
      const raw = series.values[pointIndex] ?? null;
      const value = valueFor(raw);
      if (value === null) continue;
      if (grouping === 'clustered') {
        target[pointIndex] = { start: 0, end: value, value };
        continue;
      }
      const positive = value >= 0;
      const stableTotal = (positive
        ? stableTotals[pointIndex]?.positive
        : stableTotals[pointIndex]?.negative) ?? {
        scale: 1,
        total: 0,
      };
      // Keep positive and negative accumulators separate by deriving them from
      // preceding segments. This avoids mutating the input data model.
      let start = 0;
      for (let prior = 0; prior < seriesIndex; prior += 1) {
        const priorValue = data.series[prior]?.values[pointIndex] ?? null;
        const normalized = valueFor(priorValue);
        if (normalized === null || normalized >= 0 !== positive) continue;
        start = saturatedAdd(start, normalized);
      }
      let rendered = value;
      if (grouping === 'percent-stacked') {
        rendered = stableTotal.total === 0 ? 0 : value / stableTotal.scale / stableTotal.total;
        start = 0;
        for (let prior = 0; prior < seriesIndex; prior += 1) {
          const priorValue = valueFor(data.series[prior]?.values[pointIndex] ?? null);
          if (priorValue === null || priorValue >= 0 !== positive) continue;
          start = saturatedAdd(
            start,
            stableTotal.total === 0 ? 0 : priorValue / stableTotal.scale / stableTotal.total,
          );
        }
      }
      target[pointIndex] = { start, end: saturatedAdd(start, rendered), value: rendered };
    }
  }
  return result;
};

interface BarGeometry {
  readonly size: number;
  readonly seriesStep: number;
  readonly groupOffset: number;
}

/** Resolve Excel's gapWidth/overlap equations for one category slot. */
export const barGeometry = (
  slot: number,
  seriesCount: number,
  grouping: ChartGrouping,
  spec: ChartSpec,
): BarGeometry => {
  const gapWidth = boundedOption(spec.gapWidth, 0, 500, 150);
  const overlap = boundedOption(spec.overlap, -100, 100, 0);
  const count = Math.max(1, seriesCount);
  if (grouping !== 'clustered') {
    const size = slot / (1 + gapWidth / 100);
    return {
      size: finite(size, 0),
      seriesStep: 0,
      groupOffset: finite((slot - size) / 2, 0),
    };
  }
  const seriesStepFactor = 1 - overlap / 100;
  const size = slot / (1 + (count - 1) * seriesStepFactor + gapWidth / 100);
  const cluster = size + (count - 1) * size * seriesStepFactor;
  return {
    size: finite(size, 0),
    seriesStep: finite(size * seriesStepFactor, 0),
    groupOffset: finite((slot - cluster) / 2, 0),
  };
};

interface NormalizedMagnitudes {
  readonly values: readonly number[];
  readonly total: number;
}

/** Normalize one circular ring without overflowing on spreadsheet extremes. */
export const normalizedMagnitudes = (
  values: readonly (number | null)[],
): NormalizedMagnitudes | null => {
  const magnitudes = values.map((value) =>
    value !== null && Number.isFinite(value) && value !== 0 ? Math.abs(value) : 0,
  );
  const scale = magnitudes.reduce((maximum, value) => Math.max(maximum, value), 0);
  if (!(scale > 0) || !Number.isFinite(scale)) return null;
  const normalized = magnitudes.map((value) => value / scale);
  let total = 0;
  for (const value of normalized) {
    total += value;
    if (!Number.isFinite(total)) return null;
  }
  return total > 0 ? { values: normalized, total } : null;
};

interface CircularPoint {
  readonly x: number;
  readonly y: number;
}

export const circularPoint = (
  cx: number,
  cy: number,
  radius: number,
  angle: number,
): CircularPoint => ({
  x: cx + Math.cos(angle) * radius,
  y: cy + Math.sin(angle) * radius,
});

/** Build an annular SVG path, including two arcs for each side of a full ring. */
export const annularPath = (
  cx: number,
  cy: number,
  innerRadius: number,
  outerRadius: number,
  start: number,
  angle: number,
): string => {
  const fullCircle = angle >= Math.PI * 2 - 1e-9;
  const outerStart = circularPoint(cx, cy, outerRadius, start);
  const innerStart = circularPoint(cx, cy, innerRadius, start);
  if (fullCircle) {
    const outerMid = circularPoint(cx, cy, outerRadius, start + Math.PI);
    const innerMid = circularPoint(cx, cy, innerRadius, start + Math.PI);
    return [
      `M ${outerStart.x} ${outerStart.y}`,
      `A ${outerRadius} ${outerRadius} 0 1 1 ${outerMid.x} ${outerMid.y}`,
      `A ${outerRadius} ${outerRadius} 0 1 1 ${outerStart.x} ${outerStart.y}`,
      `M ${innerStart.x} ${innerStart.y}`,
      `A ${innerRadius} ${innerRadius} 0 1 0 ${innerMid.x} ${innerMid.y}`,
      `A ${innerRadius} ${innerRadius} 0 1 0 ${innerStart.x} ${innerStart.y}`,
      'Z',
    ].join(' ');
  }
  const end = start + angle;
  const outerEnd = circularPoint(cx, cy, outerRadius, end);
  const innerEnd = circularPoint(cx, cy, innerRadius, end);
  const largeArc = angle > Math.PI ? 1 : 0;
  return [
    `M ${outerStart.x} ${outerStart.y}`,
    `A ${outerRadius} ${outerRadius} 0 ${largeArc} 1 ${outerEnd.x} ${outerEnd.y}`,
    `L ${innerEnd.x} ${innerEnd.y}`,
    `A ${innerRadius} ${innerRadius} 0 ${largeArc} 0 ${innerStart.x} ${innerStart.y}`,
    'Z',
  ].join(' ');
};
