import {
  automaticAxisDomain,
  linearScale,
  niceDomain,
  resolvePositivePowerDomain,
  unclampedLogProjection,
  validateLogBase,
} from './scales.js';
import type { ChartAxisSpec, ChartScene, ChartSpec } from './types.js';

const clamp = (value: number, min: number, max: number): number =>
  Math.min(max, Math.max(min, value));

export interface Domain {
  min: number;
  max: number;
  /** The major unit an automatic axis resolved along with its bounds. */
  majorUnit?: number;
}

export const axisDomain = (
  values: readonly number[],
  explicitMin: number | undefined,
  explicitMax: number | undefined,
  percent: boolean,
  logBase?: number,
): Domain => {
  if (logBase !== undefined) {
    return resolvePositivePowerDomain(values, explicitMin, explicitMax, logBase);
  }
  const finiteValues = values.filter(Number.isFinite);
  let min = 0;
  let max = 1;
  if (finiteValues.length > 0) {
    min = finiteValues.reduce(
      (current, value) => Math.min(current, value),
      Number.POSITIVE_INFINITY,
    );
    max = finiteValues.reduce(
      (current, value) => Math.max(current, value),
      Number.NEGATIVE_INFINITY,
    );
  }
  if (percent) {
    if (min < 0 && max > 0) {
      min = -1;
      max = 1;
    } else if (max > 0) {
      min = 0;
      max = 1;
    } else {
      min = -1;
      max = 0;
    }
  }
  const hasExplicitMin = explicitMin !== undefined && Number.isFinite(explicitMin);
  const hasExplicitMax = explicitMax !== undefined && Number.isFinite(explicitMax);
  const baseMin = min;
  const baseMax = max;
  const spanForMissingSide = Math.max(
    Number.isFinite(baseMax - baseMin) && baseMax > baseMin ? baseMax - baseMin : 0,
    Math.abs(baseMin) * 0.1,
    Math.abs(baseMax) * 0.1,
    1,
  );
  if (hasExplicitMin && hasExplicitMax) {
    min = explicitMin as number;
    max = explicitMax as number;
    if (min > max) [min, max] = [max, min];
  } else if (hasExplicitMin) {
    min = explicitMin as number;
    if (max < min) max = Math.min(Number.MAX_VALUE, min + spanForMissingSide);
  } else if (hasExplicitMax) {
    max = explicitMax as number;
    if (min > max) min = Math.max(-Number.MAX_VALUE, max - spanForMissingSide);
  }
  if (!percent && (!hasExplicitMin || !hasExplicitMax)) {
    return automaticAxisDomain(min, max, {
      ...(hasExplicitMin ? { min } : {}),
      ...(hasExplicitMax ? { max } : {}),
    });
  }
  if (min === max) {
    const padding = Math.max(Math.abs(min) * 0.1, Number.MIN_VALUE, percent ? 1 : 1);
    if (hasExplicitMin && !hasExplicitMax) max = Math.min(Number.MAX_VALUE, min + padding);
    else if (hasExplicitMax && !hasExplicitMin) min = Math.max(-Number.MAX_VALUE, max - padding);
    else {
      min = Math.max(-Number.MAX_VALUE, min - padding);
      max = Math.min(Number.MAX_VALUE, max + padding);
    }
  }
  // Treat non-finite explicit bounds exactly like omitted bounds.  Checking
  // the raw option for `undefined` here would leave NaN/Infinity in the
  // auto-expansion decision even though the finite flags above rejected them.
  if (!hasExplicitMin || !hasExplicitMax) {
    const expanded = niceDomain(min, max, 4);
    if (!hasExplicitMin) min = expanded.min;
    if (!hasExplicitMax) max = expanded.max;
  }
  return { min, max };
};

export const reversedAxis = (axis: ChartAxisSpec | undefined): boolean =>
  axis?.reverseOrder === true;

const resolvedLogBase = (axis: ChartAxisSpec | undefined): number | undefined =>
  validateLogBase(axis?.logBase);

/** Numeric axes are log-capable only for the chart families with one. */
export const categoryNumericLogBase = (spec: ChartSpec, axis: 'x' | 'y'): number | undefined => {
  if ((spec.grouping ?? 'clustered') !== 'clustered') return undefined;
  if (axis === 'x' && spec.kind === 'bar') return resolvedLogBase(spec.xAxis);
  if (axis === 'y' && (spec.kind === 'column' || spec.kind === 'line' || spec.kind === 'area')) {
    return resolvedLogBase(spec.yAxis);
  }
  return undefined;
};

export const scatterLogBases = (spec: ChartSpec): { readonly x?: number; readonly y?: number } =>
  spec.kind === 'scatter'
    ? {
        x: resolvedLogBase(spec.xAxis),
        y: resolvedLogBase(spec.yAxis),
      }
    : {};

export const categoryIndex = (count: number, index: number, reverse: boolean): number =>
  reverse ? count - 1 - index : index;

export const xCategory = (
  plot: ChartScene['plot'],
  count: number,
  index: number,
  reverse = false,
): number => {
  if (count <= 1) return plot.x + plot.width / 2;
  return plot.x + ((categoryIndex(count, index, reverse) + 0.5) * plot.width) / count;
};

export const yCategory = (
  plot: ChartScene['plot'],
  count: number,
  index: number,
  reverse = false,
): number => {
  if (count <= 1) return plot.y + plot.height / 2;
  return plot.y + ((categoryIndex(count, index, reverse) + 0.5) * plot.height) / count;
};

/** Project an ascending numeric domain onto a screen range, optionally reversing its endpoints. */
export const numericProjection = (
  value: number,
  domain: Domain,
  start: number,
  end: number,
  reverse: boolean,
  logBase?: number,
): number => {
  const from = reverse ? end : start;
  const to = reverse ? start : end;
  const projected =
    logBase === undefined
      ? (linearScale(value, domain.min, domain.max, from, to) ?? from)
      : unclampedLogProjection(value, domain.min, domain.max, from, to);
  return clamp(projected, Math.min(start, end), Math.max(start, end));
};
