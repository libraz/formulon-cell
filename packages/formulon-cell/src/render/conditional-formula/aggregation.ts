import type { CellValue } from '../../engine/types.js';
import type { FormulaAggregateName } from './types.js';

const aggregateValueA = (value: CellValue): number | null => {
  if (value.kind === 'number' && Number.isFinite(value.value)) return value.value;
  if (value.kind === 'bool') return value.value ? 1 : 0;
  if (value.kind === 'text') return 0;
  return null;
};
const aggregateResult = (
  fn: FormulaAggregateName,
  values: number[],
  valuesA: number[],
  countA: number,
  countBlank: number,
): CellValue => {
  if (fn === 'COUNT') return { kind: 'number', value: values.length };
  if (fn === 'COUNTA') return { kind: 'number', value: countA };
  if (fn === 'COUNTBLANK') return { kind: 'number', value: countBlank };
  if (fn === 'SUM') {
    return { kind: 'number', value: values.reduce((sum, value) => sum + value, 0) };
  }
  if (fn === 'AVERAGEA' || fn === 'MINA' || fn === 'MAXA') {
    if (valuesA.length === 0) return { kind: 'error', code: 15, text: '#VALUE!' };
    if (fn === 'AVERAGEA') {
      return {
        kind: 'number',
        value: valuesA.reduce((sum, value) => sum + value, 0) / valuesA.length,
      };
    }
    return {
      kind: 'number',
      value: fn === 'MINA' ? Math.min(...valuesA) : Math.max(...valuesA),
    };
  }
  if (fn === 'PRODUCT') {
    return {
      kind: 'number',
      value: values.length === 0 ? 0 : values.reduce((product, value) => product * value, 1),
    };
  }
  if (values.length === 0) return { kind: 'error', code: 15, text: '#VALUE!' };
  if (fn === 'AVERAGE') {
    return {
      kind: 'number',
      value: values.reduce((sum, value) => sum + value, 0) / values.length,
    };
  }
  if (fn === 'MEDIAN') {
    const sorted = [...values].sort((a, b) => a - b);
    const mid = Math.floor(sorted.length / 2);
    return {
      kind: 'number',
      value:
        sorted.length % 2 === 1
          ? (sorted[mid] as number)
          : ((sorted[mid - 1] as number) + (sorted[mid] as number)) / 2,
    };
  }
  if (fn === 'MODE' || fn === 'MODE.SNGL') {
    const counts = new Map<number, number>();
    let mode: number | null = null;
    let bestCount = 1;
    for (const value of values) {
      const count = (counts.get(value) ?? 0) + 1;
      counts.set(value, count);
      if (count > bestCount) {
        bestCount = count;
        mode = value;
      }
    }
    return mode === null
      ? { kind: 'error', code: 6, text: '#N/A' }
      : { kind: 'number', value: mode };
  }
  if (fn === 'DEVSQ') {
    const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
    return {
      kind: 'number',
      value: values.reduce((sum, value) => sum + (value - mean) ** 2, 0),
    };
  }
  if (fn === 'AVEDEV') {
    const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
    return {
      kind: 'number',
      value: values.reduce((sum, value) => sum + Math.abs(value - mean), 0) / values.length,
    };
  }
  if (fn === 'SKEW') {
    if (values.length < 3) return { kind: 'error', code: 1, text: '#DIV/0!' };
    const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
    const sampleVariance =
      values.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (values.length - 1);
    const sampleDeviation = Math.sqrt(sampleVariance);
    if (sampleDeviation === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
    const skew =
      (values.length / ((values.length - 1) * (values.length - 2))) *
      values.reduce((sum, value) => sum + ((value - mean) / sampleDeviation) ** 3, 0);
    return { kind: 'number', value: skew };
  }
  if (fn === 'SKEW.P') {
    if (values.length < 3) return { kind: 'error', code: 1, text: '#DIV/0!' };
    const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
    const populationVariance =
      values.reduce((sum, value) => sum + (value - mean) ** 2, 0) / values.length;
    const populationDeviation = Math.sqrt(populationVariance);
    if (populationDeviation === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
    return {
      kind: 'number',
      value:
        values.reduce((sum, value) => sum + ((value - mean) / populationDeviation) ** 3, 0) /
        values.length,
    };
  }
  if (fn === 'KURT') {
    if (values.length < 4) return { kind: 'error', code: 1, text: '#DIV/0!' };
    const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
    const sampleVariance =
      values.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (values.length - 1);
    const sampleDeviation = Math.sqrt(sampleVariance);
    if (sampleDeviation === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
    const n = values.length;
    const sumFourthPowers = values.reduce(
      (sum, value) => sum + ((value - mean) / sampleDeviation) ** 4,
      0,
    );
    const kurtosis =
      (n * (n + 1) * sumFourthPowers) / ((n - 1) * (n - 2) * (n - 3)) -
      (3 * (n - 1) ** 2) / ((n - 2) * (n - 3));
    return { kind: 'number', value: kurtosis };
  }
  if (fn === 'GEOMEAN' || fn === 'HARMEAN') {
    if (values.some((value) => value <= 0)) {
      return { kind: 'error', code: 6, text: '#NUM!' };
    }
    if (fn === 'GEOMEAN') {
      return {
        kind: 'number',
        value: Math.exp(values.reduce((sum, value) => sum + Math.log(value), 0) / values.length),
      };
    }
    return {
      kind: 'number',
      value: values.length / values.reduce((sum, value) => sum + 1 / value, 0),
    };
  }
  if (
    fn === 'VAR' ||
    fn === 'VARP' ||
    fn === 'VAR.S' ||
    fn === 'VAR.P' ||
    fn === 'STDEV' ||
    fn === 'STDEVP' ||
    fn === 'STDEV.S' ||
    fn === 'STDEV.P'
  ) {
    const sample = fn === 'VAR' || fn === 'STDEV' || fn.endsWith('.S');
    if (values.length < (sample ? 2 : 1)) {
      return { kind: 'error', code: 1, text: '#DIV/0!' };
    }
    const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
    const variance =
      values.reduce((sum, value) => sum + (value - mean) ** 2, 0) /
      (sample ? values.length - 1 : values.length);
    return {
      kind: 'number',
      value: fn.startsWith('STDEV') ? Math.sqrt(variance) : variance,
    };
  }
  return {
    kind: 'number',
    value: fn === 'MIN' ? Math.min(...values) : Math.max(...values),
  };
};
const subtotalFunction = (functionNum: number): FormulaAggregateName | null => {
  const code = Math.trunc(functionNum);
  const normalized = code >= 101 && code <= 111 ? code - 100 : code;
  switch (normalized) {
    case 1:
      return 'AVERAGE';
    case 2:
      return 'COUNT';
    case 3:
      return 'COUNTA';
    case 4:
      return 'MAX';
    case 5:
      return 'MIN';
    case 6:
      return 'PRODUCT';
    case 7:
      return 'STDEV';
    case 8:
      return 'STDEVP';
    case 9:
      return 'SUM';
    case 10:
      return 'VAR';
    case 11:
      return 'VARP';
    default:
      return null;
  }
};
const aggregateFunction = (
  functionNum: number,
):
  | { kind: 'aggregate'; fn: FormulaAggregateName }
  | { kind: 'ranked'; fn: 'LARGE' | 'SMALL' }
  | {
      kind: 'percentile';
      fn: 'PERCENTILE.INC' | 'QUARTILE.INC' | 'PERCENTILE.EXC' | 'QUARTILE.EXC';
    }
  | null => {
  switch (Math.trunc(functionNum)) {
    case 1:
      return { kind: 'aggregate', fn: 'AVERAGE' };
    case 2:
      return { kind: 'aggregate', fn: 'COUNT' };
    case 3:
      return { kind: 'aggregate', fn: 'COUNTA' };
    case 4:
      return { kind: 'aggregate', fn: 'MAX' };
    case 5:
      return { kind: 'aggregate', fn: 'MIN' };
    case 6:
      return { kind: 'aggregate', fn: 'PRODUCT' };
    case 7:
      return { kind: 'aggregate', fn: 'STDEV.S' };
    case 8:
      return { kind: 'aggregate', fn: 'STDEV.P' };
    case 9:
      return { kind: 'aggregate', fn: 'SUM' };
    case 10:
      return { kind: 'aggregate', fn: 'VAR.S' };
    case 11:
      return { kind: 'aggregate', fn: 'VAR.P' };
    case 12:
      return { kind: 'aggregate', fn: 'MEDIAN' };
    case 13:
      return { kind: 'aggregate', fn: 'MODE.SNGL' };
    case 14:
      return { kind: 'ranked', fn: 'LARGE' };
    case 15:
      return { kind: 'ranked', fn: 'SMALL' };
    case 16:
      return { kind: 'percentile', fn: 'PERCENTILE.INC' };
    case 17:
      return { kind: 'percentile', fn: 'QUARTILE.INC' };
    case 18:
      return { kind: 'percentile', fn: 'PERCENTILE.EXC' };
    case 19:
      return { kind: 'percentile', fn: 'QUARTILE.EXC' };
    default:
      return null;
  }
};
const percentileIncValue = (values: number[], k: number): number => {
  if (values.length === 1) return values[0] as number;
  const sorted = values.slice().sort((a, b) => a - b);
  const position = k * (sorted.length - 1);
  const lower = Math.floor(position);
  const upper = Math.ceil(position);
  if (lower === upper) return sorted[lower] as number;
  const fraction = position - lower;
  return (
    (sorted[lower] as number) + ((sorted[upper] as number) - (sorted[lower] as number)) * fraction
  );
};
const percentileExcValue = (values: number[], k: number): number | null => {
  if (k <= 0 || k >= 1) return null;
  const sorted = values.slice().sort((a, b) => a - b);
  const position = k * (sorted.length + 1);
  if (position < 1 || position > sorted.length) return null;
  const lower = Math.floor(position);
  const upper = Math.ceil(position);
  if (lower === upper) return sorted[lower - 1] as number;
  const lowerValue = sorted[lower - 1] as number;
  const upperValue = sorted[upper - 1] as number;
  return lowerValue + (upperValue - lowerValue) * (position - lower);
};

export {
  aggregateFunction,
  aggregateResult,
  aggregateValueA,
  percentileExcValue,
  percentileIncValue,
  subtotalFunction,
};
