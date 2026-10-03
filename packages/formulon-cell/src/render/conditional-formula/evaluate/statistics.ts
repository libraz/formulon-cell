import { addrKey } from '../../../engine/address.js';
import type { CellValue } from '../../../engine/types.js';
import { percentileExcValue, percentileIncValue } from '../aggregation.js';
import { positiveInteger, readNumber } from '../coercion.js';
import {
  regularizedBeta,
  regularizedGammaP,
  standardNormalCdf,
  studentTCdf,
} from '../distributions.js';
import type { FormulaAggregateArg, FormulaRangeArg, ParsedA1Range } from '../types.js';
import type { FormulaReaderContext } from './context.js';
import type { RangeReader } from './ranges.js';

/** Exactly the reader members this family touches. */
export type StatisticsEvaluatorContext = Pick<FormulaReaderContext, 'state' | 'sheet'> &
  Pick<
    RangeReader,
    | 'formulaRangeArgBounds'
    | 'validRangeBounds'
    | 'numericValuesFromArgs'
    | 'numericValuesInFormulaRangeArg'
    | 'numericValuesInRangeWithShape'
    | 'numericPairsInRanges'
  >;

/** Range statistics: LARGE/SMALL, percentiles and ranks, paired-range
 *  statistics and regression, PROB, Z/T/CHISQ tests, SERIESSUM, SUMPRODUCT. */
export function createStatisticsEvaluator(ctx: StatisticsEvaluatorContext) {
  const {
    state,
    sheet,
    formulaRangeArgBounds,
    validRangeBounds,
    numericValuesFromArgs,
    numericValuesInFormulaRangeArg,
    numericValuesInRangeWithShape,
    numericPairsInRanges,
  } = ctx;
  const rankedRangeValue = (
    fn: 'LARGE' | 'SMALL',
    range: FormulaRangeArg | ParsedA1Range,
    rankValue: CellValue,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const rank = positiveInteger(rankValue);
    const values = numericValuesInFormulaRangeArg(range, rowOffset, colOffset);
    if (rank === null || values === null || rank > values.length) {
      return { kind: 'error', code: 6, text: '#NUM!' };
    }
    values.sort((a, b) => (fn === 'LARGE' ? b - a : a - b));
    return { kind: 'number', value: values[rank - 1] as number };
  };
  const percentileRangeValue = (
    fn:
      | 'PERCENTILE.INC'
      | 'PERCENTILE.EXC'
      | 'PERCENTILE'
      | 'QUARTILE.INC'
      | 'QUARTILE.EXC'
      | 'QUARTILE'
      | 'PERCENTRANK'
      | 'PERCENTRANK.INC'
      | 'PERCENTRANK.EXC',
    range: FormulaRangeArg | ParsedA1Range,
    valueCell: CellValue,
    significanceCell: CellValue | null,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const value = readNumber(valueCell);
    const values = numericValuesInFormulaRangeArg(range, rowOffset, colOffset);
    if (value === null || values === null || values.length === 0) {
      return { kind: 'error', code: 6, text: '#NUM!' };
    }
    if (fn === 'PERCENTRANK' || fn === 'PERCENTRANK.INC' || fn === 'PERCENTRANK.EXC') {
      if (values.length === 1) {
        return (fn === 'PERCENTRANK' || fn === 'PERCENTRANK.INC') && values[0] === value
          ? { kind: 'number', value: 0 }
          : { kind: 'error', code: 6, text: '#N/A' };
      }
      const significance = significanceCell === null ? 3 : positiveInteger(significanceCell);
      if (significance === null) return { kind: 'error', code: 6, text: '#NUM!' };
      const sorted = values.slice().sort((a, b) => a - b);
      if (value < (sorted[0] as number) || value > (sorted[sorted.length - 1] as number)) {
        return { kind: 'error', code: 6, text: '#N/A' };
      }
      let position: number | null = null;
      for (let index = 0; index < sorted.length; index += 1) {
        if (sorted[index] === value) {
          position = index;
          break;
        }
        const next = sorted[index + 1];
        if (next !== undefined && value > (sorted[index] as number) && value < next) {
          position =
            index + (value - (sorted[index] as number)) / (next - (sorted[index] as number));
          break;
        }
      }
      if (position === null) return { kind: 'error', code: 6, text: '#N/A' };
      const factor = 10 ** significance;
      const isInclusive = fn === 'PERCENTRANK' || fn === 'PERCENTRANK.INC';
      const denominator = isInclusive ? sorted.length - 1 : sorted.length + 1;
      const numerator = isInclusive ? position : position + 1;
      return {
        kind: 'number',
        value: Math.trunc((numerator / denominator) * factor) / factor,
      };
    }
    const isQuartile = fn === 'QUARTILE' || fn === 'QUARTILE.INC' || fn === 'QUARTILE.EXC';
    const k = isQuartile ? Math.trunc(value) / 4 : value;
    if (
      ((fn === 'PERCENTILE' || fn === 'PERCENTILE.INC') && (value < 0 || value > 1)) ||
      ((fn === 'QUARTILE' || fn === 'QUARTILE.INC') && (value < 0 || value > 4)) ||
      (fn === 'QUARTILE.EXC' && (value < 1 || value > 3))
    ) {
      return { kind: 'error', code: 6, text: '#NUM!' };
    }
    if (fn === 'PERCENTILE.EXC' || fn === 'QUARTILE.EXC') {
      const result = percentileExcValue(values, k);
      return result === null
        ? { kind: 'error', code: 6, text: '#NUM!' }
        : { kind: 'number', value: result };
    }
    return { kind: 'number', value: percentileIncValue(values, k) };
  };
  const rankRangeValue = (
    fn: 'RANK' | 'RANK.EQ' | 'RANK.AVG',
    valueCell: CellValue,
    range: FormulaRangeArg | ParsedA1Range,
    orderCell: CellValue | null,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const value = readNumber(valueCell);
    const order = orderCell === null ? 0 : readNumber(orderCell);
    const values = numericValuesInFormulaRangeArg(range, rowOffset, colOffset);
    if (value === null || order === null || values === null || values.length === 0) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const same = values.filter((candidate) => candidate === value).length;
    if (same === 0) return { kind: 'error', code: 6, text: '#N/A' };
    const before = values.filter((candidate) =>
      order === 0 ? candidate > value : candidate < value,
    ).length;
    const rank = before + 1;
    return { kind: 'number', value: fn === 'RANK.AVG' ? rank + (same - 1) / 2 : rank };
  };
  const pairedRangeStatValue = (
    fn:
      | 'CORREL'
      | 'PEARSON'
      | 'COVAR'
      | 'COVARIANCE.P'
      | 'COVARIANCE.S'
      | 'SLOPE'
      | 'INTERCEPT'
      | 'RSQ'
      | 'STEYX'
      | 'SUMX2MY2'
      | 'SUMX2PY2'
      | 'SUMXMY2'
      | 'F.TEST'
      | 'FTEST',
    left: FormulaRangeArg | ParsedA1Range,
    right: FormulaRangeArg | ParsedA1Range,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    if (fn === 'F.TEST' || fn === 'FTEST') {
      const leftValues = numericValuesInFormulaRangeArg(left, rowOffset, colOffset);
      const rightValues = numericValuesInFormulaRangeArg(right, rowOffset, colOffset);
      if (!leftValues || !rightValues || leftValues.length < 2 || rightValues.length < 2) {
        return { kind: 'error', code: 1, text: '#DIV/0!' };
      }
      const variance = (values: number[]): number => {
        const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
        return values.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (values.length - 1);
      };
      const leftVariance = variance(leftValues);
      const rightVariance = variance(rightValues);
      if (leftVariance === 0 || rightVariance === 0) {
        return { kind: 'error', code: 1, text: '#DIV/0!' };
      }
      const ratio = leftVariance / rightVariance;
      const degreesLeft = leftValues.length - 1;
      const degreesRight = rightValues.length - 1;
      const transformed = (degreesLeft * ratio) / (degreesLeft * ratio + degreesRight);
      const leftTail = regularizedBeta(transformed, degreesLeft / 2, degreesRight / 2);
      if (leftTail === null) return { kind: 'error', code: 6, text: '#NUM!' };
      return { kind: 'number', value: Math.min(1, 2 * Math.min(leftTail, 1 - leftTail)) };
    }
    const pairs = numericPairsInRanges(left, right, rowOffset, colOffset);
    if (!pairs) return { kind: 'error', code: 15, text: '#VALUE!' };
    const { xs, ys } = pairs;
    if (fn === 'SUMX2MY2' || fn === 'SUMX2PY2' || fn === 'SUMXMY2') {
      if (xs.length === 0) return { kind: 'error', code: 15, text: '#VALUE!' };
      const value = xs.reduce((sum, leftValue, index) => {
        const rightValue = ys[index] as number;
        if (fn === 'SUMX2MY2') return sum + leftValue ** 2 - rightValue ** 2;
        if (fn === 'SUMX2PY2') return sum + leftValue ** 2 + rightValue ** 2;
        return sum + (leftValue - rightValue) ** 2;
      }, 0);
      return { kind: 'number', value };
    }
    if (xs.length < (fn === 'COVAR' || fn === 'COVARIANCE.P' ? 1 : 2)) {
      return { kind: 'error', code: 1, text: '#DIV/0!' };
    }
    const meanLeft = xs.reduce((sum, value) => sum + value, 0) / xs.length;
    const meanRight = ys.reduce((sum, value) => sum + value, 0) / ys.length;
    let covarianceNumerator = 0;
    let sumLeft = 0;
    let sumRight = 0;
    for (let index = 0; index < xs.length; index += 1) {
      const dLeft = (xs[index] as number) - meanLeft;
      const dRight = (ys[index] as number) - meanRight;
      covarianceNumerator += dLeft * dRight;
      sumLeft += dLeft ** 2;
      sumRight += dRight ** 2;
    }
    if (fn === 'CORREL' || fn === 'PEARSON' || fn === 'RSQ') {
      if (sumLeft === 0 || sumRight === 0) {
        return { kind: 'error', code: 1, text: '#DIV/0!' };
      }
      const correl = covarianceNumerator / Math.sqrt(sumLeft * sumRight);
      return { kind: 'number', value: fn === 'RSQ' ? correl ** 2 : correl };
    }
    if (fn === 'SLOPE' || fn === 'INTERCEPT') {
      if (sumRight === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
      const slope = covarianceNumerator / sumRight;
      return {
        kind: 'number',
        value: fn === 'SLOPE' ? slope : meanLeft - slope * meanRight,
      };
    }
    if (fn === 'STEYX') {
      if (sumRight === 0 || xs.length < 3) return { kind: 'error', code: 1, text: '#DIV/0!' };
      const slope = covarianceNumerator / sumRight;
      const intercept = meanLeft - slope * meanRight;
      const squaredError = xs.reduce((sum, y, index) => {
        const x = ys[index] as number;
        return sum + (y - (slope * x + intercept)) ** 2;
      }, 0);
      return { kind: 'number', value: Math.sqrt(squaredError / (xs.length - 2)) };
    }
    return {
      kind: 'number',
      value: covarianceNumerator / (fn === 'COVARIANCE.S' ? xs.length - 1 : xs.length),
    };
  };
  const regressionForecastValue = (
    xCell: CellValue,
    knownY: FormulaRangeArg | ParsedA1Range,
    knownX: FormulaRangeArg | ParsedA1Range,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const x = readNumber(xCell);
    if (x === null) return { kind: 'error', code: 15, text: '#VALUE!' };
    const pairs = numericPairsInRanges(knownY, knownX, rowOffset, colOffset);
    if (!pairs) return { kind: 'error', code: 15, text: '#VALUE!' };
    const { xs, ys } = pairs;
    if (xs.length < 2) return { kind: 'error', code: 1, text: '#DIV/0!' };
    const meanY = xs.reduce((sum, value) => sum + value, 0) / xs.length;
    const meanX = ys.reduce((sum, value) => sum + value, 0) / ys.length;
    let covarianceNumerator = 0;
    let sumX = 0;
    for (let index = 0; index < xs.length; index += 1) {
      const dy = (xs[index] as number) - meanY;
      const dx = (ys[index] as number) - meanX;
      covarianceNumerator += dy * dx;
      sumX += dx ** 2;
    }
    if (sumX === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
    const slope = covarianceNumerator / sumX;
    return { kind: 'number', value: slope * x + (meanY - slope * meanX) };
  };
  const probabilityRangeValue = (
    valuesRange: FormulaRangeArg | ParsedA1Range,
    probabilitiesRange: FormulaRangeArg | ParsedA1Range,
    lowerCell: CellValue,
    upperCell: CellValue | null,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const lower = readNumber(lowerCell);
    const upper = upperCell === null ? lower : readNumber(upperCell);
    if (lower === null || upper === null) return { kind: 'error', code: 15, text: '#VALUE!' };
    const valuesBounds = formulaRangeArgBounds(valuesRange, rowOffset, colOffset);
    const probabilitiesBounds = formulaRangeArgBounds(probabilitiesRange, rowOffset, colOffset);
    if (
      !valuesBounds ||
      !probabilitiesBounds ||
      !validRangeBounds(valuesBounds) ||
      !validRangeBounds(probabilitiesBounds) ||
      valuesBounds.width !== probabilitiesBounds.width ||
      valuesBounds.height !== probabilitiesBounds.height
    ) {
      return { kind: 'error', code: 6, text: '#N/A' };
    }
    let probabilityTotal = 0;
    let result = 0;
    for (let r = 0; r < valuesBounds.height; r += 1) {
      for (let c = 0; c < valuesBounds.width; c += 1) {
        const value = state.data.cells.get(
          addrKey({ sheet, row: valuesBounds.r0 + r, col: valuesBounds.c0 + c }),
        )?.value;
        const probability = state.data.cells.get(
          addrKey({ sheet, row: probabilitiesBounds.r0 + r, col: probabilitiesBounds.c0 + c }),
        )?.value;
        if (
          value?.kind !== 'number' ||
          !Number.isFinite(value.value) ||
          probability?.kind !== 'number' ||
          !Number.isFinite(probability.value)
        ) {
          return { kind: 'error', code: 15, text: '#VALUE!' };
        }
        if (probability.value < 0 || probability.value > 1) {
          return { kind: 'error', code: 6, text: '#NUM!' };
        }
        probabilityTotal += probability.value;
        if (value.value >= lower && value.value <= upper) result += probability.value;
      }
    }
    if (Math.abs(probabilityTotal - 1) > 1e-9) {
      return { kind: 'error', code: 6, text: '#NUM!' };
    }
    return { kind: 'number', value: result };
  };
  const zTestValue = (
    range: FormulaRangeArg | ParsedA1Range,
    xCell: CellValue,
    sigmaCell: CellValue | null,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const x = readNumber(xCell);
    const values = numericValuesInFormulaRangeArg(range, rowOffset, colOffset);
    const sigma = sigmaCell === null ? null : readNumber(sigmaCell);
    if (
      x === null ||
      values === null ||
      values.length === 0 ||
      (sigmaCell !== null && sigma === null)
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
    let standardDeviation = sigma;
    if (standardDeviation === null) {
      if (values.length < 2) return { kind: 'error', code: 1, text: '#DIV/0!' };
      const variance =
        values.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (values.length - 1);
      standardDeviation = Math.sqrt(variance);
    }
    if (standardDeviation <= 0) return { kind: 'error', code: 6, text: '#NUM!' };
    const z = (mean - x) / (standardDeviation / Math.sqrt(values.length));
    return { kind: 'number', value: 1 - standardNormalCdf(z) };
  };
  const tTestValue = (
    left: FormulaRangeArg | ParsedA1Range,
    right: FormulaRangeArg | ParsedA1Range,
    tailsCell: CellValue,
    typeCell: CellValue,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const rawTails = readNumber(tailsCell);
    const rawType = readNumber(typeCell);
    if (rawTails === null || rawType === null) return { kind: 'error', code: 15, text: '#VALUE!' };
    const tails = Math.trunc(rawTails);
    const type = Math.trunc(rawType);
    if ((tails !== 1 && tails !== 2) || type < 1 || type > 3) {
      return { kind: 'error', code: 6, text: '#NUM!' };
    }
    const leftValues = numericValuesInFormulaRangeArg(left, rowOffset, colOffset);
    const rightValues = numericValuesInFormulaRangeArg(right, rowOffset, colOffset);
    if (!leftValues || !rightValues) return { kind: 'error', code: 15, text: '#VALUE!' };
    const sampleStats = (values: number[]): { mean: number; variance: number } | null => {
      if (values.length < 2) return null;
      const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
      const variance =
        values.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (values.length - 1);
      return variance > 0 ? { mean, variance } : null;
    };
    let t: number;
    let degrees: number;
    if (type === 1) {
      if (leftValues.length !== rightValues.length) {
        return { kind: 'error', code: 6, text: '#N/A' };
      }
      const differences = leftValues.map((value, index) => value - (rightValues[index] as number));
      const stats = sampleStats(differences);
      if (!stats) return { kind: 'error', code: 1, text: '#DIV/0!' };
      t = stats.mean / Math.sqrt(stats.variance / differences.length);
      degrees = differences.length - 1;
    } else {
      const leftStats = sampleStats(leftValues);
      const rightStats = sampleStats(rightValues);
      if (!leftStats || !rightStats) return { kind: 'error', code: 1, text: '#DIV/0!' };
      if (type === 2) {
        degrees = leftValues.length + rightValues.length - 2;
        const pooledVariance =
          ((leftValues.length - 1) * leftStats.variance +
            (rightValues.length - 1) * rightStats.variance) /
          degrees;
        t =
          (leftStats.mean - rightStats.mean) /
          Math.sqrt(pooledVariance * (1 / leftValues.length + 1 / rightValues.length));
      } else {
        const leftComponent = leftStats.variance / leftValues.length;
        const rightComponent = rightStats.variance / rightValues.length;
        t = (leftStats.mean - rightStats.mean) / Math.sqrt(leftComponent + rightComponent);
        degrees =
          (leftComponent + rightComponent) ** 2 /
          (leftComponent ** 2 / (leftValues.length - 1) +
            rightComponent ** 2 / (rightValues.length - 1));
      }
    }
    const leftTail = studentTCdf(Math.abs(t), degrees);
    if (leftTail === null) return { kind: 'error', code: 6, text: '#NUM!' };
    const value = tails === 1 ? 1 - leftTail : 2 * (1 - leftTail);
    return Number.isFinite(value)
      ? { kind: 'number', value }
      : { kind: 'error', code: 6, text: '#NUM!' };
  };
  const chisqTestValue = (
    actualRange: FormulaRangeArg | ParsedA1Range,
    expectedRange: FormulaRangeArg | ParsedA1Range,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const actual = numericValuesInRangeWithShape(actualRange, rowOffset, colOffset);
    const expected = numericValuesInRangeWithShape(expectedRange, rowOffset, colOffset);
    if (
      !actual ||
      !expected ||
      actual.width !== expected.width ||
      actual.height !== expected.height
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const degrees = (actual.height - 1) * (actual.width - 1);
    if (degrees < 1) return { kind: 'error', code: 1, text: '#DIV/0!' };
    let statistic = 0;
    for (let index = 0; index < actual.values.length; index += 1) {
      const expectedValue = expected.values[index] as number;
      if (expectedValue <= 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
      const diff = (actual.values[index] as number) - expectedValue;
      statistic += (diff * diff) / expectedValue;
    }
    const leftTail = regularizedGammaP(degrees / 2, statistic / 2);
    if (leftTail === null) return { kind: 'error', code: 6, text: '#NUM!' };
    return { kind: 'number', value: 1 - leftTail };
  };
  const seriesSumValue = (
    xValue: CellValue,
    nValue: CellValue,
    mValue: CellValue,
    coefficients: FormulaAggregateArg[],
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const x = readNumber(xValue);
    const n = readNumber(nValue);
    const m = readNumber(mValue);
    const coefficientValues = numericValuesFromArgs(coefficients, rowOffset, colOffset);
    if (x === null || n === null || m === null || coefficientValues === null) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    if (coefficientValues.length === 0) return { kind: 'error', code: 15, text: '#VALUE!' };
    let total = 0;
    for (let index = 0; index < coefficientValues.length; index += 1) {
      total += (coefficientValues[index] as number) * x ** (n + index * m);
    }
    return Number.isFinite(total)
      ? { kind: 'number', value: total }
      : { kind: 'error', code: 6, text: '#NUM!' };
  };
  const sumProductRanges = (
    ranges: FormulaRangeArg[],
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    type Bounds = NonNullable<ReturnType<typeof formulaRangeArgBounds>>;
    const rawBounds = ranges.map((range) => formulaRangeArgBounds(range, rowOffset, colOffset));
    const bounds = rawBounds.filter((bound): bound is Bounds => bound !== null);
    const first = bounds[0];
    if (
      !first ||
      bounds.length !== ranges.length ||
      !bounds.every(
        (bound) =>
          validRangeBounds(bound) && bound.width === first.width && bound.height === first.height,
      )
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    let total = 0;
    for (let dr = 0; dr < first.height; dr += 1) {
      for (let dc = 0; dc < first.width; dc += 1) {
        let product = 1;
        for (const bound of bounds) {
          const value = state.data.cells.get(
            addrKey({ sheet, row: bound.r0 + dr, col: bound.c0 + dc }),
          )?.value;
          product *= value?.kind === 'number' && Number.isFinite(value.value) ? value.value : 0;
        }
        total += product;
      }
    }
    return { kind: 'number', value: total };
  };
  return {
    rankedRangeValue,
    percentileRangeValue,
    rankRangeValue,
    pairedRangeStatValue,
    regressionForecastValue,
    probabilityRangeValue,
    zTestValue,
    tTestValue,
    chisqTestValue,
    seriesSumValue,
    sumProductRanges,
  };
}
