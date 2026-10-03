import { splitFormulaArgs, splitFormulaArgsAllowEmpty } from '../splitter.js';
import type { FormulaAggregateArg, FormulaRangeArg } from '../types.js';
import type { FormulaCallParser, FormulaParserContext } from './context.js';

/** Exactly the parser entry points this family re-enters. */
export type StatisticsCallParserContext = Pick<
  FormulaParserContext,
  'parseFormulaOperand' | 'parseFormulaRangeArg' | 'parseFormulaAggregateArg'
>;

/** Aggregate and statistical calls: SUM-style aggregates, SUBTOTAL, AGGREGATE,
 *  LARGE/SMALL, percentiles and ranks, paired-range statistics, FORECAST, PROB,
 *  Z/T/CHISQ tests, SERIESSUM and SUMPRODUCT. */
export function createStatisticsCallParser(ctx: StatisticsCallParserContext): FormulaCallParser {
  const { parseFormulaOperand, parseFormulaRangeArg, parseFormulaAggregateArg } = ctx;
  return (fn, aggregate, sheetIndex) => {
    if (
      fn === 'SUM' ||
      fn === 'AVERAGE' ||
      fn === 'AVERAGEA' ||
      fn === 'MIN' ||
      fn === 'MINA' ||
      fn === 'MAX' ||
      fn === 'MAXA' ||
      fn === 'COUNT' ||
      fn === 'COUNTA' ||
      fn === 'COUNTBLANK' ||
      fn === 'PRODUCT' ||
      fn === 'MEDIAN' ||
      fn === 'MODE' ||
      fn === 'MODE.SNGL' ||
      fn === 'AVEDEV' ||
      fn === 'DEVSQ' ||
      fn === 'SKEW' ||
      fn === 'SKEW.P' ||
      fn === 'KURT' ||
      fn === 'GEOMEAN' ||
      fn === 'HARMEAN' ||
      fn === 'STDEV' ||
      fn === 'STDEVP' ||
      fn === 'STDEV.S' ||
      fn === 'STDEV.P' ||
      fn === 'VAR' ||
      fn === 'VARP' ||
      fn === 'VAR.S' ||
      fn === 'VAR.P'
    ) {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 1) {
        const range = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        if (range) return { kind: 'range-aggregate', fn, range };
      }
      if (args && args.length > 0) {
        const aggregateArgs: FormulaAggregateArg[] = [];
        for (const arg of args) {
          const aggregateArg = parseFormulaAggregateArg(arg, sheetIndex);
          if (!aggregateArg) return null;
          aggregateArgs.push(aggregateArg);
        }
        return { kind: 'aggregate-args', fn, args: aggregateArgs };
      }
    }
    if (fn === 'SUBTOTAL') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args && args.length >= 2) {
        const functionNum = parseFormulaOperand(args[0] ?? '', sheetIndex);
        if (!functionNum) return null;
        const subtotalArgs: FormulaAggregateArg[] = [];
        for (const arg of args.slice(1)) {
          const subtotalArg = parseFormulaAggregateArg(arg, sheetIndex);
          if (!subtotalArg) return null;
          subtotalArgs.push(subtotalArg);
        }
        return { kind: 'subtotal', functionNum, args: subtotalArgs };
      }
    }
    if (fn === 'AGGREGATE') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args && args.length >= 3) {
        const functionNum = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const options = parseFormulaOperand(args[1] ?? '', sheetIndex);
        if (!functionNum || !options) return null;
        const aggregateArgs: FormulaAggregateArg[] = [];
        for (const arg of args.slice(2)) {
          const aggregateArg = parseFormulaAggregateArg(arg, sheetIndex);
          if (!aggregateArg) return null;
          aggregateArgs.push(aggregateArg);
        }
        return { kind: 'aggregate-function', functionNum, options, args: aggregateArgs };
      }
    }
    if (fn === 'LARGE' || fn === 'SMALL') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2) {
        const range = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        const rank = parseFormulaOperand(args[1] ?? '', sheetIndex);
        if (range && rank) return { kind: 'ranked-range', fn, range, rank };
      }
    }
    if (
      fn === 'PERCENTILE.INC' ||
      fn === 'PERCENTILE.EXC' ||
      fn === 'PERCENTILE' ||
      fn === 'QUARTILE.INC' ||
      fn === 'QUARTILE.EXC' ||
      fn === 'QUARTILE'
    ) {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2) {
        const range = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        const value = parseFormulaOperand(args[1] ?? '', sheetIndex);
        if (range && value) return { kind: 'percentile-range', fn, range, value };
      }
    }
    if (fn === 'PERCENTRANK' || fn === 'PERCENTRANK.INC' || fn === 'PERCENTRANK.EXC') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args?.length === 2 || args?.length === 3) {
        const range = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        const value = parseFormulaOperand(args[1] ?? '', sheetIndex);
        if (range && value) {
          if (args.length === 2 || (args[2] ?? '').trim() === '') {
            return { kind: 'percentile-range', fn, range, value };
          }
          const significance = parseFormulaOperand(args[2] ?? '', sheetIndex);
          if (significance) return { kind: 'percentile-range', fn, range, value, significance };
        }
      }
    }
    if (fn === 'RANK' || fn === 'RANK.EQ' || fn === 'RANK.AVG') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args?.length === 2 || args?.length === 3) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const range = parseFormulaRangeArg(args[1] ?? '', sheetIndex);
        if (value && range) {
          if (args.length === 2 || (args[2] ?? '').trim() === '') {
            return { kind: 'range-rank', fn, value, range };
          }
          const order = parseFormulaOperand(args[2] ?? '', sheetIndex);
          if (order) return { kind: 'range-rank', fn, value, range, order };
        }
      }
    }
    if (
      fn === 'CORREL' ||
      fn === 'PEARSON' ||
      fn === 'COVAR' ||
      fn === 'COVARIANCE.P' ||
      fn === 'COVARIANCE.S' ||
      fn === 'SLOPE' ||
      fn === 'INTERCEPT' ||
      fn === 'RSQ' ||
      fn === 'STEYX' ||
      fn === 'SUMX2MY2' ||
      fn === 'SUMX2PY2' ||
      fn === 'SUMXMY2' ||
      fn === 'F.TEST' ||
      fn === 'FTEST'
    ) {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2) {
        const left = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        const right = parseFormulaRangeArg(args[1] ?? '', sheetIndex);
        if (left && right) return { kind: 'paired-range-stat', fn, left, right };
      }
    }
    if (fn === 'FORECAST' || fn === 'FORECAST.LINEAR') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 3) {
        const x = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const knownY = parseFormulaRangeArg(args[1] ?? '', sheetIndex);
        const knownX = parseFormulaRangeArg(args[2] ?? '', sheetIndex);
        if (x && knownY && knownX) return { kind: 'regression-forecast', fn, x, knownY, knownX };
      }
    }
    if (fn === 'PROB') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 3 || args?.length === 4) {
        const values = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        const probabilities = parseFormulaRangeArg(args[1] ?? '', sheetIndex);
        const lower = parseFormulaOperand(args[2] ?? '', sheetIndex);
        if (values && probabilities && lower) {
          if (args.length === 3) return { kind: 'probability-range', values, probabilities, lower };
          const upper = parseFormulaOperand(args[3] ?? '', sheetIndex);
          if (upper) return { kind: 'probability-range', values, probabilities, lower, upper };
        }
      }
    }
    if (fn === 'Z.TEST' || fn === 'ZTEST') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2 || args?.length === 3) {
        const range = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        const x = parseFormulaOperand(args[1] ?? '', sheetIndex);
        if (range && x) {
          if (args.length === 2) return { kind: 'z-test', range, x };
          const sigma = parseFormulaOperand(args[2] ?? '', sheetIndex);
          if (sigma) return { kind: 'z-test', range, x, sigma };
        }
      }
    }
    if (fn === 'T.TEST' || fn === 'TTEST') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 4) {
        const left = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        const right = parseFormulaRangeArg(args[1] ?? '', sheetIndex);
        const tails = parseFormulaOperand(args[2] ?? '', sheetIndex);
        const type = parseFormulaOperand(args[3] ?? '', sheetIndex);
        if (left && right && tails && type) return { kind: 't-test', left, right, tails, type };
      }
    }
    if (fn === 'CHISQ.TEST' || fn === 'CHITEST') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2) {
        const actual = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        const expected = parseFormulaRangeArg(args[1] ?? '', sheetIndex);
        if (actual && expected) return { kind: 'chisq-test', actual, expected };
      }
    }
    if (fn === 'SERIESSUM') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 4) {
        const x = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const n = parseFormulaOperand(args[1] ?? '', sheetIndex);
        const m = parseFormulaOperand(args[2] ?? '', sheetIndex);
        const coefficientsArg = parseFormulaAggregateArg(args[3] ?? '', sheetIndex);
        if (x && n && m && coefficientsArg) {
          return {
            kind: 'series-sum',
            x,
            n,
            m,
            coefficients: [coefficientsArg],
          };
        }
      }
    }
    if (fn === 'SUMPRODUCT') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args && args.length > 0) {
        const ranges = args.map((arg) => parseFormulaRangeArg(arg, sheetIndex));
        if (ranges.every((range) => range !== null)) {
          return { kind: 'sumproduct', ranges: ranges as FormulaRangeArg[] };
        }
      }
    }
    return undefined;
  };
}
