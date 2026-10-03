import {
  colToLetters,
  parseA1Range,
  parseA1Ref,
  parseR1C1Range,
  parseR1C1Ref,
} from './references.js';
import {
  splitFormulaArgs,
  splitFormulaArgsAllowEmpty,
  splitFormulaArithmetic,
  splitFormulaComparison,
  stripOuterParens,
} from './splitter.js';
import type {
  FormulaAggregateArg,
  FormulaCondition,
  FormulaDateArg,
  FormulaOperand,
  FormulaRangeArg,
  FormulaRangeOperand,
} from './types.js';

const MAX_FORMULA_AGGREGATE_CELLS = 10000;
const FORMULA_NUMBER_LITERAL = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;
const FORMULA_VALUE_NUMBER_LITERAL =
  /^[+-]?(?:(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?%?$/;

function parseFormulaRangeOperand(raw: string, sheetIndex: number): FormulaRangeOperand | null {
  const body = stripOuterParens(raw.trim());
  const aggregate = body.match(/^([A-Za-z][A-Za-z0-9.]*)\s*\((.*)\)$/);
  if (!aggregate) return null;
  const fn = (aggregate[1] ?? '').toUpperCase();
  if (fn === 'OFFSET') {
    const args = splitFormulaArgs(aggregate[2] ?? '');
    if (args && args.length >= 3 && args.length <= 5) {
      const reference = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
      const rows = parseFormulaOperand(args[1] ?? '', sheetIndex);
      const cols = parseFormulaOperand(args[2] ?? '', sheetIndex);
      if (reference && rows && cols) {
        const height =
          args.length >= 4 ? parseFormulaOperand(args[3] ?? '', sheetIndex) : undefined;
        if (args.length >= 4 && !height) return null;
        const width = args.length >= 5 ? parseFormulaOperand(args[4] ?? '', sheetIndex) : undefined;
        if (args.length >= 5 && !width) return null;
        return {
          kind: 'offset-range',
          reference,
          rows,
          cols,
          ...(height ? { height } : {}),
          ...(width ? { width } : {}),
        };
      }
    }
  }
  if (fn === 'INDIRECT') {
    const args = splitFormulaArgs(aggregate[2] ?? '');
    if (args?.length === 1 || args?.length === 2) {
      const refText = parseFormulaOperand(args[0] ?? '', sheetIndex);
      if (!refText) return null;
      if (args.length === 1) return { kind: 'indirect-range', refText };
      const a1 = parseFormulaOperand(args[1] ?? '', sheetIndex);
      if (a1) return { kind: 'indirect-range', refText, a1 };
    }
  }
  return null;
}

function parseFormulaAggregateArg(raw: string, sheetIndex: number): FormulaAggregateArg | null {
  const rangeArg = parseFormulaRangeArg(raw, sheetIndex);
  if (rangeArg) return rangeArg;
  const operand = parseFormulaOperand(raw, sheetIndex);
  return operand ? { kind: 'operand', operand } : null;
}

function parseFormulaRangeArg(raw: string, sheetIndex: number): FormulaRangeArg | null {
  const range = parseA1Range(raw, sheetIndex);
  if (range) return { kind: 'range', range };
  const dynamicRange = parseFormulaRangeOperand(raw, sheetIndex);
  if (dynamicRange) return { kind: 'dynamic-range', range: dynamicRange };
  return null;
}

function parseFormulaOperand(raw: string, sheetIndex: number): FormulaOperand | null {
  const body = stripOuterParens(raw.trim());
  const ref = parseA1Ref(body, sheetIndex);
  if (ref) return { kind: 'ref', ref };
  const aggregate = body.match(/^([A-Za-z][A-Za-z0-9.]*)\s*\((.*)\)$/);
  if (aggregate) {
    const fn = (aggregate[1] ?? '').toUpperCase();
    if ((fn === 'TRUE' || fn === 'FALSE') && (aggregate[2] ?? '').trim() === '') {
      return { kind: 'literal', value: { kind: 'bool', value: fn === 'TRUE' } };
    }
    if (fn === 'AND' || fn === 'OR' || fn === 'NOT' || fn === 'XOR') {
      const condition = parseFormulaCondition(body, sheetIndex);
      if (condition) return { kind: 'condition-value', condition };
    }
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
    if (fn === 'FVSCHEDULE') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2) {
        const principal = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const schedule = parseFormulaRangeArg(args[1] ?? '', sheetIndex);
        if (principal && schedule) return { kind: 'fv-schedule', principal, schedule };
      }
    }
    if (fn === 'NPV') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args && args.length >= 2) {
        const rate = parseFormulaOperand(args[0] ?? '', sheetIndex);
        if (!rate) return null;
        const values: FormulaAggregateArg[] = [];
        for (const arg of args.slice(1)) {
          const valueArg = parseFormulaAggregateArg(arg, sheetIndex);
          if (!valueArg) return null;
          values.push(valueArg);
        }
        return { kind: 'npv', rate, values };
      }
    }
    if (fn === 'MIRR') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 3) {
        const values = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        const financeRate = parseFormulaOperand(args[1] ?? '', sheetIndex);
        const reinvestRate = parseFormulaOperand(args[2] ?? '', sheetIndex);
        if (values && financeRate && reinvestRate) {
          return { kind: 'mirr', values, financeRate, reinvestRate };
        }
      }
    }
    if (fn === 'XNPV') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 3) {
        const rate = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const values = parseFormulaRangeArg(args[1] ?? '', sheetIndex);
        const dates = parseFormulaRangeArg(args[2] ?? '', sheetIndex);
        if (rate && values && dates) return { kind: 'xnpv', rate, values, dates };
      }
    }
    if (fn === 'XIRR') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args?.length === 2 || args?.length === 3) {
        const values = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        const dates = parseFormulaRangeArg(args[1] ?? '', sheetIndex);
        if (values && dates) {
          if (args.length === 2 || (args[2] ?? '').trim() === '') {
            return { kind: 'xirr', values, dates };
          }
          const guess = parseFormulaOperand(args[2] ?? '', sheetIndex);
          if (guess) return { kind: 'xirr', values, dates, guess };
        }
      }
    }
    if (fn === 'IRR') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args?.length === 1 || args?.length === 2) {
        const values = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        if (values) {
          if (args.length === 1 || (args[1] ?? '').trim() === '') {
            return { kind: 'irr', values };
          }
          const guess = parseFormulaOperand(args[1] ?? '', sheetIndex);
          if (guess) return { kind: 'irr', values, guess };
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
    if (fn === 'COUNTIF') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2) {
        const range = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        const criteria = parseFormulaOperand(args[1] ?? '', sheetIndex);
        if (range && criteria) return { kind: 'countif', range, criteria };
      }
    }
    if (fn === 'COUNTIFS') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args && args.length >= 2 && args.length % 2 === 0) {
        const pairs: { range: FormulaRangeArg; criteria: FormulaOperand }[] = [];
        for (let i = 0; i < args.length; i += 2) {
          const range = parseFormulaRangeArg(args[i] ?? '', sheetIndex);
          const criteria = parseFormulaOperand(args[i + 1] ?? '', sheetIndex);
          if (!range || !criteria) return null;
          pairs.push({ range, criteria });
        }
        return { kind: 'countifs', pairs };
      }
    }
    if (fn === 'SUMIF') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2 || args?.length === 3) {
        const range = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        const criteria = parseFormulaOperand(args[1] ?? '', sheetIndex);
        const sumRange = parseFormulaRangeArg(args[2] ?? args[0] ?? '', sheetIndex);
        if (range && criteria && sumRange) return { kind: 'sumif', range, criteria, sumRange };
      }
    }
    if (fn === 'AVERAGEIF') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2 || args?.length === 3) {
        const range = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        const criteria = parseFormulaOperand(args[1] ?? '', sheetIndex);
        const averageRange = parseFormulaRangeArg(args[2] ?? args[0] ?? '', sheetIndex);
        if (range && criteria && averageRange) {
          return { kind: 'averageif', range, criteria, averageRange };
        }
      }
    }
    if (fn === 'SUMIFS') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args && args.length >= 3 && args.length % 2 === 1) {
        const sumRange = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        if (!sumRange) return null;
        const pairs: { range: FormulaRangeArg; criteria: FormulaOperand }[] = [];
        for (let i = 1; i < args.length; i += 2) {
          const range = parseFormulaRangeArg(args[i] ?? '', sheetIndex);
          const criteria = parseFormulaOperand(args[i + 1] ?? '', sheetIndex);
          if (!range || !criteria) return null;
          pairs.push({ range, criteria });
        }
        return { kind: 'sumifs', sumRange, pairs };
      }
    }
    if (fn === 'AVERAGEIFS') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args && args.length >= 3 && args.length % 2 === 1) {
        const averageRange = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        if (!averageRange) return null;
        const pairs: { range: FormulaRangeArg; criteria: FormulaOperand }[] = [];
        for (let i = 1; i < args.length; i += 2) {
          const range = parseFormulaRangeArg(args[i] ?? '', sheetIndex);
          const criteria = parseFormulaOperand(args[i + 1] ?? '', sheetIndex);
          if (!range || !criteria) return null;
          pairs.push({ range, criteria });
        }
        return { kind: 'averageifs', averageRange, pairs };
      }
    }
    if (fn === 'MINIFS' || fn === 'MAXIFS') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args && args.length >= 3 && args.length % 2 === 1) {
        const valueRange = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        if (!valueRange) return null;
        const pairs: { range: FormulaRangeArg; criteria: FormulaOperand }[] = [];
        for (let i = 1; i < args.length; i += 2) {
          const range = parseFormulaRangeArg(args[i] ?? '', sheetIndex);
          const criteria = parseFormulaOperand(args[i + 1] ?? '', sheetIndex);
          if (!range || !criteria) return null;
          pairs.push({ range, criteria });
        }
        return { kind: 'minmaxifs', fn, valueRange, pairs };
      }
    }
    if (fn === 'LEN') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 1) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        if (value) return { kind: 'text-length', value };
      }
    }
    if (fn === 'FORMULATEXT') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 1) {
        const formulaRef = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        if (formulaRef) return { kind: 'formula-text', ref: formulaRef };
      }
    }
    if (fn === 'CELL') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 1 || args?.length === 2) {
        const infoType = parseFormulaOperand(args[0] ?? '', sheetIndex);
        if (!infoType) return null;
        if (args.length === 1) return { kind: 'cell-info', infoType };
        const ref = parseFormulaRangeArg(args[1] ?? '', sheetIndex);
        if (ref) return { kind: 'cell-info', infoType, ref };
      }
    }
    if (fn === 'SHEET' || fn === 'SHEETS') {
      const rawArgs = aggregate[2] ?? '';
      if (rawArgs.trim() === '' && fn === 'SHEET') return { kind: 'sheet-info', fn };
      const args = splitFormulaArgs(rawArgs);
      if (args?.length === 1) {
        const range = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        if (range) return { kind: 'sheet-info', fn, range };
      }
    }
    if (fn === 'HYPERLINK') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 1 || args?.length === 2) {
        const link = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const friendlyName =
          args.length === 2 ? parseFormulaOperand(args[1] ?? '', sheetIndex) : undefined;
        if (link && friendlyName !== null) return { kind: 'hyperlink', link, friendlyName };
      }
    }
    if (fn === 'SEARCH' || fn === 'FIND') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args?.length === 2 || args?.length === 3) {
        const needle = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const haystack = parseFormulaOperand(args[1] ?? '', sheetIndex);
        if (needle && haystack) {
          if (args.length === 2) return { kind: 'text-search', fn, needle, haystack };
          const start =
            (args[2] ?? '').trim() === ''
              ? { kind: 'literal' as const, value: { kind: 'number' as const, value: 1 } }
              : parseFormulaOperand(args[2] ?? '', sheetIndex);
          if (start) return { kind: 'text-search', fn, needle, haystack, start };
        }
      }
    }
    if (fn === 'LEFT' || fn === 'RIGHT') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args?.length === 1 || args?.length === 2) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const count =
          (args[1] ?? '').trim() === ''
            ? { kind: 'literal' as const, value: { kind: 'number' as const, value: 1 } }
            : parseFormulaOperand(args[1] ?? '1', sheetIndex);
        if (value && count) return { kind: 'text-slice', fn, value, count };
      }
    }
    if (fn === 'MID') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 3) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const start = parseFormulaOperand(args[1] ?? '', sheetIndex);
        const count = parseFormulaOperand(args[2] ?? '', sheetIndex);
        if (value && start && count) return { kind: 'text-slice', fn, value, start, count };
      }
    }
    if (fn === 'CONCATENATE' || fn === 'CONCAT') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args && args.length > 0) {
        const values = args.map((arg) =>
          arg.trim() === ''
            ? { kind: 'literal' as const, value: { kind: 'blank' as const } }
            : parseFormulaOperand(arg, sheetIndex),
        );
        if (values.every((value) => value !== null)) {
          return { kind: 'text-concat-function', values: values as FormulaOperand[] };
        }
      }
    }
    if (fn === 'SUBSTITUTE') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args?.length === 3 || args?.length === 4) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const oldText = parseFormulaOperand(args[1] ?? '', sheetIndex);
        const newText = parseFormulaOperand(args[2] ?? '', sheetIndex);
        if (value && oldText && newText) {
          if (args.length === 3 || (args[3] ?? '').trim() === '') {
            return { kind: 'text-substitute', value, oldText, newText };
          }
          const instance = parseFormulaOperand(args[3] ?? '', sheetIndex);
          if (instance) return { kind: 'text-substitute', value, oldText, newText, instance };
        }
      }
    }
    if (fn === 'REPLACE') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 4) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const start = parseFormulaOperand(args[1] ?? '', sheetIndex);
        const count = parseFormulaOperand(args[2] ?? '', sheetIndex);
        const newText = parseFormulaOperand(args[3] ?? '', sheetIndex);
        if (value && start && count && newText) {
          return { kind: 'text-replace', value, start, count, newText };
        }
      }
    }
    if (fn === 'REPT') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const count = parseFormulaOperand(args[1] ?? '', sheetIndex);
        if (value && count) return { kind: 'text-repeat', value, count };
      }
    }
    if (fn === 'TEXTBEFORE' || fn === 'TEXTAFTER') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (
        args &&
        args.length >= 2 &&
        args.length <= 6 &&
        (args[0] ?? '').trim() !== '' &&
        (args[1] ?? '').trim() !== ''
      ) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const delimiter = parseFormulaOperand(args[1] ?? '', sheetIndex);
        const instance =
          args.length >= 3 && (args[2] ?? '').trim() !== ''
            ? parseFormulaOperand(args[2] ?? '', sheetIndex)
            : undefined;
        const matchMode =
          args.length >= 4 && (args[3] ?? '').trim() !== ''
            ? parseFormulaOperand(args[3] ?? '', sheetIndex)
            : undefined;
        const matchEnd =
          args.length >= 5 && (args[4] ?? '').trim() !== ''
            ? parseFormulaOperand(args[4] ?? '', sheetIndex)
            : undefined;
        const ifNotFound =
          args.length >= 6 && (args[5] ?? '').trim() !== ''
            ? parseFormulaOperand(args[5] ?? '', sheetIndex)
            : undefined;
        if (
          value &&
          delimiter &&
          instance !== null &&
          matchMode !== null &&
          matchEnd !== null &&
          ifNotFound !== null
        ) {
          return {
            kind: 'text-before-after',
            fn,
            value,
            delimiter,
            instance,
            matchMode,
            matchEnd,
            ifNotFound,
          };
        }
      }
    }
    if (fn === 'TEXTJOIN') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args && args.length >= 3) {
        const delimiter = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const ignoreEmpty = parseFormulaOperand(args[1] ?? '', sheetIndex);
        const values = args
          .slice(2)
          .map((arg) =>
            arg.trim() === ''
              ? { kind: 'literal' as const, value: { kind: 'blank' as const } }
              : parseFormulaOperand(arg, sheetIndex),
          );
        if (delimiter && ignoreEmpty && values.every((value) => value !== null)) {
          return {
            kind: 'text-join',
            delimiter,
            ignoreEmpty,
            values: values as FormulaOperand[],
          };
        }
      }
    }
    if (
      fn === 'LOWER' ||
      fn === 'UPPER' ||
      fn === 'TRIM' ||
      fn === 'CLEAN' ||
      fn === 'PROPER' ||
      fn === 'ENCODEURL'
    ) {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 1) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        if (value) return { kind: 'text-transform', fn, value };
      }
    }
    if (fn === 'EXACT') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2) {
        const left = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const right = parseFormulaOperand(args[1] ?? '', sheetIndex);
        if (left && right) return { kind: 'text-exact', left, right };
      }
    }
    if (fn === 'TEXT') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const pattern = parseFormulaOperand(args[1] ?? '', sheetIndex);
        if (value && pattern) return { kind: 'text-format', value, pattern };
      }
    }
    if (fn === 'DOLLAR' || fn === 'FIXED') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (
        args &&
        args.length >= 1 &&
        args.length <= (fn === 'DOLLAR' ? 2 : 3) &&
        (args[0] ?? '').trim() !== ''
      ) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const decimals =
          args.length >= 2 && (args[1] ?? '').trim() !== ''
            ? parseFormulaOperand(args[1] ?? '', sheetIndex)
            : undefined;
        const noCommas =
          fn === 'FIXED' && args.length >= 3 && (args[2] ?? '').trim() !== ''
            ? parseFormulaOperand(args[2] ?? '', sheetIndex)
            : undefined;
        if (value && decimals !== null && noCommas !== null) {
          return { kind: 'text-fixed-format', fn, value, decimals, noCommas };
        }
      }
    }
    if (fn === 'VALUE') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 1) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        if (value) return { kind: 'text-value', value };
      }
    }
    if (fn === 'VALUETOTEXT') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args && args.length >= 1 && args.length <= 2 && (args[0] ?? '').trim() !== '') {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const format =
          args.length >= 2 && (args[1] ?? '').trim() !== ''
            ? parseFormulaOperand(args[1] ?? '', sheetIndex)
            : undefined;
        if (value && format !== null) return { kind: 'value-to-text', value, format };
      }
    }
    if (fn === 'NUMBERVALUE') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args && args.length >= 1 && args.length <= 3 && (args[0] ?? '').trim() !== '') {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const decimalSeparator =
          args.length >= 2 && (args[1] ?? '').trim() !== ''
            ? parseFormulaOperand(args[1] ?? '', sheetIndex)
            : undefined;
        const groupSeparator =
          args.length >= 3 && (args[2] ?? '').trim() !== ''
            ? parseFormulaOperand(args[2] ?? '', sheetIndex)
            : undefined;
        if (value && decimalSeparator !== null && groupSeparator !== null) {
          return { kind: 'text-number-value', value, decimalSeparator, groupSeparator };
        }
      }
    }
    if (fn === 'N' || fn === 'T') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 1) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        if (value) return { kind: 'scalar-coerce', fn, value };
      }
    }
    if (fn === 'ROW' || fn === 'COLUMN') {
      const rawArgs = aggregate[2] ?? '';
      if (rawArgs.trim() === '') return { kind: 'position', fn };
      const args = splitFormulaArgs(rawArgs);
      if (args?.length === 1) {
        const ref = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        if (ref) return { kind: 'position', fn, ref };
      }
    }
    if (fn === 'ROWS' || fn === 'COLUMNS' || fn === 'AREAS') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 1) {
        const range = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        if (range) return { kind: 'range-dimension', fn, range };
      }
    }
    if (
      fn === 'ABS' ||
      fn === 'MOD' ||
      fn === 'ROUND' ||
      fn === 'ROUNDUP' ||
      fn === 'ROUNDDOWN' ||
      fn === 'MROUND' ||
      fn === 'QUOTIENT' ||
      fn === 'INT' ||
      fn === 'TRUNC' ||
      fn === 'SQRT' ||
      fn === 'POWER' ||
      fn === 'PI' ||
      fn === 'RADIANS' ||
      fn === 'DEGREES' ||
      fn === 'SIN' ||
      fn === 'COS' ||
      fn === 'TAN' ||
      fn === 'SEC' ||
      fn === 'CSC' ||
      fn === 'COT' ||
      fn === 'ASIN' ||
      fn === 'ACOS' ||
      fn === 'ATAN' ||
      fn === 'ATAN2' ||
      fn === 'ACOT' ||
      fn === 'SINH' ||
      fn === 'COSH' ||
      fn === 'TANH' ||
      fn === 'COTH' ||
      fn === 'SECH' ||
      fn === 'CSCH' ||
      fn === 'ASINH' ||
      fn === 'ACOSH' ||
      fn === 'ATANH' ||
      fn === 'ACOTH' ||
      fn === 'EXP' ||
      fn === 'LN' ||
      fn === 'LOG' ||
      fn === 'LOG10' ||
      fn === 'CHAR' ||
      fn === 'CODE' ||
      fn === 'UNICHAR' ||
      fn === 'UNICODE' ||
      fn === 'ADDRESS' ||
      fn === 'TYPE' ||
      fn === 'ERROR.TYPE' ||
      fn === 'FISHER' ||
      fn === 'FISHERINV' ||
      fn === 'ERF' ||
      fn === 'ERF.PRECISE' ||
      fn === 'ERFC' ||
      fn === 'ERFC.PRECISE' ||
      fn === 'GAUSS' ||
      fn === 'BASE' ||
      fn === 'DECIMAL' ||
      fn === 'BIN2DEC' ||
      fn === 'DEC2BIN' ||
      fn === 'HEX2DEC' ||
      fn === 'DEC2HEX' ||
      fn === 'OCT2DEC' ||
      fn === 'DEC2OCT' ||
      fn === 'BIN2HEX' ||
      fn === 'HEX2BIN' ||
      fn === 'BIN2OCT' ||
      fn === 'OCT2BIN' ||
      fn === 'HEX2OCT' ||
      fn === 'OCT2HEX' ||
      fn === 'ROMAN' ||
      fn === 'ARABIC' ||
      fn === 'DELTA' ||
      fn === 'GESTEP' ||
      fn === 'BITAND' ||
      fn === 'BITOR' ||
      fn === 'BITXOR' ||
      fn === 'BITLSHIFT' ||
      fn === 'BITRSHIFT' ||
      fn === 'SQRTPI' ||
      fn === 'SUMSQ' ||
      fn === 'SIGN' ||
      fn === 'GAMMA' ||
      fn === 'GAMMALN' ||
      fn === 'GAMMALN.PRECISE' ||
      fn === 'GCD' ||
      fn === 'LCM' ||
      fn === 'FACT' ||
      fn === 'FACTDOUBLE' ||
      fn === 'COMBIN' ||
      fn === 'COMBINA' ||
      fn === 'PERMUT' ||
      fn === 'PERMUTATIONA' ||
      fn === 'MULTINOMIAL' ||
      fn === 'EVEN' ||
      fn === 'ODD' ||
      fn === 'STANDARDIZE' ||
      fn === 'PHI' ||
      fn === 'CONFIDENCE' ||
      fn === 'CONFIDENCE.NORM' ||
      fn === 'CONFIDENCE.T' ||
      fn === 'PMT' ||
      fn === 'PV' ||
      fn === 'FV' ||
      fn === 'NPER' ||
      fn === 'RATE' ||
      fn === 'IPMT' ||
      fn === 'PPMT' ||
      fn === 'CUMIPMT' ||
      fn === 'CUMPRINC' ||
      fn === 'ISPMT' ||
      fn === 'EFFECT' ||
      fn === 'NOMINAL' ||
      fn === 'DOLLARDE' ||
      fn === 'DOLLARFR' ||
      fn === 'DISC' ||
      fn === 'INTRATE' ||
      fn === 'PRICEDISC' ||
      fn === 'RECEIVED' ||
      fn === 'ACCRINTM' ||
      fn === 'TBILLPRICE' ||
      fn === 'TBILLYIELD' ||
      fn === 'TBILLEQ' ||
      fn === 'RRI' ||
      fn === 'PDURATION' ||
      fn === 'SLN' ||
      fn === 'SYD' ||
      fn === 'DDB' ||
      fn === 'DB' ||
      fn === 'NORMSDIST' ||
      fn === 'NORMDIST' ||
      fn === 'NORM.S.DIST' ||
      fn === 'NORM.DIST' ||
      fn === 'NORMSINV' ||
      fn === 'NORM.S.INV' ||
      fn === 'NORMINV' ||
      fn === 'NORM.INV' ||
      fn === 'LOGINV' ||
      fn === 'LOGNORM.INV' ||
      fn === 'LOGNORMDIST' ||
      fn === 'LOGNORM.DIST' ||
      fn === 'GAMMADIST' ||
      fn === 'GAMMA.DIST' ||
      fn === 'GAMMAINV' ||
      fn === 'GAMMA.INV' ||
      fn === 'BETADIST' ||
      fn === 'BETA.DIST' ||
      fn === 'BETAINV' ||
      fn === 'BETA.INV' ||
      fn === 'FDIST' ||
      fn === 'F.DIST' ||
      fn === 'F.DIST.RT' ||
      fn === 'FINV' ||
      fn === 'F.INV' ||
      fn === 'F.INV.RT' ||
      fn === 'TDIST' ||
      fn === 'T.DIST' ||
      fn === 'T.DIST.2T' ||
      fn === 'T.DIST.RT' ||
      fn === 'TINV' ||
      fn === 'T.INV' ||
      fn === 'T.INV.2T' ||
      fn === 'CHIDIST' ||
      fn === 'CHISQ.DIST' ||
      fn === 'CHISQ.DIST.RT' ||
      fn === 'CHIINV' ||
      fn === 'CHISQ.INV' ||
      fn === 'CHISQ.INV.RT' ||
      fn === 'WEIBULL' ||
      fn === 'WEIBULL.DIST' ||
      fn === 'BINOMDIST' ||
      fn === 'BINOM.DIST' ||
      fn === 'CRITBINOM' ||
      fn === 'BINOM.INV' ||
      fn === 'NEGBINOMDIST' ||
      fn === 'NEGBINOM.DIST' ||
      fn === 'HYPGEOMDIST' ||
      fn === 'HYPGEOM.DIST' ||
      fn === 'POISSON' ||
      fn === 'POISSON.DIST' ||
      fn === 'EXPONDIST' ||
      fn === 'EXPON.DIST' ||
      fn === 'CEILING' ||
      fn === 'FLOOR' ||
      fn === 'CEILING.MATH' ||
      fn === 'FLOOR.MATH' ||
      fn === 'CEILING.PRECISE' ||
      fn === 'FLOOR.PRECISE' ||
      fn === 'ISO.CEILING'
    ) {
      if (fn === 'PI' && (aggregate[2] ?? '').trim() === '') {
        return { kind: 'numeric-function', fn, args: [] };
      }
      const args =
        fn === 'CEILING.MATH' ||
        fn === 'CEILING.PRECISE' ||
        fn === 'FLOOR.PRECISE' ||
        fn === 'ISO.CEILING' ||
        fn === 'FLOOR.MATH' ||
        fn === 'LOG' ||
        fn === 'TRUNC' ||
        fn === 'DELTA' ||
        fn === 'GESTEP' ||
        fn === 'PMT' ||
        fn === 'PV' ||
        fn === 'FV' ||
        fn === 'NPER' ||
        fn === 'RATE' ||
        fn === 'IPMT' ||
        fn === 'PPMT' ||
        fn === 'CUMIPMT' ||
        fn === 'CUMPRINC' ||
        fn === 'DISC' ||
        fn === 'INTRATE' ||
        fn === 'PRICEDISC' ||
        fn === 'RECEIVED' ||
        fn === 'ACCRINTM' ||
        fn === 'TBILLPRICE' ||
        fn === 'TBILLYIELD' ||
        fn === 'TBILLEQ' ||
        fn === 'DDB' ||
        fn === 'DB' ||
        fn === 'ADDRESS'
          ? splitFormulaArgsAllowEmpty(aggregate[2] ?? '')
          : splitFormulaArgs(aggregate[2] ?? '');
      const validLength =
        fn === 'ABS' ||
        fn === 'INT' ||
        fn === 'SQRT' ||
        fn === 'RADIANS' ||
        fn === 'DEGREES' ||
        fn === 'SIN' ||
        fn === 'COS' ||
        fn === 'TAN' ||
        fn === 'SEC' ||
        fn === 'CSC' ||
        fn === 'COT' ||
        fn === 'ASIN' ||
        fn === 'ACOS' ||
        fn === 'ATAN' ||
        fn === 'ACOT' ||
        fn === 'SINH' ||
        fn === 'COSH' ||
        fn === 'TANH' ||
        fn === 'COTH' ||
        fn === 'SECH' ||
        fn === 'CSCH' ||
        fn === 'ASINH' ||
        fn === 'ACOSH' ||
        fn === 'ATANH' ||
        fn === 'ACOTH' ||
        fn === 'EXP' ||
        fn === 'LN' ||
        fn === 'LOG10' ||
        fn === 'CHAR' ||
        fn === 'CODE' ||
        fn === 'UNICHAR' ||
        fn === 'UNICODE' ||
        fn === 'TYPE' ||
        fn === 'ERROR.TYPE' ||
        fn === 'FISHER' ||
        fn === 'FISHERINV' ||
        fn === 'ERF.PRECISE' ||
        fn === 'ERFC' ||
        fn === 'ERFC.PRECISE' ||
        fn === 'GAUSS' ||
        fn === 'ARABIC' ||
        fn === 'SQRTPI' ||
        fn === 'SIGN' ||
        fn === 'GAMMA' ||
        fn === 'GAMMALN' ||
        fn === 'GAMMALN.PRECISE' ||
        fn === 'FACT' ||
        fn === 'FACTDOUBLE' ||
        fn === 'EVEN' ||
        fn === 'ODD' ||
        fn === 'NORMSDIST' ||
        fn === 'PHI'
          ? args?.length === 1
          : fn === 'ADDRESS'
            ? args !== null && args.length >= 2 && args.length <= 5
            : fn === 'GCD' || fn === 'LCM' || fn === 'SUMSQ' || fn === 'MULTINOMIAL'
              ? args !== null && args.length > 0
              : fn === 'LOG'
                ? args?.length === 1 || args?.length === 2
                : fn === 'ERF'
                  ? args?.length === 1 || args?.length === 2
                  : fn === 'BASE'
                    ? args?.length === 2 || args?.length === 3
                    : fn === 'DEC2BIN' ||
                        fn === 'DEC2HEX' ||
                        fn === 'DEC2OCT' ||
                        fn === 'BIN2HEX' ||
                        fn === 'HEX2BIN' ||
                        fn === 'BIN2OCT' ||
                        fn === 'OCT2BIN' ||
                        fn === 'HEX2OCT' ||
                        fn === 'OCT2HEX'
                      ? args?.length === 1 || args?.length === 2
                      : fn === 'DECIMAL'
                        ? args?.length === 2
                        : fn === 'ROMAN'
                          ? args?.length === 1 || args?.length === 2
                          : fn === 'BIN2DEC' || fn === 'HEX2DEC' || fn === 'OCT2DEC'
                            ? args?.length === 1
                            : fn === 'DELTA' || fn === 'GESTEP'
                              ? args?.length === 1 || args?.length === 2
                              : fn === 'BITAND' ||
                                  fn === 'BITOR' ||
                                  fn === 'BITXOR' ||
                                  fn === 'BITLSHIFT' ||
                                  fn === 'BITRSHIFT'
                                ? args?.length === 2
                                : fn === 'STANDARDIZE'
                                  ? args?.length === 3
                                  : fn === 'RATE'
                                    ? args?.length === 3 ||
                                      args?.length === 4 ||
                                      args?.length === 5 ||
                                      args?.length === 6
                                    : fn === 'PMT' || fn === 'PV' || fn === 'FV' || fn === 'NPER'
                                      ? args?.length === 3 ||
                                        args?.length === 4 ||
                                        args?.length === 5
                                      : fn === 'IPMT' || fn === 'PPMT'
                                        ? args?.length === 4 ||
                                          args?.length === 5 ||
                                          args?.length === 6
                                        : fn === 'CUMIPMT' || fn === 'CUMPRINC'
                                          ? args?.length === 6
                                          : fn === 'ISPMT'
                                            ? args?.length === 4
                                            : fn === 'EFFECT' || fn === 'NOMINAL'
                                              ? args?.length === 2
                                              : fn === 'DOLLARDE' || fn === 'DOLLARFR'
                                                ? args?.length === 2
                                                : fn === 'DISC' ||
                                                    fn === 'INTRATE' ||
                                                    fn === 'PRICEDISC' ||
                                                    fn === 'RECEIVED'
                                                  ? args?.length === 4 || args?.length === 5
                                                  : fn === 'ACCRINTM'
                                                    ? args?.length === 3 ||
                                                      args?.length === 4 ||
                                                      args?.length === 5
                                                    : fn === 'TBILLPRICE' ||
                                                        fn === 'TBILLYIELD' ||
                                                        fn === 'TBILLEQ'
                                                      ? args?.length === 3
                                                      : fn === 'RRI' || fn === 'PDURATION'
                                                        ? args?.length === 3
                                                        : fn === 'SLN'
                                                          ? args?.length === 3
                                                          : fn === 'SYD'
                                                            ? args?.length === 4
                                                            : fn === 'DDB'
                                                              ? args?.length === 4 ||
                                                                args?.length === 5
                                                              : fn === 'DB'
                                                                ? args?.length === 4 ||
                                                                  args?.length === 5
                                                                : fn === 'NORM.S.DIST'
                                                                  ? args?.length === 2
                                                                  : fn === 'CONFIDENCE' ||
                                                                      fn === 'CONFIDENCE.NORM' ||
                                                                      fn === 'CONFIDENCE.T'
                                                                    ? args?.length === 3
                                                                    : fn === 'NORMSINV' ||
                                                                        fn === 'NORM.S.INV'
                                                                      ? args?.length === 1
                                                                      : fn === 'NORMINV' ||
                                                                          fn === 'NORM.INV' ||
                                                                          fn === 'LOGINV' ||
                                                                          fn === 'LOGNORM.INV'
                                                                        ? args?.length === 3
                                                                        : fn === 'NORMDIST' ||
                                                                            fn === 'NORM.DIST'
                                                                          ? args?.length === 4
                                                                          : fn === 'LOGNORMDIST'
                                                                            ? args?.length === 3
                                                                            : fn ===
                                                                                  'LOGNORM.DIST' ||
                                                                                fn ===
                                                                                  'GAMMADIST' ||
                                                                                fn ===
                                                                                  'GAMMA.DIST' ||
                                                                                fn === 'WEIBULL' ||
                                                                                fn ===
                                                                                  'WEIBULL.DIST'
                                                                              ? args?.length === 4
                                                                              : fn === 'GAMMAINV' ||
                                                                                  fn === 'GAMMA.INV'
                                                                                ? args?.length === 3
                                                                                : fn ===
                                                                                      'BETADIST' ||
                                                                                    fn === 'BETAINV'
                                                                                  ? args?.length ===
                                                                                      3 ||
                                                                                    args?.length ===
                                                                                      4 ||
                                                                                    args?.length ===
                                                                                      5
                                                                                  : fn ===
                                                                                      'BETA.DIST'
                                                                                    ? args?.length ===
                                                                                        4 ||
                                                                                      args?.length ===
                                                                                        5 ||
                                                                                      args?.length ===
                                                                                        6
                                                                                    : fn ===
                                                                                        'BETA.INV'
                                                                                      ? args?.length ===
                                                                                          3 ||
                                                                                        args?.length ===
                                                                                          4 ||
                                                                                        args?.length ===
                                                                                          5
                                                                                      : fn ===
                                                                                            'FDIST' ||
                                                                                          fn ===
                                                                                            'F.DIST.RT' ||
                                                                                          fn ===
                                                                                            'FINV' ||
                                                                                          fn ===
                                                                                            'F.INV' ||
                                                                                          fn ===
                                                                                            'F.INV.RT'
                                                                                        ? args?.length ===
                                                                                          3
                                                                                        : fn ===
                                                                                            'F.DIST'
                                                                                          ? args?.length ===
                                                                                            4
                                                                                          : fn ===
                                                                                                'TDIST' ||
                                                                                              fn ===
                                                                                                'T.DIST'
                                                                                            ? args?.length ===
                                                                                              3
                                                                                            : fn ===
                                                                                                  'T.DIST.2T' ||
                                                                                                fn ===
                                                                                                  'T.DIST.RT' ||
                                                                                                fn ===
                                                                                                  'TINV' ||
                                                                                                fn ===
                                                                                                  'T.INV' ||
                                                                                                fn ===
                                                                                                  'T.INV.2T'
                                                                                              ? args?.length ===
                                                                                                2
                                                                                              : fn ===
                                                                                                    'CHIDIST' ||
                                                                                                  fn ===
                                                                                                    'CHISQ.DIST.RT' ||
                                                                                                  fn ===
                                                                                                    'CHIINV' ||
                                                                                                  fn ===
                                                                                                    'CHISQ.INV' ||
                                                                                                  fn ===
                                                                                                    'CHISQ.INV.RT'
                                                                                                ? args?.length ===
                                                                                                  2
                                                                                                : fn ===
                                                                                                    'CHISQ.DIST'
                                                                                                  ? args?.length ===
                                                                                                    3
                                                                                                  : fn ===
                                                                                                        'BINOMDIST' ||
                                                                                                      fn ===
                                                                                                        'BINOM.DIST'
                                                                                                    ? args?.length ===
                                                                                                      4
                                                                                                    : fn ===
                                                                                                          'CRITBINOM' ||
                                                                                                        fn ===
                                                                                                          'BINOM.INV'
                                                                                                      ? args?.length ===
                                                                                                        3
                                                                                                      : fn ===
                                                                                                          'NEGBINOMDIST'
                                                                                                        ? args?.length ===
                                                                                                          3
                                                                                                        : fn ===
                                                                                                            'NEGBINOM.DIST'
                                                                                                          ? args?.length ===
                                                                                                            4
                                                                                                          : fn ===
                                                                                                              'HYPGEOMDIST'
                                                                                                            ? args?.length ===
                                                                                                              4
                                                                                                            : fn ===
                                                                                                                'HYPGEOM.DIST'
                                                                                                              ? args?.length ===
                                                                                                                5
                                                                                                              : fn ===
                                                                                                                    'POISSON' ||
                                                                                                                  fn ===
                                                                                                                    'POISSON.DIST' ||
                                                                                                                  fn ===
                                                                                                                    'EXPONDIST' ||
                                                                                                                  fn ===
                                                                                                                    'EXPON.DIST'
                                                                                                                ? args?.length ===
                                                                                                                  3
                                                                                                                : fn ===
                                                                                                                      'CEILING' ||
                                                                                                                    fn ===
                                                                                                                      'FLOOR' ||
                                                                                                                    fn ===
                                                                                                                      'MROUND' ||
                                                                                                                    fn ===
                                                                                                                      'QUOTIENT'
                                                                                                                  ? args?.length ===
                                                                                                                    2
                                                                                                                  : fn ===
                                                                                                                        'COMBIN' ||
                                                                                                                      fn ===
                                                                                                                        'COMBINA' ||
                                                                                                                      fn ===
                                                                                                                        'PERMUT' ||
                                                                                                                      fn ===
                                                                                                                        'PERMUTATIONA'
                                                                                                                    ? args?.length ===
                                                                                                                      2
                                                                                                                    : fn ===
                                                                                                                          'CEILING.MATH' ||
                                                                                                                        fn ===
                                                                                                                          'FLOOR.MATH'
                                                                                                                      ? args?.length ===
                                                                                                                          1 ||
                                                                                                                        args?.length ===
                                                                                                                          2 ||
                                                                                                                        args?.length ===
                                                                                                                          3
                                                                                                                      : fn ===
                                                                                                                            'CEILING.PRECISE' ||
                                                                                                                          fn ===
                                                                                                                            'FLOOR.PRECISE' ||
                                                                                                                          fn ===
                                                                                                                            'ISO.CEILING'
                                                                                                                        ? args?.length ===
                                                                                                                            1 ||
                                                                                                                          args?.length ===
                                                                                                                            2
                                                                                                                        : fn ===
                                                                                                                            'TRUNC'
                                                                                                                          ? args?.length ===
                                                                                                                              1 ||
                                                                                                                            args?.length ===
                                                                                                                              2
                                                                                                                          : args?.length ===
                                                                                                                            2;
      if (args && validLength) {
        if (
          (fn === 'CEILING.MATH' ||
            fn === 'CEILING.PRECISE' ||
            fn === 'FLOOR.PRECISE' ||
            fn === 'ISO.CEILING' ||
            fn === 'FLOOR.MATH' ||
            fn === 'LOG' ||
            fn === 'TRUNC' ||
            fn === 'DELTA' ||
            fn === 'GESTEP' ||
            fn === 'PMT' ||
            fn === 'PV' ||
            fn === 'FV' ||
            fn === 'NPER' ||
            fn === 'RATE' ||
            fn === 'IPMT' ||
            fn === 'PPMT' ||
            fn === 'CUMIPMT' ||
            fn === 'CUMPRINC' ||
            fn === 'DISC' ||
            fn === 'INTRATE' ||
            fn === 'PRICEDISC' ||
            fn === 'RECEIVED' ||
            fn === 'ACCRINTM' ||
            fn === 'TBILLPRICE' ||
            fn === 'TBILLYIELD' ||
            fn === 'TBILLEQ' ||
            fn === 'DDB' ||
            fn === 'DB' ||
            fn === 'ADDRESS') &&
          (args[0] ?? '').trim() === ''
        ) {
          return null;
        }
        const operands = args.map((arg, index) => {
          if (
            (fn === 'CEILING.MATH' ||
              fn === 'FLOOR.MATH' ||
              fn === 'CEILING.PRECISE' ||
              fn === 'FLOOR.PRECISE' ||
              fn === 'ISO.CEILING') &&
            arg.trim() === ''
          ) {
            const value =
              fn === 'CEILING.PRECISE' || fn === 'FLOOR.PRECISE' || fn === 'ISO.CEILING'
                ? 1
                : index === 1
                  ? 1
                  : 0;
            return { kind: 'literal' as const, value: { kind: 'number' as const, value } };
          }
          if (fn === 'LOG' && index === 1 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'number' as const, value: 10 } };
          }
          if (fn === 'TRUNC' && index === 1 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'number' as const, value: 0 } };
          }
          if ((fn === 'DELTA' || fn === 'GESTEP') && index === 1 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'number' as const, value: 0 } };
          }
          if (fn === 'ADDRESS' && index === 2 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'number' as const, value: 1 } };
          }
          if (fn === 'ADDRESS' && index === 3 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'bool' as const, value: true } };
          }
          if (fn === 'ADDRESS' && index === 4 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'text' as const, value: '' } };
          }
          if (
            arg.trim() === '' &&
            (((fn === 'PMT' || fn === 'PV' || fn === 'FV' || fn === 'NPER' || fn === 'RATE') &&
              index >= 3) ||
              ((fn === 'IPMT' || fn === 'PPMT') && index >= 4))
          ) {
            return {
              kind: 'literal' as const,
              value: { kind: 'number' as const, value: fn === 'RATE' && index === 5 ? 0.1 : 0 },
            };
          }
          if (fn === 'DDB' && index === 4 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'number' as const, value: 2 } };
          }
          if (fn === 'DB' && index === 4 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'number' as const, value: 12 } };
          }
          if (
            (fn === 'DISC' || fn === 'INTRATE' || fn === 'PRICEDISC' || fn === 'RECEIVED') &&
            index === 4 &&
            arg.trim() === ''
          ) {
            return { kind: 'literal' as const, value: { kind: 'number' as const, value: 0 } };
          }
          if (fn === 'ACCRINTM' && index === 3 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'number' as const, value: 1000 } };
          }
          if (fn === 'ACCRINTM' && index === 4 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'number' as const, value: 0 } };
          }
          return parseFormulaOperand(arg, sheetIndex);
        });
        if (operands.every((operand) => operand !== null)) {
          return { kind: 'numeric-function', fn, args: operands as FormulaOperand[] };
        }
      }
    }
    if (fn === 'ISEVEN' || fn === 'ISODD') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 1) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        if (value) return { kind: 'numeric-predicate', fn, value };
      }
    }
    if (
      fn === 'DATE' ||
      fn === 'YEAR' ||
      fn === 'MONTH' ||
      fn === 'DAY' ||
      fn === 'WEEKDAY' ||
      fn === 'WEEKNUM' ||
      fn === 'ISOWEEKNUM' ||
      fn === 'TODAY' ||
      fn === 'NOW' ||
      fn === 'TIME' ||
      fn === 'EDATE' ||
      fn === 'EOMONTH' ||
      fn === 'DAYS' ||
      fn === 'DAYS360' ||
      fn === 'DATEDIF' ||
      fn === 'YEARFRAC' ||
      fn === 'DATEVALUE' ||
      fn === 'TIMEVALUE' ||
      fn === 'NETWORKDAYS' ||
      fn === 'NETWORKDAYS.INTL' ||
      fn === 'WORKDAY' ||
      fn === 'WORKDAY.INTL' ||
      fn === 'HOUR' ||
      fn === 'MINUTE' ||
      fn === 'SECOND'
    ) {
      const args =
        fn === 'WEEKDAY' ||
        fn === 'WEEKNUM' ||
        fn === 'DAYS360' ||
        fn === 'YEARFRAC' ||
        fn === 'NETWORKDAYS.INTL' ||
        fn === 'WORKDAY.INTL'
          ? splitFormulaArgsAllowEmpty(aggregate[2] ?? '')
          : splitFormulaArgs(aggregate[2] ?? '');
      const isIntlBusinessDayFn = fn === 'NETWORKDAYS.INTL' || fn === 'WORKDAY.INTL';
      const validLength =
        fn === 'TODAY' || fn === 'NOW'
          ? (aggregate[2] ?? '').trim() === ''
          : fn === 'DATE' || fn === 'TIME'
            ? args?.length === 3
            : fn === 'EDATE' ||
                fn === 'EOMONTH' ||
                fn === 'DAYS' ||
                fn === 'NETWORKDAYS' ||
                fn === 'WORKDAY' ||
                fn === 'NETWORKDAYS.INTL' ||
                fn === 'WORKDAY.INTL'
              ? args?.length === 2 ||
                ((fn === 'NETWORKDAYS' || fn === 'WORKDAY') && args?.length === 3) ||
                (isIntlBusinessDayFn && (args?.length === 3 || args?.length === 4))
              : fn === 'DATEDIF'
                ? args?.length === 3
                : fn === 'DAYS360'
                  ? args?.length === 2 || args?.length === 3
                  : fn === 'YEARFRAC'
                    ? args?.length === 2 || args?.length === 3
                    : fn === 'DATEVALUE' || fn === 'TIMEVALUE'
                      ? args?.length === 1
                      : fn === 'WEEKDAY' || fn === 'WEEKNUM'
                        ? args?.length === 1 || args?.length === 2
                        : args?.length === 1;
      if (fn === 'TODAY' || fn === 'NOW') {
        if (validLength) return { kind: 'date-function', fn, args: [] };
      } else if (args && validLength) {
        if (
          (fn === 'WEEKDAY' || fn === 'WEEKNUM' || fn === 'DAYS360' || fn === 'YEARFRAC') &&
          (args[0] ?? '').trim() === ''
        ) {
          return null;
        }
        const operands = args.map((arg, index) => {
          if ((fn === 'WEEKDAY' || fn === 'WEEKNUM') && index === 1 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'number' as const, value: 1 } };
          }
          if (fn === 'DAYS360' && index === 2 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'bool' as const, value: false } };
          }
          if (fn === 'YEARFRAC' && index === 2 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'number' as const, value: 0 } };
          }
          if (isIntlBusinessDayFn && index === 2 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'number' as const, value: 1 } };
          }
          if (
            ((fn === 'NETWORKDAYS' || fn === 'WORKDAY') && index === 2) ||
            (isIntlBusinessDayFn && index === 3)
          ) {
            const range = parseFormulaRangeArg(arg, sheetIndex);
            if (range) return range;
          }
          return parseFormulaOperand(arg, sheetIndex);
        });
        if (operands.every((operand) => operand !== null)) {
          return { kind: 'date-function', fn, args: operands as FormulaDateArg[] };
        }
      }
    }
    if (fn === 'NA') {
      const rawArgs = aggregate[2] ?? '';
      if (rawArgs.trim() === '') {
        return { kind: 'literal', value: { kind: 'error', code: 6, text: '#N/A' } };
      }
    }
    if (fn === 'IFERROR' || fn === 'IFNA') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const fallback = parseFormulaOperand(args[1] ?? '', sheetIndex);
        if (value && fallback) return { kind: 'error-fallback', fn, value, fallback };
      }
    }
    if (fn === 'MATCH') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2 || args?.length === 3) {
        const lookup = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const range = parseFormulaRangeArg(args[1] ?? '', sheetIndex);
        if (lookup && range) {
          if (args.length === 2) return { kind: 'match', lookup, range };
          const matchType = parseFormulaOperand(args[2] ?? '', sheetIndex);
          if (matchType) return { kind: 'match', lookup, range, matchType };
        }
      }
    }
    if (fn === 'OFFSET') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args && args.length >= 3 && args.length <= 5) {
        const reference = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        const rows = parseFormulaOperand(args[1] ?? '', sheetIndex);
        const cols = parseFormulaOperand(args[2] ?? '', sheetIndex);
        if (reference && rows && cols) {
          const height =
            args.length >= 4 ? parseFormulaOperand(args[3] ?? '', sheetIndex) : undefined;
          if (args.length >= 4 && !height) return null;
          const width =
            args.length >= 5 ? parseFormulaOperand(args[4] ?? '', sheetIndex) : undefined;
          if (args.length >= 5 && !width) return null;
          return {
            kind: 'offset',
            reference,
            rows,
            cols,
            ...(height ? { height } : {}),
            ...(width ? { width } : {}),
          };
        }
      }
    }
    if (fn === 'INDIRECT') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 1 || args?.length === 2) {
        const refText = parseFormulaOperand(args[0] ?? '', sheetIndex);
        if (!refText) return null;
        if (args.length === 1) return { kind: 'indirect', refText };
        const a1 = parseFormulaOperand(args[1] ?? '', sheetIndex);
        if (a1) return { kind: 'indirect', refText, a1 };
      }
    }
    if (fn === 'XMATCH') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args && args.length >= 2 && args.length <= 4) {
        const lookup = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const range = parseFormulaRangeArg(args[1] ?? '', sheetIndex);
        if (lookup && range) {
          if (args.length === 2) return { kind: 'xmatch', lookup, range };
          const rawMatchMode = args[2] ?? '';
          const matchMode =
            rawMatchMode.trim() === '' ? undefined : parseFormulaOperand(rawMatchMode, sheetIndex);
          if (rawMatchMode.trim() !== '' && !matchMode) return null;
          if (args.length === 3) {
            return { kind: 'xmatch', lookup, range, ...(matchMode ? { matchMode } : {}) };
          }
          const searchMode = parseFormulaOperand(args[3] ?? '', sheetIndex);
          if (searchMode) {
            return {
              kind: 'xmatch',
              lookup,
              range,
              ...(matchMode ? { matchMode } : {}),
              searchMode,
            };
          }
        }
      }
    }
    if (fn === 'INDEX') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2 || args?.length === 3) {
        const range = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        const row = parseFormulaOperand(args[1] ?? '', sheetIndex);
        if (range && row) {
          if (args.length === 2) return { kind: 'index', range, row };
          const col = parseFormulaOperand(args[2] ?? '', sheetIndex);
          if (col) return { kind: 'index', range, row, col };
        }
      }
    }
    if (fn === 'VLOOKUP' || fn === 'HLOOKUP') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 3 || args?.length === 4) {
        const lookup = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const range = parseFormulaRangeArg(args[1] ?? '', sheetIndex);
        const index = parseFormulaOperand(args[2] ?? '', sheetIndex);
        const rangeLookup =
          args.length === 3
            ? { kind: 'literal' as const, value: { kind: 'bool' as const, value: true } }
            : parseFormulaOperand(args[3] ?? '', sheetIndex);
        if (lookup && range && index && rangeLookup) {
          return { kind: 'lookup', fn, lookup, range, index, rangeLookup };
        }
      }
    }
    if (fn === 'XLOOKUP') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args && args.length >= 3 && args.length <= 6) {
        const lookup = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const lookupRange = parseFormulaRangeArg(args[1] ?? '', sheetIndex);
        const returnRange = parseFormulaRangeArg(args[2] ?? '', sheetIndex);
        if (lookup && lookupRange && returnRange) {
          if (args.length === 3) return { kind: 'xlookup', lookup, lookupRange, returnRange };
          const rawIfNotFound = args[3] ?? '';
          const ifNotFound =
            rawIfNotFound.trim() === ''
              ? undefined
              : parseFormulaOperand(rawIfNotFound, sheetIndex);
          if (rawIfNotFound.trim() !== '' && !ifNotFound) return null;
          if (args.length === 4) {
            return {
              kind: 'xlookup',
              lookup,
              lookupRange,
              returnRange,
              ...(ifNotFound ? { ifNotFound } : {}),
            };
          }
          const rawMatchMode = args[4] ?? '';
          const matchMode =
            rawMatchMode.trim() === '' ? undefined : parseFormulaOperand(rawMatchMode, sheetIndex);
          if (rawMatchMode.trim() !== '' && !matchMode) return null;
          if (args.length === 5) {
            return {
              kind: 'xlookup',
              lookup,
              lookupRange,
              returnRange,
              ...(ifNotFound ? { ifNotFound } : {}),
              ...(matchMode ? { matchMode } : {}),
            };
          }
          const searchMode = parseFormulaOperand(args[5] ?? '', sheetIndex);
          if (searchMode) {
            return {
              kind: 'xlookup',
              lookup,
              lookupRange,
              returnRange,
              ...(ifNotFound ? { ifNotFound } : {}),
              ...(matchMode ? { matchMode } : {}),
              searchMode,
            };
          }
        }
      }
    }
    if (fn === 'LOOKUP') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2 || args?.length === 3) {
        const lookup = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const lookupRange = parseFormulaRangeArg(args[1] ?? '', sheetIndex);
        if (lookup && lookupRange) {
          if (args.length === 2) return { kind: 'vector-lookup', lookup, lookupRange };
          const resultRange = parseFormulaRangeArg(args[2] ?? '', sheetIndex);
          if (resultRange) return { kind: 'vector-lookup', lookup, lookupRange, resultRange };
        }
      }
    }
    if (fn === 'CHOOSE') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args && args.length >= 2) {
        const index = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const choices = args.slice(1).map((arg) => parseFormulaOperand(arg, sheetIndex));
        if (index && choices.every((choice) => choice !== null)) {
          return { kind: 'choose', index, choices: choices as FormulaOperand[] };
        }
      }
    }
    if (fn === 'SWITCH') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args && args.length >= 3) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        if (!value) return null;
        const hasDefault = args.length % 2 === 0;
        const caseArgs = hasDefault ? args.slice(1, -1) : args.slice(1);
        const cases: { match: FormulaOperand; result: FormulaOperand }[] = [];
        for (let i = 0; i < caseArgs.length; i += 2) {
          const match = parseFormulaOperand(caseArgs[i] ?? '', sheetIndex);
          const result = parseFormulaOperand(caseArgs[i + 1] ?? '', sheetIndex);
          if (!match || !result) return null;
          cases.push({ match, result });
        }
        if (hasDefault) {
          const defaultValue = parseFormulaOperand(args[args.length - 1] ?? '', sheetIndex);
          if (defaultValue) return { kind: 'switch', value, cases, defaultValue };
        } else {
          return { kind: 'switch', value, cases };
        }
      }
    }
    if (fn === 'IF') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args && (args.length === 2 || args.length === 3)) {
        const condition = parseFormulaCondition(args[0] ?? '', sheetIndex);
        const whenTrue =
          args[1] === ''
            ? { kind: 'literal' as const, value: { kind: 'number' as const, value: 0 } }
            : parseFormulaOperand(args[1] ?? '', sheetIndex);
        const whenFalse =
          args.length === 2
            ? { kind: 'literal' as const, value: { kind: 'bool' as const, value: false } }
            : args[2] === ''
              ? { kind: 'literal' as const, value: { kind: 'number' as const, value: 0 } }
              : parseFormulaOperand(args[2] ?? '', sheetIndex);
        if (condition && whenTrue && whenFalse)
          return { kind: 'if', condition, whenTrue, whenFalse };
      }
    }
    if (fn === 'IFS') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args && args.length >= 2 && args.length % 2 === 0) {
        const branches: { condition: FormulaCondition; result: FormulaOperand }[] = [];
        for (let i = 0; i < args.length; i += 2) {
          const condition = parseFormulaCondition(args[i] ?? '', sheetIndex);
          const result = parseFormulaOperand(args[i + 1] ?? '', sheetIndex);
          if (!condition || !result) return null;
          branches.push({ condition, result });
        }
        return { kind: 'ifs', branches };
      }
    }
  }
  const arithmetic = splitFormulaArithmetic(body);
  if (arithmetic) {
    const left = parseFormulaOperand(arithmetic.left, sheetIndex);
    const right = parseFormulaOperand(arithmetic.right, sheetIndex);
    if (left && right) return { kind: 'binary', op: arithmetic.op, left, right };
  }
  if (
    (body.startsWith('"') && body.endsWith('"')) ||
    (body.startsWith("'") && body.endsWith("'"))
  ) {
    return { kind: 'literal', value: { kind: 'text', value: body.slice(1, -1) } };
  }
  if (FORMULA_NUMBER_LITERAL.test(body)) {
    return { kind: 'literal', value: { kind: 'number', value: Number(body) } };
  }
  if (/^true$/i.test(body)) return { kind: 'literal', value: { kind: 'bool', value: true } };
  if (/^false$/i.test(body)) return { kind: 'literal', value: { kind: 'bool', value: false } };
  return null;
}

function parseFormulaCondition(raw: string, sheetIndex: number): FormulaCondition | null {
  const body = stripOuterParens(raw.trim());
  if (/^true$/i.test(body)) return { kind: 'bool', value: true };
  if (/^false$/i.test(body)) return { kind: 'bool', value: false };
  const comparison = splitFormulaComparison(body);
  if (comparison) {
    const left = parseFormulaOperand(comparison.left, sheetIndex);
    const right = parseFormulaOperand(comparison.right, sheetIndex);
    if (left && right) return { kind: 'comparison', left, op: comparison.op, right };
  }
  const fnCall = body.match(/^([A-Za-z]+)\s*\((.*)\)$/);
  if (fnCall) {
    const fn = (fnCall[1] ?? '').toUpperCase();
    if ((fn === 'TRUE' || fn === 'FALSE') && (fnCall[2] ?? '').trim() === '') {
      return { kind: 'bool', value: fn === 'TRUE' };
    }
    if (fn === 'AND' || fn === 'OR' || fn === 'NOT' || fn === 'XOR') {
      const args = splitFormulaArgs(fnCall[2] ?? '');
      if (args === null || args.length === 0 || (fn === 'NOT' && args.length !== 1)) return null;
      const conditions = args.map((arg) => parseFormulaCondition(arg, sheetIndex));
      if (conditions.some((condition) => condition === null)) return null;
      return { kind: 'logical', fn, args: conditions as FormulaCondition[] };
    }
    if (
      fn === 'ISBLANK' ||
      fn === 'ISERROR' ||
      fn === 'ISERR' ||
      fn === 'ISNA' ||
      fn === 'ISNUMBER' ||
      fn === 'ISTEXT' ||
      fn === 'ISLOGICAL' ||
      fn === 'ISNONTEXT' ||
      fn === 'ISFORMULA' ||
      fn === 'ISREF'
    ) {
      const args = splitFormulaArgs(fnCall[2] ?? '');
      if (args?.length !== 1) return null;
      if (fn === 'ISREF' || fn === 'ISFORMULA') {
        const range = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        if (range) return { kind: 'is', fn, value: range };
        if (fn === 'ISREF') return null;
      }
      const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
      if (!value) return null;
      return { kind: 'is', fn, value };
    }
  }
  const operand = parseFormulaOperand(body, sheetIndex);
  return operand ? { kind: 'operand', value: operand } : null;
}

export {
  colToLetters,
  FORMULA_NUMBER_LITERAL,
  FORMULA_VALUE_NUMBER_LITERAL,
  MAX_FORMULA_AGGREGATE_CELLS,
  parseA1Range,
  parseA1Ref,
  parseFormulaOperand,
  parseFormulaRangeArg,
  parseR1C1Range,
  parseR1C1Ref,
  splitFormulaArgs,
  splitFormulaArgsAllowEmpty,
  splitFormulaComparison,
  stripOuterParens,
};
