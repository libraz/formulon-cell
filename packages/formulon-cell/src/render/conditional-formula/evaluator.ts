import { addrKey } from '../../engine/address.js';
import type { CellValue } from '../../engine/types.js';
import type { ConditionalRule, State } from '../../store/store.js';
import { aggregateFunction, subtotalFunction } from './aggregation.js';
import { booleanValue, readNumber, textValue } from './coercion.js';
import type { FormulaReaderContext, NumericFunctionName } from './evaluate/context.js';
import { createCriteriaEvaluator } from './evaluate/criteria.js';
import { createDateEvaluator } from './evaluate/dates.js';
import { createFinancialEvaluator } from './evaluate/financial.js';
import { createLookupEvaluator } from './evaluate/lookup.js';
import { createMathEvaluator } from './evaluate/math.js';
import { createProbabilityEvaluator } from './evaluate/probability.js';
import { createRangeReader } from './evaluate/ranges.js';
import { createStatisticsEvaluator } from './evaluate/statistics.js';
import { compareValues } from './matching.js';
import {
  parseFormulaOperand,
  parseFormulaRangeArg,
  splitFormulaArgs,
  splitFormulaArgsAllowEmpty,
  splitFormulaComparison,
  stripOuterParens,
} from './parser.js';
import {
  beforeAfterText,
  coerceScalar,
  concatTextValue,
  exactText,
  fixedFormatText,
  formatText,
  numberValueText,
  repeatText,
  replaceText,
  searchText,
  sliceText,
  substituteText,
  transformText,
  valueText,
  valueToText,
} from './text-functions.js';
import type { FormulaCondition, FormulaOperand } from './types.js';

interface FormulaPredicate {
  /** Evaluate against a cell's value. Returns true to apply the format. */
  test(v: CellValue): boolean;
}

interface FormulaCellPredicate {
  /** Evaluate against the destination cell address within the rule range. */
  test(row: number, col: number): boolean;
}

/** Parse a v1 lightweight predicate: a leading comparison operator
 *  followed by a numeric or quoted-string literal. Anything more complex
 *  returns null and the rule becomes a no-op. */
export function parseFormulaPredicate(raw: string): FormulaPredicate | null {
  const trimmed = raw.trim();
  if (trimmed === '') return null;
  // Strip leading `=` for the comparator-prefix path; an `=`-prefixed
  // expression that doesn't fit a comparator template is reserved for
  // engine-side `evaluateText` (not implemented in v1) — return null.
  let body = trimmed;
  if (body.startsWith('=')) body = body.slice(1).trim();
  // Match: <op><whitespace?><literal>
  const m = body.match(/^(>=|<=|<>|>|<|=)\s*(.+)$/);
  if (!m) return null;
  const op = m[1] as '>' | '<' | '>=' | '<=' | '=' | '<>';
  const rhs = m[2]?.trim() ?? '';
  if (rhs === '') return null;
  // Quoted string literal.
  if ((rhs.startsWith('"') && rhs.endsWith('"')) || (rhs.startsWith("'") && rhs.endsWith("'"))) {
    const inner = rhs.slice(1, -1);
    return {
      test(v): boolean {
        const text = v.kind === 'text' ? v.value : v.kind === 'number' ? String(v.value) : null;
        if (text === null) return false;
        return op === '<>' ? text !== inner : op === '=' ? text === inner : false;
      },
    };
  }
  // Numeric literal.
  const num = Number.parseFloat(rhs);
  if (Number.isNaN(num)) return null;
  return {
    test(v): boolean {
      if (v.kind !== 'number') return false;
      const x = v.value;
      switch (op) {
        case '>':
          return x > num;
        case '<':
          return x < num;
        case '>=':
          return x >= num;
        case '<=':
          return x <= num;
        case '=':
          return x === num;
        case '<>':
          return x !== num;
        default:
          return false;
      }
    },
  };
}

export function compileFormulaCellPredicate(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'formula' }>,
): FormulaCellPredicate | null {
  const body = rule.formula.trim().replace(/^=/, '').trim();
  return parseFormulaBooleanExpression(state, rule, body);
}

function parseFormulaBooleanExpression(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'formula' }>,
  body: string,
): FormulaCellPredicate | null {
  const inner = stripOuterParens(body.trim());
  if (/^true$/i.test(inner)) return { test: () => true };
  if (/^false$/i.test(inner)) return { test: () => false };
  const comparisonPredicate = parseFormulaComparisonPredicate(state, rule, inner);
  if (comparisonPredicate) return comparisonPredicate;
  const logical = inner.match(/^([A-Za-z]+)\s*\((.*)\)$/);
  if (logical) {
    const name = (logical[1] ?? '').toUpperCase();
    if (name === 'AND' || name === 'OR' || name === 'NOT' || name === 'XOR') {
      const args = splitFormulaArgs(logical[2] ?? '');
      if (args === null || args.length === 0) return null;
      const predicates = args.map((arg) => parseFormulaBooleanExpression(state, rule, arg));
      if (predicates.some((predicate) => predicate === null)) return null;
      if (name === 'NOT') {
        if (predicates.length !== 1) return null;
        const predicate = predicates[0] as FormulaCellPredicate;
        return { test: (row, col) => !predicate.test(row, col) };
      }
      const parsed = predicates as FormulaCellPredicate[];
      return {
        test(row, col): boolean {
          if (name === 'AND') return parsed.every((predicate) => predicate.test(row, col));
          if (name === 'OR') return parsed.some((predicate) => predicate.test(row, col));
          return parsed.filter((predicate) => predicate.test(row, col)).length % 2 === 1;
        },
      };
    }
    if (name === 'IF') {
      const args = splitFormulaArgsAllowEmpty(logical[2] ?? '');
      if (!args || (args.length !== 2 && args.length !== 3)) return null;
      const condition = parseFormulaBooleanExpression(state, rule, args[0] ?? '');
      const whenTrue =
        args[1] === ''
          ? { test: () => false }
          : parseFormulaBooleanExpression(state, rule, args[1] ?? '');
      const whenFalse =
        args.length === 2
          ? { test: () => false }
          : args[2] === ''
            ? { test: () => false }
            : parseFormulaBooleanExpression(state, rule, args[2] ?? '');
      if (!condition || !whenTrue || !whenFalse) return null;
      return {
        test(row, col): boolean {
          return condition.test(row, col) ? whenTrue.test(row, col) : whenFalse.test(row, col);
        },
      };
    }
    if (
      name === 'ISBLANK' ||
      name === 'ISERROR' ||
      name === 'ISERR' ||
      name === 'ISNA' ||
      name === 'ISNUMBER' ||
      name === 'ISTEXT' ||
      name === 'ISLOGICAL' ||
      name === 'ISNONTEXT' ||
      name === 'ISFORMULA' ||
      name === 'ISREF'
    ) {
      const args = splitFormulaArgs(logical[2] ?? '');
      if (args?.length !== 1) return null;
      if (name === 'ISREF') {
        const range = parseFormulaRangeArg(args[0] ?? '', state.data.sheetIndex);
        return { test: () => range !== null };
      }
      const formulaRef =
        name === 'ISFORMULA' ? parseFormulaRangeArg(args[0] ?? '', state.data.sheetIndex) : null;
      const operand = parseFormulaOperand(args[0] ?? '', state.data.sheetIndex);
      if (!operand && !formulaRef) return null;
      const readOperand = makeFormulaOperandReader(
        state,
        state.data.sheetIndex,
        rule.range.r0,
        rule.range.c0,
      );
      return {
        test(row, col): boolean {
          const rowOffset = row - rule.range.r0;
          const colOffset = col - rule.range.c0;
          if (name === 'ISFORMULA') {
            if (formulaRef) {
              const value = readOperand(
                { kind: 'formula-text', ref: formulaRef },
                rowOffset,
                colOffset,
              );
              return value.kind === 'text';
            }
            if (!operand) return false;
            if (operand.kind !== 'ref') return false;
            const targetRow = operand.ref.absRow ? operand.ref.row : operand.ref.row + rowOffset;
            const targetCol = operand.ref.absCol ? operand.ref.col : operand.ref.col + colOffset;
            const cell = state.data.cells.get(
              addrKey({ sheet: state.data.sheetIndex, row: targetRow, col: targetCol }),
            );
            return typeof cell?.formula === 'string' && cell.formula.length > 0;
          }
          if (!operand) return false;
          const value = readOperand(operand, rowOffset, colOffset);
          if (name === 'ISBLANK') return value.kind === 'blank';
          if (name === 'ISERROR') return value.kind === 'error';
          if (name === 'ISERR') return value.kind === 'error' && value.text !== '#N/A';
          if (name === 'ISNA') return value.kind === 'error' && value.text === '#N/A';
          if (name === 'ISNUMBER') return value.kind === 'number';
          if (name === 'ISTEXT') return value.kind === 'text';
          if (name === 'ISLOGICAL') return value.kind === 'bool';
          return name === 'ISNONTEXT' && value.kind !== 'text';
        },
      };
    }
  }
  const operand = parseFormulaOperand(inner, state.data.sheetIndex);
  if (operand) {
    const readOperand = makeFormulaOperandReader(
      state,
      state.data.sheetIndex,
      rule.range.r0,
      rule.range.c0,
    );
    return {
      test(row, col): boolean {
        const rowOffset = row - rule.range.r0;
        const colOffset = col - rule.range.c0;
        const value = readOperand(operand, rowOffset, colOffset);
        return value.kind === 'bool' && value.value;
      },
    };
  }
  return null;
}

function parseFormulaComparisonPredicate(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'formula' }>,
  body: string,
): FormulaCellPredicate | null {
  const comparison = splitFormulaComparison(body);
  if (!comparison) return null;
  const sheet = state.data.sheetIndex;
  const left = parseFormulaOperand(comparison.left, sheet);
  const right = parseFormulaOperand(comparison.right, sheet);
  if (!left || !right) return null;
  const readOperand = makeFormulaOperandReader(state, sheet, rule.range.r0, rule.range.c0);
  return {
    test(row, col): boolean {
      const rowOffset = row - rule.range.r0;
      const colOffset = col - rule.range.c0;
      const leftValue = readOperand(left, rowOffset, colOffset);
      const rightValue = readOperand(right, rowOffset, colOffset);
      return compareValues(leftValue, comparison.op, rightValue);
    },
  };
}

function makeFormulaOperandReader(
  state: State,
  sheet: number,
  anchorRow: number,
  anchorCol: number,
): (operand: FormulaOperand, rowOffset: number, colOffset: number) => CellValue {
  const reader: FormulaReaderContext = { state, sheet, anchorRow, anchorCol, readOperand };
  const ranges = createRangeReader(reader);
  const {
    resolveRef,
    aggregateRange,
    aggregateArgs,
    validRangeBounds,
    formulaRangeArgBounds,
    singleCellRefPosition,
  } = ranges;
  const familyContext = { ...reader, ...ranges };
  const {
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
  } = createStatisticsEvaluator(familyContext);
  const {
    countMatchingRange,
    countMatchingRanges,
    sumMatchingRange,
    averageMatchingRange,
    sumMatchingRanges,
    averageMatchingRanges,
    minMaxMatchingRanges,
  } = createCriteriaEvaluator(familyContext);
  const {
    financialFunction,
    npvValue,
    mirrValue,
    xnpvValue,
    irrValue,
    xirrValue,
    fvScheduleValue,
  } = createFinancialEvaluator(familyContext);
  const { probabilityFunction } = createProbabilityEvaluator(familyContext);
  const { mathFunction, numericPredicate } = createMathEvaluator(familyContext);
  const numericFunction = (
    fn: NumericFunctionName,
    args: FormulaOperand[],
    rowOffset: number,
    colOffset: number,
  ): CellValue =>
    financialFunction(fn, args, rowOffset, colOffset) ??
    probabilityFunction(fn, args, rowOffset, colOffset) ??
    mathFunction(fn, args, rowOffset, colOffset);
  const { dateFunction } = createDateEvaluator(familyContext);
  const {
    matchExactRange,
    xmatchRange,
    indexRange,
    offsetValue,
    indirectValue,
    tableLookup,
    xlookupRange,
    vectorLookup,
    cellInfo,
    sheetInfo,
  } = createLookupEvaluator(familyContext);
  const joinText = (
    delimiterValue: CellValue,
    ignoreEmptyValue: CellValue,
    values: FormulaOperand[],
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const delimiter = textValue(delimiterValue);
    const ignoreEmpty = booleanValue(ignoreEmptyValue);
    if (delimiter === null || ignoreEmpty === null)
      return { kind: 'error', code: 15, text: '#VALUE!' };
    const parts: string[] = [];
    for (const value of values) {
      const text = concatTextValue(readOperand(value, rowOffset, colOffset));
      if (ignoreEmpty && text === '') continue;
      parts.push(text);
    }
    const joined = parts.join(delimiter);
    return joined.length > 32767
      ? { kind: 'error', code: 15, text: '#VALUE!' }
      : { kind: 'text', value: joined };
  };
  const chooseValue = (
    indexValue: CellValue,
    choices: FormulaOperand[],
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const indexNumber = readNumber(indexValue);
    if (indexNumber === null) return { kind: 'error', code: 15, text: '#VALUE!' };
    const index = Math.trunc(indexNumber);
    if (index < 1 || index > choices.length) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    return readOperand(choices[index - 1] as FormulaOperand, rowOffset, colOffset);
  };
  const switchValue = (
    value: CellValue,
    cases: { match: FormulaOperand; result: FormulaOperand }[],
    defaultValue: FormulaOperand | undefined,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    for (const item of cases) {
      const match = readOperand(item.match, rowOffset, colOffset);
      if (compareValues(value, '=', match)) return readOperand(item.result, rowOffset, colOffset);
    }
    return defaultValue
      ? readOperand(defaultValue, rowOffset, colOffset)
      : { kind: 'error', code: 6, text: '#N/A' };
  };
  const evalCondition = (
    condition: FormulaCondition,
    rowOffset: number,
    colOffset: number,
  ): boolean => {
    if (condition.kind === 'bool') return condition.value;
    if (condition.kind === 'logical') {
      if (condition.fn === 'NOT')
        return !evalCondition(condition.args[0] as FormulaCondition, rowOffset, colOffset);
      if (condition.fn === 'AND') {
        return condition.args.every((arg) => evalCondition(arg, rowOffset, colOffset));
      }
      if (condition.fn === 'OR') {
        return condition.args.some((arg) => evalCondition(arg, rowOffset, colOffset));
      }
      return (
        condition.args.filter((arg) => evalCondition(arg, rowOffset, colOffset)).length % 2 === 1
      );
    }
    if (condition.kind === 'comparison') {
      return compareValues(
        readOperand(condition.left, rowOffset, colOffset),
        condition.op,
        readOperand(condition.right, rowOffset, colOffset),
      );
    }
    if (condition.kind === 'operand') {
      const value = readOperand(condition.value, rowOffset, colOffset);
      return value.kind === 'bool' && value.value;
    }
    if (condition.fn === 'ISREF') return true;
    if (condition.fn === 'ISFORMULA') {
      const ref =
        condition.value.kind === 'range' || condition.value.kind === 'dynamic-range'
          ? condition.value
          : condition.value.kind === 'ref'
            ? {
                kind: 'range' as const,
                range: { start: condition.value.ref, end: condition.value.ref },
              }
            : null;
      const position = ref ? singleCellRefPosition(ref, rowOffset, colOffset) : null;
      if (!position) return false;
      const cell = state.data.cells.get(addrKey({ sheet, row: position.row, col: position.col }));
      return typeof cell?.formula === 'string' && cell.formula.length > 0;
    }
    if (condition.value.kind === 'range' || condition.value.kind === 'dynamic-range') return false;
    const value = readOperand(condition.value, rowOffset, colOffset);
    if (condition.fn === 'ISBLANK') return value.kind === 'blank';
    if (condition.fn === 'ISERROR') return value.kind === 'error';
    if (condition.fn === 'ISERR') return value.kind === 'error' && value.text !== '#N/A';
    if (condition.fn === 'ISNA') return value.kind === 'error' && value.text === '#N/A';
    if (condition.fn === 'ISNUMBER') return value.kind === 'number';
    if (condition.fn === 'ISTEXT') return value.kind === 'text';
    if (condition.fn === 'ISLOGICAL') return value.kind === 'bool';
    return condition.fn === 'ISNONTEXT' && value.kind !== 'text';
  };
  const ifsValue = (
    branches: { condition: FormulaCondition; result: FormulaOperand }[],
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    for (const branch of branches) {
      if (evalCondition(branch.condition, rowOffset, colOffset)) {
        return readOperand(branch.result, rowOffset, colOffset);
      }
    }
    return { kind: 'error', code: 6, text: '#N/A' };
  };
  function readOperand(operand: FormulaOperand, rowOffset: number, colOffset: number): CellValue {
    if (operand.kind === 'literal') return operand.value;
    if (operand.kind === 'range-aggregate') {
      return aggregateRange(operand.range, operand.fn, rowOffset, colOffset);
    }
    if (operand.kind === 'aggregate-args') {
      return aggregateArgs(operand.args, operand.fn, rowOffset, colOffset);
    }
    if (operand.kind === 'subtotal') {
      const functionNum = readNumber(readOperand(operand.functionNum, rowOffset, colOffset));
      const fn = functionNum === null ? null : subtotalFunction(functionNum);
      return fn === null
        ? { kind: 'error', code: 15, text: '#VALUE!' }
        : aggregateArgs(operand.args, fn, rowOffset, colOffset);
    }
    if (operand.kind === 'aggregate-function') {
      const functionNum = readNumber(readOperand(operand.functionNum, rowOffset, colOffset));
      const options = readNumber(readOperand(operand.options, rowOffset, colOffset));
      const fn = functionNum === null ? null : aggregateFunction(functionNum);
      if (fn === null || options === null || Math.trunc(options) < 0 || Math.trunc(options) > 7) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (fn.kind === 'aggregate') {
        return aggregateArgs(operand.args, fn.fn, rowOffset, colOffset);
      }
      const [rangeArg, valueArg] = operand.args;
      if (
        !rangeArg ||
        (rangeArg.kind !== 'range' && rangeArg.kind !== 'dynamic-range') ||
        !valueArg ||
        valueArg.kind !== 'operand'
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      return fn.kind === 'ranked'
        ? rankedRangeValue(
            fn.fn,
            rangeArg,
            readOperand(valueArg.operand, rowOffset, colOffset),
            rowOffset,
            colOffset,
          )
        : percentileRangeValue(
            fn.fn,
            rangeArg,
            readOperand(valueArg.operand, rowOffset, colOffset),
            null,
            rowOffset,
            colOffset,
          );
    }
    if (operand.kind === 'series-sum') {
      return seriesSumValue(
        readOperand(operand.x, rowOffset, colOffset),
        readOperand(operand.n, rowOffset, colOffset),
        readOperand(operand.m, rowOffset, colOffset),
        operand.coefficients,
        rowOffset,
        colOffset,
      );
    }
    if (operand.kind === 'ranked-range') {
      return rankedRangeValue(
        operand.fn,
        operand.range,
        readOperand(operand.rank, rowOffset, colOffset),
        rowOffset,
        colOffset,
      );
    }
    if (operand.kind === 'percentile-range') {
      return percentileRangeValue(
        operand.fn,
        operand.range,
        readOperand(operand.value, rowOffset, colOffset),
        operand.significance ? readOperand(operand.significance, rowOffset, colOffset) : null,
        rowOffset,
        colOffset,
      );
    }
    if (operand.kind === 'range-rank') {
      return rankRangeValue(
        operand.fn,
        readOperand(operand.value, rowOffset, colOffset),
        operand.range,
        operand.order ? readOperand(operand.order, rowOffset, colOffset) : null,
        rowOffset,
        colOffset,
      );
    }
    if (operand.kind === 'paired-range-stat') {
      return pairedRangeStatValue(operand.fn, operand.left, operand.right, rowOffset, colOffset);
    }
    if (operand.kind === 'regression-forecast') {
      return regressionForecastValue(
        readOperand(operand.x, rowOffset, colOffset),
        operand.knownY,
        operand.knownX,
        rowOffset,
        colOffset,
      );
    }
    if (operand.kind === 'probability-range') {
      return probabilityRangeValue(
        operand.values,
        operand.probabilities,
        readOperand(operand.lower, rowOffset, colOffset),
        operand.upper ? readOperand(operand.upper, rowOffset, colOffset) : null,
        rowOffset,
        colOffset,
      );
    }
    if (operand.kind === 'z-test') {
      return zTestValue(
        operand.range,
        readOperand(operand.x, rowOffset, colOffset),
        operand.sigma ? readOperand(operand.sigma, rowOffset, colOffset) : null,
        rowOffset,
        colOffset,
      );
    }
    if (operand.kind === 't-test') {
      return tTestValue(
        operand.left,
        operand.right,
        readOperand(operand.tails, rowOffset, colOffset),
        readOperand(operand.type, rowOffset, colOffset),
        rowOffset,
        colOffset,
      );
    }
    if (operand.kind === 'chisq-test') {
      return chisqTestValue(operand.actual, operand.expected, rowOffset, colOffset);
    }
    if (operand.kind === 'npv') {
      return npvValue(operand, rowOffset, colOffset);
    }
    if (operand.kind === 'mirr') {
      return mirrValue(operand, rowOffset, colOffset);
    }
    if (operand.kind === 'xnpv') {
      return xnpvValue(operand, rowOffset, colOffset);
    }
    if (operand.kind === 'irr') {
      return irrValue(operand, rowOffset, colOffset);
    }
    if (operand.kind === 'xirr') {
      return xirrValue(operand, rowOffset, colOffset);
    }
    if (operand.kind === 'fv-schedule') {
      return fvScheduleValue(operand, rowOffset, colOffset);
    }
    if (operand.kind === 'sumproduct') {
      return sumProductRanges(operand.ranges, rowOffset, colOffset);
    }
    if (operand.kind === 'countif') {
      const criteria = readOperand(operand.criteria, rowOffset, colOffset);
      return countMatchingRange(operand.range, criteria, rowOffset, colOffset);
    }
    if (operand.kind === 'countifs') {
      const pairs = operand.pairs.map((pair) => ({
        range: pair.range,
        criteria: readOperand(pair.criteria, rowOffset, colOffset),
      }));
      return countMatchingRanges(pairs, rowOffset, colOffset);
    }
    if (operand.kind === 'sumif') {
      const criteria = readOperand(operand.criteria, rowOffset, colOffset);
      return sumMatchingRange(operand.range, criteria, operand.sumRange, rowOffset, colOffset);
    }
    if (operand.kind === 'averageif') {
      const criteria = readOperand(operand.criteria, rowOffset, colOffset);
      return averageMatchingRange(
        operand.range,
        criteria,
        operand.averageRange,
        rowOffset,
        colOffset,
      );
    }
    if (operand.kind === 'sumifs') {
      const pairs = operand.pairs.map((pair) => ({
        range: pair.range,
        criteria: readOperand(pair.criteria, rowOffset, colOffset),
      }));
      return sumMatchingRanges(operand.sumRange, pairs, rowOffset, colOffset);
    }
    if (operand.kind === 'averageifs') {
      const pairs = operand.pairs.map((pair) => ({
        range: pair.range,
        criteria: readOperand(pair.criteria, rowOffset, colOffset),
      }));
      return averageMatchingRanges(operand.averageRange, pairs, rowOffset, colOffset);
    }
    if (operand.kind === 'minmaxifs') {
      const pairs = operand.pairs.map((pair) => ({
        range: pair.range,
        criteria: readOperand(pair.criteria, rowOffset, colOffset),
      }));
      return minMaxMatchingRanges(operand.valueRange, operand.fn, pairs, rowOffset, colOffset);
    }
    if (operand.kind === 'text-length') {
      const value = textValue(readOperand(operand.value, rowOffset, colOffset));
      return value === null
        ? { kind: 'error', code: 15, text: '#VALUE!' }
        : { kind: 'number', value: value.length };
    }
    if (operand.kind === 'formula-text') {
      const position = singleCellRefPosition(operand.ref, rowOffset, colOffset);
      if (!position) return { kind: 'error', code: 15, text: '#VALUE!' };
      const cell = state.data.cells.get(addrKey({ sheet, row: position.row, col: position.col }));
      return typeof cell?.formula === 'string' && cell.formula.length > 0
        ? { kind: 'text', value: cell.formula }
        : { kind: 'error', code: 6, text: '#N/A' };
    }
    if (operand.kind === 'hyperlink') {
      if (operand.friendlyName) return readOperand(operand.friendlyName, rowOffset, colOffset);
      const link = textValue(readOperand(operand.link, rowOffset, colOffset));
      return link === null
        ? { kind: 'error', code: 15, text: '#VALUE!' }
        : { kind: 'text', value: link };
    }
    if (operand.kind === 'text-search') {
      return searchText(
        operand.fn,
        readOperand(operand.needle, rowOffset, colOffset),
        readOperand(operand.haystack, rowOffset, colOffset),
        operand.start ? readOperand(operand.start, rowOffset, colOffset) : null,
      );
    }
    if (operand.kind === 'text-slice') {
      return sliceText(
        operand.fn,
        readOperand(operand.value, rowOffset, colOffset),
        operand.start ? readOperand(operand.start, rowOffset, colOffset) : null,
        readOperand(operand.count, rowOffset, colOffset),
      );
    }
    if (operand.kind === 'text-transform') {
      return transformText(operand.fn, readOperand(operand.value, rowOffset, colOffset));
    }
    if (operand.kind === 'text-substitute') {
      return substituteText(
        readOperand(operand.value, rowOffset, colOffset),
        readOperand(operand.oldText, rowOffset, colOffset),
        readOperand(operand.newText, rowOffset, colOffset),
        operand.instance ? readOperand(operand.instance, rowOffset, colOffset) : null,
      );
    }
    if (operand.kind === 'text-replace') {
      return replaceText(
        readOperand(operand.value, rowOffset, colOffset),
        readOperand(operand.start, rowOffset, colOffset),
        readOperand(operand.count, rowOffset, colOffset),
        readOperand(operand.newText, rowOffset, colOffset),
      );
    }
    if (operand.kind === 'text-repeat') {
      return repeatText(
        readOperand(operand.value, rowOffset, colOffset),
        readOperand(operand.count, rowOffset, colOffset),
      );
    }
    if (operand.kind === 'text-before-after') {
      return beforeAfterText(
        operand.fn,
        readOperand(operand.value, rowOffset, colOffset),
        readOperand(operand.delimiter, rowOffset, colOffset),
        operand.instance ? readOperand(operand.instance, rowOffset, colOffset) : null,
        operand.matchMode ? readOperand(operand.matchMode, rowOffset, colOffset) : null,
        operand.matchEnd ? readOperand(operand.matchEnd, rowOffset, colOffset) : null,
        operand.ifNotFound ? readOperand(operand.ifNotFound, rowOffset, colOffset) : null,
      );
    }
    if (operand.kind === 'text-join') {
      return joinText(
        readOperand(operand.delimiter, rowOffset, colOffset),
        readOperand(operand.ignoreEmpty, rowOffset, colOffset),
        operand.values,
        rowOffset,
        colOffset,
      );
    }
    if (operand.kind === 'text-exact') {
      return exactText(
        readOperand(operand.left, rowOffset, colOffset),
        readOperand(operand.right, rowOffset, colOffset),
      );
    }
    if (operand.kind === 'text-format') {
      return formatText(
        readOperand(operand.value, rowOffset, colOffset),
        readOperand(operand.pattern, rowOffset, colOffset),
      );
    }
    if (operand.kind === 'text-fixed-format') {
      return fixedFormatText(
        operand.fn,
        readOperand(operand.value, rowOffset, colOffset),
        operand.decimals ? readOperand(operand.decimals, rowOffset, colOffset) : null,
        operand.noCommas ? readOperand(operand.noCommas, rowOffset, colOffset) : null,
      );
    }
    if (operand.kind === 'text-value') {
      return valueText(readOperand(operand.value, rowOffset, colOffset));
    }
    if (operand.kind === 'text-number-value') {
      return numberValueText(
        readOperand(operand.value, rowOffset, colOffset),
        operand.decimalSeparator
          ? readOperand(operand.decimalSeparator, rowOffset, colOffset)
          : null,
        operand.groupSeparator ? readOperand(operand.groupSeparator, rowOffset, colOffset) : null,
      );
    }
    if (operand.kind === 'value-to-text') {
      return valueToText(
        readOperand(operand.value, rowOffset, colOffset),
        operand.format ? readOperand(operand.format, rowOffset, colOffset) : null,
      );
    }
    if (operand.kind === 'scalar-coerce') {
      return coerceScalar(operand.fn, readOperand(operand.value, rowOffset, colOffset));
    }
    if (operand.kind === 'condition-value') {
      return { kind: 'bool', value: evalCondition(operand.condition, rowOffset, colOffset) };
    }
    if (operand.kind === 'range-dimension') {
      const bounds = formulaRangeArgBounds(operand.range, rowOffset, colOffset);
      if (!bounds || !validRangeBounds(bounds)) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const rows = bounds.height;
      const cols = bounds.width;
      if (operand.fn === 'AREAS') return { kind: 'number', value: 1 };
      return { kind: 'number', value: operand.fn === 'ROWS' ? rows : cols };
    }
    if (operand.kind === 'position') {
      if (operand.ref) {
        const position = singleCellRefPosition(operand.ref, rowOffset, colOffset);
        if (!position) return { kind: 'error', code: 15, text: '#VALUE!' };
        return {
          kind: 'number',
          value: operand.fn === 'ROW' ? position.row + 1 : position.col + 1,
        };
      }
      return {
        kind: 'number',
        value: operand.fn === 'ROW' ? anchorRow + rowOffset + 1 : anchorCol + colOffset + 1,
      };
    }
    if (operand.kind === 'numeric-function') {
      return numericFunction(operand.fn, operand.args, rowOffset, colOffset);
    }
    if (operand.kind === 'numeric-predicate') {
      return numericPredicate(operand.fn, operand.value, rowOffset, colOffset);
    }
    if (operand.kind === 'date-function') {
      return dateFunction(operand.fn, operand.args, rowOffset, colOffset);
    }
    if (operand.kind === 'error-fallback') {
      const value = readOperand(operand.value, rowOffset, colOffset);
      if (value.kind !== 'error') return value;
      if (operand.fn === 'IFNA' && value.text !== '#N/A') return value;
      return readOperand(operand.fallback, rowOffset, colOffset);
    }
    if (operand.kind === 'match') {
      return matchExactRange(
        readOperand(operand.lookup, rowOffset, colOffset),
        operand.range,
        operand.matchType
          ? readOperand(operand.matchType, rowOffset, colOffset)
          : { kind: 'number', value: 1 },
        rowOffset,
        colOffset,
      );
    }
    if (operand.kind === 'offset') {
      return offsetValue(
        operand.reference,
        readOperand(operand.rows, rowOffset, colOffset),
        readOperand(operand.cols, rowOffset, colOffset),
        operand.height ? readOperand(operand.height, rowOffset, colOffset) : null,
        operand.width ? readOperand(operand.width, rowOffset, colOffset) : null,
        rowOffset,
        colOffset,
      );
    }
    if (operand.kind === 'indirect') {
      return indirectValue(
        readOperand(operand.refText, rowOffset, colOffset),
        operand.a1 ? readOperand(operand.a1, rowOffset, colOffset) : null,
        rowOffset,
        colOffset,
      );
    }
    if (operand.kind === 'xmatch') {
      return xmatchRange(
        readOperand(operand.lookup, rowOffset, colOffset),
        operand.range,
        operand.matchMode ? readOperand(operand.matchMode, rowOffset, colOffset) : null,
        operand.searchMode ? readOperand(operand.searchMode, rowOffset, colOffset) : null,
        rowOffset,
        colOffset,
      );
    }
    if (operand.kind === 'index') {
      return indexRange(
        operand.range,
        readOperand(operand.row, rowOffset, colOffset),
        operand.col ? readOperand(operand.col, rowOffset, colOffset) : null,
        rowOffset,
        colOffset,
      );
    }
    if (operand.kind === 'lookup') {
      return tableLookup(
        operand.fn,
        readOperand(operand.lookup, rowOffset, colOffset),
        operand.range,
        readOperand(operand.index, rowOffset, colOffset),
        readOperand(operand.rangeLookup, rowOffset, colOffset),
        rowOffset,
        colOffset,
      );
    }
    if (operand.kind === 'xlookup') {
      return xlookupRange(
        readOperand(operand.lookup, rowOffset, colOffset),
        operand.lookupRange,
        operand.returnRange,
        operand.ifNotFound ? readOperand(operand.ifNotFound, rowOffset, colOffset) : null,
        operand.matchMode ? readOperand(operand.matchMode, rowOffset, colOffset) : null,
        operand.searchMode ? readOperand(operand.searchMode, rowOffset, colOffset) : null,
        rowOffset,
        colOffset,
      );
    }
    if (operand.kind === 'vector-lookup') {
      return vectorLookup(
        readOperand(operand.lookup, rowOffset, colOffset),
        operand.lookupRange,
        operand.resultRange,
        rowOffset,
        colOffset,
      );
    }
    if (operand.kind === 'cell-info') {
      return cellInfo(
        readOperand(operand.infoType, rowOffset, colOffset),
        operand.ref,
        rowOffset,
        colOffset,
      );
    }
    if (operand.kind === 'sheet-info') {
      return sheetInfo(operand.fn, operand.range, rowOffset, colOffset);
    }
    if (operand.kind === 'choose') {
      return chooseValue(
        readOperand(operand.index, rowOffset, colOffset),
        operand.choices,
        rowOffset,
        colOffset,
      );
    }
    if (operand.kind === 'switch') {
      return switchValue(
        readOperand(operand.value, rowOffset, colOffset),
        operand.cases,
        operand.defaultValue,
        rowOffset,
        colOffset,
      );
    }
    if (operand.kind === 'if') {
      return readOperand(
        evalCondition(operand.condition, rowOffset, colOffset)
          ? operand.whenTrue
          : operand.whenFalse,
        rowOffset,
        colOffset,
      );
    }
    if (operand.kind === 'ifs') {
      return ifsValue(operand.branches, rowOffset, colOffset);
    }
    if (operand.kind === 'text-concat-function') {
      return {
        kind: 'text',
        value: operand.values
          .map((value) => concatTextValue(readOperand(value, rowOffset, colOffset)))
          .join(''),
      };
    }
    if (operand.kind === 'binary') {
      const leftValue = readOperand(operand.left, rowOffset, colOffset);
      const rightValue = readOperand(operand.right, rowOffset, colOffset);
      if (operand.op === '&') {
        return { kind: 'text', value: concatTextValue(leftValue) + concatTextValue(rightValue) };
      }
      if (leftValue.kind !== 'number' || rightValue.kind !== 'number') {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      let value: number;
      switch (operand.op) {
        case '+':
          value = leftValue.value + rightValue.value;
          break;
        case '-':
          value = leftValue.value - rightValue.value;
          break;
        case '*':
          value = leftValue.value * rightValue.value;
          break;
        case '/':
          value = rightValue.value === 0 ? Number.NaN : leftValue.value / rightValue.value;
          break;
        case '^':
          value = leftValue.value ** rightValue.value;
          break;
      }
      return Number.isFinite(value)
        ? { kind: 'number', value }
        : { kind: 'error', code: 1, text: '#DIV/0!' };
    }
    const [row, col] = resolveRef(operand.ref, rowOffset, colOffset);
    return state.data.cells.get(addrKey({ sheet, row, col }))?.value ?? { kind: 'blank' };
  }
  return readOperand;
}
