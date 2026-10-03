import { addrKey } from '../../engine/address.js';
import type { CellValue } from '../../engine/types.js';
import type { ConditionalRule, State } from '../../store/store.js';
import {
  aggregateFunction,
  aggregateResult,
  aggregateValueA,
  percentileExcValue,
  percentileIncValue,
  subtotalFunction,
} from './aggregation.js';
import {
  datedif,
  dateFromSerial,
  dateSerialFromParts,
  dateValueText,
  dayOfYear,
  days360,
  defaultWeekendDays,
  isoWeekNumber,
  networkDays,
  serialTimeFraction,
  timeValueText,
  todaySerial,
  weekdayValue,
  weekendDaysFromValue,
  weekStartForReturnType,
  workday,
  yearFrac,
} from './calendar.js';
import { booleanValue, positiveInteger, readLogical, readNumber, textValue } from './coercion.js';
import {
  binomialProbability,
  combination,
  doubleFactorial,
  erf,
  factorial,
  gamma,
  hypergeometricProbability,
  inverseRegularizedBeta,
  inverseRegularizedGammaP,
  inverseStandardNormal,
  inverseStudentTCdf,
  logGamma,
  negativeBinomialProbability,
  poissonProbability,
  regularizedBeta,
  regularizedGammaP,
  standardNormalCdf,
  standardNormalPdf,
  studentTCdf,
  studentTPdf,
} from './distributions.js';
import {
  approximateMatchIndex,
  approximateXmatchIndex,
  compareValues,
  exactMatchValues,
  isApproximateLookupMode,
  isExactLookupMode,
  matchesCountIfCriteria,
} from './matching.js';
import {
  baseDigits,
  bitOperand,
  engineeringBaseText,
  engineeringBaseValue,
  maxBitValue,
  romanText,
  romanValue,
} from './numerals.js';
import {
  colToLetters,
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
import type {
  FormulaAggregateArg,
  FormulaAggregateName,
  FormulaCondition,
  FormulaDateArg,
  FormulaOperand,
  FormulaRangeArg,
  FormulaRangeOperand,
  ParsedA1Range,
  ParsedRef,
} from './types.js';

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
  const resolveRef = (ref: ParsedRef, rowOffset: number, colOffset: number): [number, number] => [
    ref.absRow ? ref.row : ref.row + rowOffset,
    ref.absCol ? ref.col : ref.col + colOffset,
  ];
  const rangeAggregateStats = (
    range: FormulaRangeArg,
    rowOffset: number,
    colOffset: number,
  ): { values: number[]; valuesA: number[]; countA: number; countBlank: number } | null => {
    const bounds = formulaRangeArgBounds(range, rowOffset, colOffset);
    if (!bounds) return null;
    const { r0, r1, c0, c1 } = bounds;
    if (!validRangeBounds(bounds)) return null;
    const values: number[] = [];
    const valuesA: number[] = [];
    let countA = 0;
    let countBlank = 0;
    for (let r = r0; r <= r1; r += 1) {
      for (let c = c0; c <= c1; c += 1) {
        const value = state.data.cells.get(addrKey({ sheet, row: r, col: c }))?.value ?? {
          kind: 'blank' as const,
        };
        if (value.kind === 'blank') countBlank += 1;
        else countA += 1;
        if (value?.kind === 'number' && Number.isFinite(value.value)) values.push(value.value);
        const valueA = aggregateValueA(value);
        if (valueA !== null) valuesA.push(valueA);
      }
    }
    return { values, valuesA, countA, countBlank };
  };
  const aggregateRange = (
    range: FormulaRangeArg,
    fn: FormulaAggregateName,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const stats = rangeAggregateStats(range, rowOffset, colOffset);
    if (!stats) return { kind: 'error', code: 15, text: '#VALUE!' };
    return aggregateResult(fn, stats.values, stats.valuesA, stats.countA, stats.countBlank);
  };
  const aggregateArgs = (
    args: FormulaAggregateArg[],
    fn: FormulaAggregateName,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const values: number[] = [];
    const valuesA: number[] = [];
    let countA = 0;
    let countBlank = 0;
    for (const arg of args) {
      if (arg.kind === 'range' || arg.kind === 'dynamic-range') {
        const bounds = formulaRangeArgBounds(arg, rowOffset, colOffset);
        const stats = bounds ? rangeAggregateStatsFromBounds(bounds) : null;
        if (!stats) return { kind: 'error', code: 15, text: '#VALUE!' };
        values.push(...stats.values);
        valuesA.push(...stats.valuesA);
        countA += stats.countA;
        countBlank += stats.countBlank;
        continue;
      }
      const value = readOperand(arg.operand, rowOffset, colOffset);
      if (value.kind === 'blank') countBlank += 1;
      else countA += 1;
      if (value.kind === 'number' && Number.isFinite(value.value)) values.push(value.value);
      const valueA = aggregateValueA(value);
      if (valueA !== null) valuesA.push(valueA);
    }
    return aggregateResult(fn, values, valuesA, countA, countBlank);
  };
  const rangeBounds = (
    range: ParsedA1Range,
    rowOffset: number,
    colOffset: number,
  ): { r0: number; r1: number; c0: number; c1: number; width: number; height: number } => {
    const [startRow, startCol] = resolveRef(range.start, rowOffset, colOffset);
    const [endRow, endCol] = resolveRef(range.end, rowOffset, colOffset);
    const r0 = Math.min(startRow, endRow);
    const r1 = Math.max(startRow, endRow);
    const c0 = Math.min(startCol, endCol);
    const c1 = Math.max(startCol, endCol);
    const width = c1 - c0 + 1;
    const height = r1 - r0 + 1;
    return { r0, r1, c0, c1, width, height };
  };
  const fixedRangeBounds = (
    range: ParsedA1Range,
  ): { r0: number; r1: number; c0: number; c1: number; width: number; height: number } => {
    const r0 = Math.min(range.start.row, range.end.row);
    const r1 = Math.max(range.start.row, range.end.row);
    const c0 = Math.min(range.start.col, range.end.col);
    const c1 = Math.max(range.start.col, range.end.col);
    return { r0, r1, c0, c1, width: c1 - c0 + 1, height: r1 - r0 + 1 };
  };
  const validRangeBounds = (bounds: { width: number; height: number }): boolean =>
    bounds.width > 0 &&
    bounds.height > 0 &&
    bounds.width * bounds.height <= MAX_FORMULA_AGGREGATE_CELLS;
  const dynamicRangeBounds = (
    range: FormulaRangeOperand,
    rowOffset: number,
    colOffset: number,
  ): { r0: number; r1: number; c0: number; c1: number; width: number; height: number } | null => {
    if (range.kind === 'offset-range') {
      const rows = readNumber(readOperand(range.rows, rowOffset, colOffset));
      const cols = readNumber(readOperand(range.cols, rowOffset, colOffset));
      const rawHeight = range.height
        ? readNumber(readOperand(range.height, rowOffset, colOffset))
        : null;
      const rawWidth = range.width
        ? readNumber(readOperand(range.width, rowOffset, colOffset))
        : null;
      if (
        rows === null ||
        cols === null ||
        (range.height && rawHeight === null) ||
        (range.width && rawWidth === null)
      ) {
        return null;
      }
      const base = formulaRangeArgBounds(range.reference, rowOffset, colOffset);
      if (!base || !validRangeBounds(base)) return null;
      const height = rawHeight === null ? base.height : Math.trunc(rawHeight);
      const width = rawWidth === null ? base.width : Math.trunc(rawWidth);
      const r0 = base.r0 + Math.trunc(rows);
      const c0 = base.c0 + Math.trunc(cols);
      const bounds = { r0, r1: r0 + height - 1, c0, c1: c0 + width - 1, width, height };
      return bounds.r0 < 0 ||
        bounds.c0 < 0 ||
        bounds.r1 > 1048575 ||
        bounds.c1 > 16383 ||
        !validRangeBounds(bounds)
        ? null
        : bounds;
    }
    const refText = textValue(readOperand(range.refText, rowOffset, colOffset));
    const a1 = range.a1 ? readLogical(readOperand(range.a1, rowOffset, colOffset)) : true;
    if (refText === null || a1 === null) return null;
    const parsed = a1
      ? parseA1Range(refText, sheet)
      : parseR1C1Range(refText, sheet, anchorRow + rowOffset, anchorCol + colOffset);
    if (!parsed) return null;
    const bounds = fixedRangeBounds(parsed);
    return bounds.r0 < 0 ||
      bounds.c0 < 0 ||
      bounds.r1 > 1048575 ||
      bounds.c1 > 16383 ||
      !validRangeBounds(bounds)
      ? null
      : bounds;
  };
  const formulaRangeArgBounds = (
    arg: FormulaRangeArg | ParsedA1Range,
    rowOffset: number,
    colOffset: number,
  ): { r0: number; r1: number; c0: number; c1: number; width: number; height: number } | null =>
    !('kind' in arg)
      ? rangeBounds(arg, rowOffset, colOffset)
      : arg.kind === 'range'
        ? rangeBounds(arg.range, rowOffset, colOffset)
        : dynamicRangeBounds(arg.range, rowOffset, colOffset);
  const singleCellRefPosition = (
    ref: FormulaRangeArg,
    rowOffset: number,
    colOffset: number,
  ): { row: number; col: number } | null => {
    const bounds = formulaRangeArgBounds(ref, rowOffset, colOffset);
    if (bounds?.width !== 1 || bounds.height !== 1) return null;
    return { row: bounds.r0, col: bounds.c0 };
  };
  const rangeAggregateStatsFromBounds = (bounds: {
    r0: number;
    r1: number;
    c0: number;
    c1: number;
    width: number;
    height: number;
  }): { values: number[]; valuesA: number[]; countA: number; countBlank: number } | null => {
    const { r0, r1, c0, c1 } = bounds;
    if (!validRangeBounds(bounds)) return null;
    const values: number[] = [];
    const valuesA: number[] = [];
    let countA = 0;
    let countBlank = 0;
    for (let r = r0; r <= r1; r += 1) {
      for (let c = c0; c <= c1; c += 1) {
        const value = state.data.cells.get(addrKey({ sheet, row: r, col: c }))?.value ?? {
          kind: 'blank' as const,
        };
        if (value.kind === 'blank') countBlank += 1;
        else countA += 1;
        if (value?.kind === 'number' && Number.isFinite(value.value)) values.push(value.value);
        const valueA = aggregateValueA(value);
        if (valueA !== null) valuesA.push(valueA);
      }
    }
    return { values, valuesA, countA, countBlank };
  };
  const numericValuesInBounds = (bounds: {
    r0: number;
    r1: number;
    c0: number;
    c1: number;
    width: number;
    height: number;
  }): number[] | null => {
    const { r0, r1, c0, c1 } = bounds;
    if (!validRangeBounds(bounds)) return null;
    const values: number[] = [];
    for (let r = r0; r <= r1; r += 1) {
      for (let c = c0; c <= c1; c += 1) {
        const value = state.data.cells.get(addrKey({ sheet, row: r, col: c }))?.value;
        if (value?.kind === 'number' && Number.isFinite(value.value)) values.push(value.value);
      }
    }
    return values;
  };
  const numericValuesFromArgs = (
    args: FormulaAggregateArg[],
    rowOffset: number,
    colOffset: number,
  ): number[] | null => {
    const values: number[] = [];
    for (const arg of args) {
      if (arg.kind === 'range' || arg.kind === 'dynamic-range') {
        const bounds = formulaRangeArgBounds(arg, rowOffset, colOffset);
        const rangeValues = bounds ? numericValuesInBounds(bounds) : null;
        if (rangeValues === null) return null;
        values.push(...rangeValues);
        continue;
      }
      const value = readNumber(readOperand(arg.operand, rowOffset, colOffset));
      if (value === null) return null;
      values.push(value);
    }
    return values;
  };
  const numericValuesInFormulaRangeArg = (
    range: FormulaRangeArg | ParsedA1Range,
    rowOffset: number,
    colOffset: number,
  ): number[] | null => {
    const bounds = formulaRangeArgBounds(range, rowOffset, colOffset);
    return bounds ? numericValuesInBounds(bounds) : null;
  };
  const numericValuesInRangeWithShape = (
    range: FormulaRangeArg | ParsedA1Range,
    rowOffset: number,
    colOffset: number,
  ): { values: number[]; width: number; height: number } | null => {
    const bounds = formulaRangeArgBounds(range, rowOffset, colOffset);
    if (!bounds) return null;
    const { r0, r1, c0, c1, width, height } = bounds;
    if (!validRangeBounds(bounds)) return null;
    const values: number[] = [];
    for (let r = r0; r <= r1; r += 1) {
      for (let c = c0; c <= c1; c += 1) {
        const value = state.data.cells.get(addrKey({ sheet, row: r, col: c }))?.value;
        if (value?.kind !== 'number' || !Number.isFinite(value.value)) return null;
        values.push(value.value);
      }
    }
    return { values, width, height };
  };
  const numericPairsInRanges = (
    left: FormulaRangeArg | ParsedA1Range,
    right: FormulaRangeArg | ParsedA1Range,
    rowOffset: number,
    colOffset: number,
  ): { xs: number[]; ys: number[] } | null => {
    const leftBounds = formulaRangeArgBounds(left, rowOffset, colOffset);
    const rightBounds = formulaRangeArgBounds(right, rowOffset, colOffset);
    if (
      !leftBounds ||
      !rightBounds ||
      !validRangeBounds(leftBounds) ||
      !validRangeBounds(rightBounds) ||
      leftBounds.width !== rightBounds.width ||
      leftBounds.height !== rightBounds.height
    ) {
      return null;
    }
    const xs: number[] = [];
    const ys: number[] = [];
    for (let r = 0; r < leftBounds.height; r += 1) {
      for (let c = 0; c < leftBounds.width; c += 1) {
        const leftValue = state.data.cells.get(
          addrKey({ sheet, row: leftBounds.r0 + r, col: leftBounds.c0 + c }),
        )?.value;
        const rightValue = state.data.cells.get(
          addrKey({ sheet, row: rightBounds.r0 + r, col: rightBounds.c0 + c }),
        )?.value;
        if (
          leftValue?.kind === 'number' &&
          Number.isFinite(leftValue.value) &&
          rightValue?.kind === 'number' &&
          Number.isFinite(rightValue.value)
        ) {
          xs.push(leftValue.value);
          ys.push(rightValue.value);
        }
      }
    }
    return { xs, ys };
  };
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
  const countMatchingRange = (
    range: FormulaRangeArg,
    criteria: CellValue,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const bounds = formulaRangeArgBounds(range, rowOffset, colOffset);
    if (!bounds || !validRangeBounds(bounds)) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    let count = 0;
    for (let r = bounds.r0; r <= bounds.r1; r += 1) {
      for (let c = bounds.c0; c <= bounds.c1; c += 1) {
        const value = state.data.cells.get(addrKey({ sheet, row: r, col: c }))?.value ?? {
          kind: 'blank' as const,
        };
        if (matchesCountIfCriteria(value, criteria)) count += 1;
      }
    }
    return { kind: 'number', value: count };
  };
  const countMatchingRanges = (
    pairs: { range: FormulaRangeArg; criteria: CellValue }[],
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const nullableBounds = pairs.map((pair) =>
      formulaRangeArgBounds(pair.range, rowOffset, colOffset),
    );
    if (!nullableBounds.every((bound) => bound !== null && validRangeBounds(bound))) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const bounds = nullableBounds as NonNullable<(typeof nullableBounds)[number]>[];
    const first = bounds[0];
    if (!first) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    if (bounds.some((bound) => bound.width !== first.width || bound.height !== first.height)) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    let count = 0;
    for (let dr = 0; dr < first.height; dr += 1) {
      for (let dc = 0; dc < first.width; dc += 1) {
        const matches = pairs.every((pair, index) => {
          const bound = bounds[index] as typeof first;
          const value = state.data.cells.get(
            addrKey({ sheet, row: bound.r0 + dr, col: bound.c0 + dc }),
          )?.value ?? { kind: 'blank' as const };
          return matchesCountIfCriteria(value, pair.criteria);
        });
        if (matches) count += 1;
      }
    }
    return { kind: 'number', value: count };
  };
  const sumMatchingRange = (
    range: FormulaRangeArg,
    criteria: CellValue,
    sumRange: FormulaRangeArg,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const criteriaBounds = formulaRangeArgBounds(range, rowOffset, colOffset);
    const sumBounds = formulaRangeArgBounds(sumRange, rowOffset, colOffset);
    if (
      !criteriaBounds ||
      !sumBounds ||
      !validRangeBounds(criteriaBounds) ||
      !validRangeBounds(sumBounds)
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    if (sumBounds.width !== criteriaBounds.width || sumBounds.height !== criteriaBounds.height) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    let sum = 0;
    for (let dr = 0; dr < criteriaBounds.height; dr += 1) {
      for (let dc = 0; dc < criteriaBounds.width; dc += 1) {
        const criteriaValue = state.data.cells.get(
          addrKey({ sheet, row: criteriaBounds.r0 + dr, col: criteriaBounds.c0 + dc }),
        )?.value ?? { kind: 'blank' as const };
        if (!matchesCountIfCriteria(criteriaValue, criteria)) continue;
        const sumValue = state.data.cells.get(
          addrKey({ sheet, row: sumBounds.r0 + dr, col: sumBounds.c0 + dc }),
        )?.value;
        if (sumValue?.kind === 'number' && Number.isFinite(sumValue.value)) sum += sumValue.value;
      }
    }
    return { kind: 'number', value: sum };
  };
  const averageMatchingRange = (
    range: FormulaRangeArg,
    criteria: CellValue,
    averageRange: FormulaRangeArg,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const criteriaBounds = formulaRangeArgBounds(range, rowOffset, colOffset);
    const averageBounds = formulaRangeArgBounds(averageRange, rowOffset, colOffset);
    if (
      !criteriaBounds ||
      !averageBounds ||
      !validRangeBounds(criteriaBounds) ||
      !validRangeBounds(averageBounds)
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    if (
      averageBounds.width !== criteriaBounds.width ||
      averageBounds.height !== criteriaBounds.height
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    let sum = 0;
    let count = 0;
    for (let dr = 0; dr < criteriaBounds.height; dr += 1) {
      for (let dc = 0; dc < criteriaBounds.width; dc += 1) {
        const criteriaValue = state.data.cells.get(
          addrKey({ sheet, row: criteriaBounds.r0 + dr, col: criteriaBounds.c0 + dc }),
        )?.value ?? { kind: 'blank' as const };
        if (!matchesCountIfCriteria(criteriaValue, criteria)) continue;
        const averageValue = state.data.cells.get(
          addrKey({ sheet, row: averageBounds.r0 + dr, col: averageBounds.c0 + dc }),
        )?.value;
        if (averageValue?.kind === 'number' && Number.isFinite(averageValue.value)) {
          sum += averageValue.value;
          count += 1;
        }
      }
    }
    return count > 0
      ? { kind: 'number', value: sum / count }
      : { kind: 'error', code: 1, text: '#DIV/0!' };
  };
  const sumMatchingRanges = (
    sumRange: FormulaRangeArg,
    pairs: { range: FormulaRangeArg; criteria: CellValue }[],
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const sumBounds = formulaRangeArgBounds(sumRange, rowOffset, colOffset);
    const nullableBounds = pairs.map((pair) =>
      formulaRangeArgBounds(pair.range, rowOffset, colOffset),
    );
    if (
      !sumBounds ||
      !validRangeBounds(sumBounds) ||
      !nullableBounds.every((bound) => bound !== null && validRangeBounds(bound))
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const bounds = nullableBounds as NonNullable<(typeof nullableBounds)[number]>[];
    if (
      bounds.some((bound) => bound.width !== sumBounds.width || bound.height !== sumBounds.height)
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    let sum = 0;
    for (let dr = 0; dr < sumBounds.height; dr += 1) {
      for (let dc = 0; dc < sumBounds.width; dc += 1) {
        const matches = pairs.every((pair, index) => {
          const bound = bounds[index] as typeof sumBounds;
          const value = state.data.cells.get(
            addrKey({ sheet, row: bound.r0 + dr, col: bound.c0 + dc }),
          )?.value ?? { kind: 'blank' as const };
          return matchesCountIfCriteria(value, pair.criteria);
        });
        if (!matches) continue;
        const sumValue = state.data.cells.get(
          addrKey({ sheet, row: sumBounds.r0 + dr, col: sumBounds.c0 + dc }),
        )?.value;
        if (sumValue?.kind === 'number' && Number.isFinite(sumValue.value)) sum += sumValue.value;
      }
    }
    return { kind: 'number', value: sum };
  };
  const averageMatchingRanges = (
    averageRange: FormulaRangeArg,
    pairs: { range: FormulaRangeArg; criteria: CellValue }[],
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const averageBounds = formulaRangeArgBounds(averageRange, rowOffset, colOffset);
    const nullableBounds = pairs.map((pair) =>
      formulaRangeArgBounds(pair.range, rowOffset, colOffset),
    );
    if (
      !averageBounds ||
      !validRangeBounds(averageBounds) ||
      !nullableBounds.every((bound) => bound !== null && validRangeBounds(bound))
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const bounds = nullableBounds as NonNullable<(typeof nullableBounds)[number]>[];
    if (
      bounds.some(
        (bound) => bound.width !== averageBounds.width || bound.height !== averageBounds.height,
      )
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    let sum = 0;
    let count = 0;
    for (let dr = 0; dr < averageBounds.height; dr += 1) {
      for (let dc = 0; dc < averageBounds.width; dc += 1) {
        const matches = pairs.every((pair, index) => {
          const bound = bounds[index] as typeof averageBounds;
          const value = state.data.cells.get(
            addrKey({ sheet, row: bound.r0 + dr, col: bound.c0 + dc }),
          )?.value ?? { kind: 'blank' as const };
          return matchesCountIfCriteria(value, pair.criteria);
        });
        if (!matches) continue;
        const averageValue = state.data.cells.get(
          addrKey({ sheet, row: averageBounds.r0 + dr, col: averageBounds.c0 + dc }),
        )?.value;
        if (averageValue?.kind === 'number' && Number.isFinite(averageValue.value)) {
          sum += averageValue.value;
          count += 1;
        }
      }
    }
    return count > 0
      ? { kind: 'number', value: sum / count }
      : { kind: 'error', code: 1, text: '#DIV/0!' };
  };
  const minMaxMatchingRanges = (
    valueRange: FormulaRangeArg,
    fn: 'MINIFS' | 'MAXIFS',
    pairs: { range: FormulaRangeArg; criteria: CellValue }[],
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const valueBounds = formulaRangeArgBounds(valueRange, rowOffset, colOffset);
    const nullableBounds = pairs.map((pair) =>
      formulaRangeArgBounds(pair.range, rowOffset, colOffset),
    );
    if (
      !valueBounds ||
      !validRangeBounds(valueBounds) ||
      !nullableBounds.every((bound) => bound !== null && validRangeBounds(bound))
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const bounds = nullableBounds as NonNullable<(typeof nullableBounds)[number]>[];
    if (
      bounds.some(
        (bound) => bound.width !== valueBounds.width || bound.height !== valueBounds.height,
      )
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    let result: number | null = null;
    for (let dr = 0; dr < valueBounds.height; dr += 1) {
      for (let dc = 0; dc < valueBounds.width; dc += 1) {
        const matches = pairs.every((pair, index) => {
          const bound = bounds[index] as typeof valueBounds;
          const value = state.data.cells.get(
            addrKey({ sheet, row: bound.r0 + dr, col: bound.c0 + dc }),
          )?.value ?? { kind: 'blank' as const };
          return matchesCountIfCriteria(value, pair.criteria);
        });
        if (!matches) continue;
        const value = state.data.cells.get(
          addrKey({ sheet, row: valueBounds.r0 + dr, col: valueBounds.c0 + dc }),
        )?.value;
        if (value?.kind !== 'number' || !Number.isFinite(value.value)) continue;
        result =
          result === null
            ? value.value
            : fn === 'MINIFS'
              ? Math.min(result, value.value)
              : Math.max(result, value.value);
      }
    }
    return result === null
      ? { kind: 'error', code: 1, text: '#DIV/0!' }
      : { kind: 'number', value: result };
  };
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
  const roundAwayFromZero = (value: number): number =>
    Math.sign(value) * Math.round(Math.abs(value));
  const roundUpAwayFromZero = (value: number): number =>
    Math.sign(value) * Math.ceil(Math.abs(value));
  const gcdPair = (a: number, b: number): number => {
    let x = Math.abs(a);
    let y = Math.abs(b);
    while (y !== 0) {
      const next = x % y;
      x = y;
      y = next;
    }
    return x;
  };
  const numericFunction = (
    fn:
      | 'ABS'
      | 'MOD'
      | 'ROUND'
      | 'ROUNDUP'
      | 'ROUNDDOWN'
      | 'MROUND'
      | 'QUOTIENT'
      | 'INT'
      | 'TRUNC'
      | 'SQRT'
      | 'POWER'
      | 'PI'
      | 'RADIANS'
      | 'DEGREES'
      | 'SIN'
      | 'COS'
      | 'TAN'
      | 'SEC'
      | 'CSC'
      | 'COT'
      | 'ASIN'
      | 'ACOS'
      | 'ATAN'
      | 'ATAN2'
      | 'ACOT'
      | 'SINH'
      | 'COSH'
      | 'TANH'
      | 'COTH'
      | 'SECH'
      | 'CSCH'
      | 'ASINH'
      | 'ACOSH'
      | 'ATANH'
      | 'ACOTH'
      | 'EXP'
      | 'LN'
      | 'LOG'
      | 'LOG10'
      | 'CHAR'
      | 'CODE'
      | 'UNICHAR'
      | 'UNICODE'
      | 'ADDRESS'
      | 'TYPE'
      | 'ERROR.TYPE'
      | 'FISHER'
      | 'FISHERINV'
      | 'ERF'
      | 'ERF.PRECISE'
      | 'ERFC'
      | 'ERFC.PRECISE'
      | 'GAUSS'
      | 'BASE'
      | 'DECIMAL'
      | 'BIN2DEC'
      | 'DEC2BIN'
      | 'HEX2DEC'
      | 'DEC2HEX'
      | 'OCT2DEC'
      | 'DEC2OCT'
      | 'BIN2HEX'
      | 'HEX2BIN'
      | 'BIN2OCT'
      | 'OCT2BIN'
      | 'HEX2OCT'
      | 'OCT2HEX'
      | 'ROMAN'
      | 'ARABIC'
      | 'DELTA'
      | 'GESTEP'
      | 'BITAND'
      | 'BITOR'
      | 'BITXOR'
      | 'BITLSHIFT'
      | 'BITRSHIFT'
      | 'SQRTPI'
      | 'SUMSQ'
      | 'SIGN'
      | 'GAMMA'
      | 'GAMMALN'
      | 'GAMMALN.PRECISE'
      | 'GCD'
      | 'LCM'
      | 'FACT'
      | 'FACTDOUBLE'
      | 'COMBIN'
      | 'COMBINA'
      | 'PERMUT'
      | 'PERMUTATIONA'
      | 'MULTINOMIAL'
      | 'EVEN'
      | 'ODD'
      | 'STANDARDIZE'
      | 'PHI'
      | 'CONFIDENCE'
      | 'CONFIDENCE.NORM'
      | 'CONFIDENCE.T'
      | 'PMT'
      | 'PV'
      | 'FV'
      | 'NPER'
      | 'RATE'
      | 'IPMT'
      | 'PPMT'
      | 'CUMIPMT'
      | 'CUMPRINC'
      | 'ISPMT'
      | 'EFFECT'
      | 'NOMINAL'
      | 'DOLLARDE'
      | 'DOLLARFR'
      | 'DISC'
      | 'INTRATE'
      | 'PRICEDISC'
      | 'RECEIVED'
      | 'ACCRINTM'
      | 'TBILLPRICE'
      | 'TBILLYIELD'
      | 'TBILLEQ'
      | 'RRI'
      | 'PDURATION'
      | 'SLN'
      | 'SYD'
      | 'DDB'
      | 'DB'
      | 'NORMSDIST'
      | 'NORMDIST'
      | 'NORM.S.DIST'
      | 'NORM.DIST'
      | 'NORMSINV'
      | 'NORM.S.INV'
      | 'NORMINV'
      | 'NORM.INV'
      | 'LOGINV'
      | 'LOGNORM.INV'
      | 'LOGNORMDIST'
      | 'LOGNORM.DIST'
      | 'GAMMADIST'
      | 'GAMMA.DIST'
      | 'GAMMAINV'
      | 'GAMMA.INV'
      | 'BETADIST'
      | 'BETA.DIST'
      | 'BETAINV'
      | 'BETA.INV'
      | 'FDIST'
      | 'F.DIST'
      | 'F.DIST.RT'
      | 'FINV'
      | 'F.INV'
      | 'F.INV.RT'
      | 'TDIST'
      | 'T.DIST'
      | 'T.DIST.2T'
      | 'T.DIST.RT'
      | 'TINV'
      | 'T.INV'
      | 'T.INV.2T'
      | 'CHIDIST'
      | 'CHISQ.DIST'
      | 'CHISQ.DIST.RT'
      | 'CHIINV'
      | 'CHISQ.INV'
      | 'CHISQ.INV.RT'
      | 'WEIBULL'
      | 'WEIBULL.DIST'
      | 'BINOMDIST'
      | 'BINOM.DIST'
      | 'CRITBINOM'
      | 'BINOM.INV'
      | 'NEGBINOMDIST'
      | 'NEGBINOM.DIST'
      | 'HYPGEOMDIST'
      | 'HYPGEOM.DIST'
      | 'POISSON'
      | 'POISSON.DIST'
      | 'EXPONDIST'
      | 'EXPON.DIST'
      | 'CEILING'
      | 'FLOOR'
      | 'CEILING.MATH'
      | 'FLOOR.MATH'
      | 'CEILING.PRECISE'
      | 'FLOOR.PRECISE'
      | 'ISO.CEILING',
    args: FormulaOperand[],
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const numericResult = (value: number): CellValue =>
      Number.isFinite(value)
        ? { kind: 'number', value }
        : { kind: 'error', code: 6, text: '#NUM!' };
    const quoteAddressSheet = (sheet: string): string =>
      /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`;
    const financialType = (value: number): number | null => {
      const type = Math.trunc(value);
      return type === 0 || type === 1 ? type : null;
    };
    const annuityFactor = (rate: number, periods: number): number | null => {
      if (rate === 0) return periods;
      const factor = (1 + rate) ** periods;
      const value = (factor - 1) / rate;
      return Number.isFinite(value) ? value : null;
    };
    const financialFutureValue = (
      rate: number,
      periods: number,
      payment: number,
      presentValue: number,
      type: number,
    ): number => {
      if (rate === 0) return presentValue + payment * periods;
      const factor = (1 + rate) ** periods;
      return presentValue * factor + payment * (1 + rate * type) * ((factor - 1) / rate);
    };
    const financialPayment = (
      rate: number,
      periods: number,
      presentValue: number,
      futureValue: number,
      type: number,
    ): number | null => {
      if (periods === 0) return null;
      if (rate === 0) return -(presentValue + futureValue) / periods;
      const factor = (1 + rate) ** periods;
      const denominator = (1 + rate * type) * (factor - 1);
      if (denominator === 0) return null;
      const value = -((futureValue + presentValue * factor) * rate) / denominator;
      return Number.isFinite(value) ? value : null;
    };
    const depreciationInputs = (
      costOperand: FormulaOperand | undefined,
      salvageOperand: FormulaOperand | undefined,
      lifeOperand: FormulaOperand | undefined,
    ): { cost: number; salvage: number; life: number } | CellValue => {
      if (!costOperand || !salvageOperand || !lifeOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const cost = readNumber(readOperand(costOperand, rowOffset, colOffset));
      const salvage = readNumber(readOperand(salvageOperand, rowOffset, colOffset));
      const life = readNumber(readOperand(lifeOperand, rowOffset, colOffset));
      if (cost === null || salvage === null || life === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (cost < 0 || salvage < 0 || life <= 0) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      return { cost, salvage, life };
    };
    if (fn === 'SLN' || fn === 'SYD' || fn === 'DDB' || fn === 'DB') {
      const [costOperand, salvageOperand, lifeOperand, periodOperand, factorOperand] = args;
      const inputs = depreciationInputs(costOperand, salvageOperand, lifeOperand);
      if ('kind' in inputs) return inputs;
      const { cost, salvage, life } = inputs;
      if (fn === 'SLN') return numericResult((cost - salvage) / life);
      const rawPeriod = readNumber(
        readOperand(periodOperand as FormulaOperand, rowOffset, colOffset),
      );
      if (rawPeriod === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      const period = Math.trunc(rawPeriod);
      if (period < 1 || (fn !== 'DB' && period > life)) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      if (fn === 'SYD') {
        return numericResult(((cost - salvage) * (life - period + 1) * 2) / (life * (life + 1)));
      }
      if (fn === 'DB') {
        const rawMonth =
          factorOperand === undefined
            ? 12
            : readNumber(readOperand(factorOperand, rowOffset, colOffset));
        if (rawMonth === null) return { kind: 'error', code: 15, text: '#VALUE!' };
        const month = Math.trunc(rawMonth);
        if (month < 1 || month > 12 || salvage > cost) {
          return { kind: 'error', code: 6, text: '#NUM!' };
        }
        const maxPeriod = month === 12 ? Math.trunc(life) : Math.trunc(life) + 1;
        if (period > maxPeriod) return { kind: 'error', code: 6, text: '#NUM!' };
        const rate = Math.round((1 - (salvage / cost) ** (1 / life)) * 1000) / 1000;
        let accumulated = 0;
        let depreciation = 0;
        for (let currentPeriod = 1; currentPeriod <= period; currentPeriod += 1) {
          if (currentPeriod === 1) {
            depreciation = cost * rate * (month / 12);
          } else if (currentPeriod === Math.trunc(life) + 1) {
            depreciation = (cost - accumulated) * rate * ((12 - month) / 12);
          } else {
            depreciation = (cost - accumulated) * rate;
          }
          accumulated += depreciation;
        }
        return numericResult(depreciation);
      }
      const rawFactor =
        factorOperand === undefined
          ? 2
          : readNumber(readOperand(factorOperand, rowOffset, colOffset));
      if (rawFactor === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      if (rawFactor <= 0) return { kind: 'error', code: 6, text: '#NUM!' };
      let bookValue = cost;
      let depreciation = 0;
      for (let currentPeriod = 1; currentPeriod <= period; currentPeriod += 1) {
        depreciation = Math.min(bookValue * (rawFactor / life), Math.max(0, bookValue - salvage));
        bookValue -= depreciation;
      }
      return numericResult(depreciation);
    }
    if (fn === 'CUMIPMT' || fn === 'CUMPRINC') {
      const [
        rateOperand,
        periodsOperand,
        presentValueOperand,
        startOperand,
        endOperand,
        typeOperand,
      ] = args;
      if (
        !rateOperand ||
        !periodsOperand ||
        !presentValueOperand ||
        !startOperand ||
        !endOperand ||
        !typeOperand
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const rate = readNumber(readOperand(rateOperand, rowOffset, colOffset));
      const periods = readNumber(readOperand(periodsOperand, rowOffset, colOffset));
      const presentValue = readNumber(readOperand(presentValueOperand, rowOffset, colOffset));
      const rawStart = readNumber(readOperand(startOperand, rowOffset, colOffset));
      const rawEnd = readNumber(readOperand(endOperand, rowOffset, colOffset));
      const rawType = readNumber(readOperand(typeOperand, rowOffset, colOffset));
      if (
        rate === null ||
        periods === null ||
        presentValue === null ||
        rawStart === null ||
        rawEnd === null ||
        rawType === null
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const type = financialType(rawType);
      const startPeriod = Math.trunc(rawStart);
      const endPeriod = Math.trunc(rawEnd);
      if (
        type === null ||
        rate <= 0 ||
        periods <= 0 ||
        presentValue <= 0 ||
        startPeriod < 1 ||
        endPeriod < 1 ||
        startPeriod > endPeriod ||
        endPeriod > periods
      ) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const payment = financialPayment(rate, periods, presentValue, 0, type);
      if (payment === null) return { kind: 'error', code: 6, text: '#NUM!' };
      let balance = presentValue;
      let cumulative = 0;
      for (let currentPeriod = 1; currentPeriod <= endPeriod; currentPeriod += 1) {
        if (type === 1) balance += payment;
        const interest = currentPeriod === 1 && type === 1 ? 0 : -balance * rate;
        const principal = payment - interest;
        if (currentPeriod >= startPeriod) {
          cumulative += fn === 'CUMIPMT' ? interest : principal;
        }
        if (type === 0) balance += principal;
        else balance -= interest;
      }
      return numericResult(cumulative);
    }
    if (fn === 'ISPMT') {
      const [rateOperand, periodOperand, periodsOperand, presentValueOperand] = args;
      if (!rateOperand || !periodOperand || !periodsOperand || !presentValueOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const rate = readNumber(readOperand(rateOperand, rowOffset, colOffset));
      const period = readNumber(readOperand(periodOperand, rowOffset, colOffset));
      const periods = readNumber(readOperand(periodsOperand, rowOffset, colOffset));
      const presentValue = readNumber(readOperand(presentValueOperand, rowOffset, colOffset));
      if (rate === null || period === null || periods === null || presentValue === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (periods === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
      return numericResult((-presentValue * rate * (periods - period)) / periods);
    }
    if (fn === 'EFFECT' || fn === 'NOMINAL') {
      const [rateOperand, periodsOperand] = args;
      if (!rateOperand || !periodsOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const rate = readNumber(readOperand(rateOperand, rowOffset, colOffset));
      const rawPeriods = readNumber(readOperand(periodsOperand, rowOffset, colOffset));
      if (rate === null || rawPeriods === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const periods = Math.trunc(rawPeriods);
      if (rate <= 0 || periods < 1) return { kind: 'error', code: 6, text: '#NUM!' };
      if (fn === 'EFFECT') return numericResult((1 + rate / periods) ** periods - 1);
      return numericResult(periods * ((1 + rate) ** (1 / periods) - 1));
    }
    if (fn === 'DOLLARDE' || fn === 'DOLLARFR') {
      const [dollarOperand, fractionOperand] = args;
      if (!dollarOperand || !fractionOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const dollar = readNumber(readOperand(dollarOperand, rowOffset, colOffset));
      const rawFraction = readNumber(readOperand(fractionOperand, rowOffset, colOffset));
      if (dollar === null || rawFraction === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const fraction = Math.trunc(rawFraction);
      if (fraction < 1) return { kind: 'error', code: 6, text: '#NUM!' };
      const sign = Math.sign(dollar) || 1;
      const absolute = Math.abs(dollar);
      const integer = Math.trunc(absolute);
      const fractional = absolute - integer;
      if (fn === 'DOLLARDE') {
        return numericResult(sign * (integer + (fractional * 100) / fraction));
      }
      return numericResult(sign * (integer + (fractional * fraction) / 100));
    }
    if (fn === 'DISC' || fn === 'INTRATE' || fn === 'PRICEDISC' || fn === 'RECEIVED') {
      const [settlementOperand, maturityOperand, thirdOperand, redemptionOperand, basisOperand] =
        args;
      if (!settlementOperand || !maturityOperand || !thirdOperand || !redemptionOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const settlement = readNumber(readOperand(settlementOperand, rowOffset, colOffset));
      const maturity = readNumber(readOperand(maturityOperand, rowOffset, colOffset));
      const third = readNumber(readOperand(thirdOperand, rowOffset, colOffset));
      const redemption = readNumber(readOperand(redemptionOperand, rowOffset, colOffset));
      const basis =
        basisOperand === undefined
          ? 0
          : readNumber(readOperand(basisOperand, rowOffset, colOffset));
      if (
        settlement === null ||
        maturity === null ||
        third === null ||
        redemption === null ||
        basis === null
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (maturity <= settlement || third <= 0 || redemption <= 0) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const yearFraction = yearFrac(settlement, maturity, basis);
      if (yearFraction.kind !== 'number' || yearFraction.value <= 0) return yearFraction;
      if (fn === 'PRICEDISC') {
        return numericResult(redemption * (1 - third * yearFraction.value));
      }
      if (fn === 'RECEIVED') {
        const denominator = 1 - redemption * yearFraction.value;
        if (denominator === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
        return numericResult(third / denominator);
      }
      const denominator = fn === 'DISC' ? redemption : third;
      return numericResult((redemption - third) / denominator / yearFraction.value);
    }
    if (fn === 'ACCRINTM') {
      const [issueOperand, settlementOperand, rateOperand, parOperand, basisOperand] = args;
      if (!issueOperand || !settlementOperand || !rateOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const issue = readNumber(readOperand(issueOperand, rowOffset, colOffset));
      const settlement = readNumber(readOperand(settlementOperand, rowOffset, colOffset));
      const rate = readNumber(readOperand(rateOperand, rowOffset, colOffset));
      const par =
        parOperand === undefined ? 1000 : readNumber(readOperand(parOperand, rowOffset, colOffset));
      const basis =
        basisOperand === undefined
          ? 0
          : readNumber(readOperand(basisOperand, rowOffset, colOffset));
      if (
        issue === null ||
        settlement === null ||
        rate === null ||
        par === null ||
        basis === null
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (settlement <= issue || rate <= 0 || par <= 0) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const yearFraction = yearFrac(issue, settlement, basis);
      if (yearFraction.kind !== 'number' || yearFraction.value <= 0) return yearFraction;
      return numericResult(par * rate * yearFraction.value);
    }
    if (fn === 'TBILLPRICE' || fn === 'TBILLYIELD' || fn === 'TBILLEQ') {
      const [settlementOperand, maturityOperand, thirdOperand] = args;
      if (!settlementOperand || !maturityOperand || !thirdOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const settlement = readNumber(readOperand(settlementOperand, rowOffset, colOffset));
      const maturity = readNumber(readOperand(maturityOperand, rowOffset, colOffset));
      const third = readNumber(readOperand(thirdOperand, rowOffset, colOffset));
      if (settlement === null || maturity === null || third === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const days = Math.trunc(maturity) - Math.trunc(settlement);
      if (days <= 0 || days > 365 || third <= 0) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      if (fn === 'TBILLPRICE') {
        return numericResult(100 * (1 - (third * days) / 360));
      }
      if (fn === 'TBILLYIELD') {
        return numericResult(((100 - third) / third) * (360 / days));
      }
      const denominator = 360 - third * days;
      if (denominator <= 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
      return numericResult((365 * third) / denominator);
    }
    if (fn === 'RRI' || fn === 'PDURATION') {
      const [periodsOperand, presentValueOperand, futureValueOperand] = args;
      if (!periodsOperand || !presentValueOperand || !futureValueOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const first = readNumber(readOperand(periodsOperand, rowOffset, colOffset));
      const presentValue = readNumber(readOperand(presentValueOperand, rowOffset, colOffset));
      const futureValue = readNumber(readOperand(futureValueOperand, rowOffset, colOffset));
      if (first === null || presentValue === null || futureValue === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (fn === 'RRI') {
        if (first <= 0 || presentValue <= 0 || futureValue <= 0) {
          return { kind: 'error', code: 6, text: '#NUM!' };
        }
        return numericResult((futureValue / presentValue) ** (1 / first) - 1);
      }
      const rate = first;
      if (rate <= 0 || presentValue <= 0 || futureValue <= 0) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      return numericResult(Math.log(futureValue / presentValue) / Math.log(1 + rate));
    }
    if (fn === 'IPMT' || fn === 'PPMT') {
      const [
        rateOperand,
        periodOperand,
        periodsOperand,
        presentValueOperand,
        futureValueOperand,
        typeOperand,
      ] = args;
      if (!rateOperand || !periodOperand || !periodsOperand || !presentValueOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const rate = readNumber(readOperand(rateOperand, rowOffset, colOffset));
      const rawPeriod = readNumber(readOperand(periodOperand, rowOffset, colOffset));
      const periods = readNumber(readOperand(periodsOperand, rowOffset, colOffset));
      const presentValue = readNumber(readOperand(presentValueOperand, rowOffset, colOffset));
      const futureValue =
        futureValueOperand === undefined
          ? 0
          : readNumber(readOperand(futureValueOperand, rowOffset, colOffset));
      const rawType =
        typeOperand === undefined ? 0 : readNumber(readOperand(typeOperand, rowOffset, colOffset));
      if (
        rate === null ||
        rawPeriod === null ||
        periods === null ||
        presentValue === null ||
        futureValue === null ||
        rawType === null
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const type = financialType(rawType);
      const period = Math.trunc(rawPeriod);
      if (type === null || period < 1 || period > periods || periods === 0) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const payment = financialPayment(rate, periods, presentValue, futureValue, type);
      if (payment === null) return { kind: 'error', code: 6, text: '#NUM!' };
      if (rate === 0) {
        const interest = 0;
        return { kind: 'number', value: fn === 'IPMT' ? interest : payment - interest };
      }
      let balance = presentValue;
      let interest = 0;
      for (let currentPeriod = 1; currentPeriod <= period; currentPeriod += 1) {
        if (type === 1) balance += payment;
        interest = currentPeriod === 1 && type === 1 ? 0 : -balance * rate;
        if (type === 0) balance += payment - interest;
        else balance -= interest;
      }
      return numericResult(fn === 'IPMT' ? interest : payment - interest);
    }
    if (fn === 'PMT' || fn === 'PV' || fn === 'FV' || fn === 'NPER' || fn === 'RATE') {
      const [rateOperand, periodsOperand, thirdOperand, fourthOperand, typeOperand, guessOperand] =
        args;
      if (!rateOperand || !periodsOperand || !thirdOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const rate = readNumber(readOperand(rateOperand, rowOffset, colOffset));
      const periods = readNumber(readOperand(periodsOperand, rowOffset, colOffset));
      const third = readNumber(readOperand(thirdOperand, rowOffset, colOffset));
      const fourth =
        fourthOperand === undefined
          ? 0
          : readNumber(readOperand(fourthOperand, rowOffset, colOffset));
      const rawType =
        typeOperand === undefined ? 0 : readNumber(readOperand(typeOperand, rowOffset, colOffset));
      const guess =
        guessOperand === undefined
          ? 0.1
          : readNumber(readOperand(guessOperand, rowOffset, colOffset));
      if (
        rate === null ||
        periods === null ||
        third === null ||
        fourth === null ||
        rawType === null ||
        guess === null
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const type = financialType(rawType);
      if (type === null || (fn === 'RATE' ? rate === 0 : periods === 0)) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      if (fn === 'RATE') {
        const payment = periods;
        const presentValue = third;
        const futureValue = fourth;
        const zeroRateValue =
          financialFutureValue(0, rate, payment, presentValue, type) + futureValue;
        if (Math.abs(zeroRateValue) < 1e-9) return { kind: 'number', value: 0 };
        let current = guess;
        for (let iteration = 0; iteration < 50; iteration += 1) {
          if (current <= -1) return { kind: 'error', code: 6, text: '#NUM!' };
          const value =
            financialFutureValue(current, rate, payment, presentValue, type) + futureValue;
          if (!Number.isFinite(value)) return { kind: 'error', code: 6, text: '#NUM!' };
          if (Math.abs(value) < 1e-9) return numericResult(current);
          const step = Math.max(Math.abs(current) * 1e-6, 1e-7);
          const high =
            financialFutureValue(current + step, rate, payment, presentValue, type) + futureValue;
          const low =
            financialFutureValue(current - step, rate, payment, presentValue, type) + futureValue;
          const derivative = (high - low) / (2 * step);
          if (!Number.isFinite(derivative) || derivative === 0) {
            return { kind: 'error', code: 6, text: '#NUM!' };
          }
          const next = current - value / derivative;
          if (!Number.isFinite(next)) return { kind: 'error', code: 6, text: '#NUM!' };
          if (Math.abs(next - current) < 1e-10) return numericResult(next);
          current = next;
        }
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      if (fn === 'NPER') {
        const payment = periods;
        const presentValue = third;
        const futureValue = fourth;
        if (rate === 0) {
          if (payment === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
          return numericResult(-(presentValue + futureValue) / payment);
        }
        const adjustedPayment = payment * (1 + rate * type);
        const numerator = adjustedPayment - futureValue * rate;
        const denominator = adjustedPayment + presentValue * rate;
        const ratio = numerator / denominator;
        if (ratio <= 0 || rate <= -1) {
          return { kind: 'error', code: 6, text: '#NUM!' };
        }
        return numericResult(Math.log(ratio) / Math.log(1 + rate));
      }
      if (fn === 'PMT') {
        const presentValue = third;
        const futureValue = fourth;
        if (rate === 0) return numericResult(-(presentValue + futureValue) / periods);
        const factor = (1 + rate) ** periods;
        const denominator = (1 + rate * type) * (factor - 1);
        if (denominator === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
        return numericResult(-((futureValue + presentValue * factor) * rate) / denominator);
      }
      const payment = third;
      const factor = (1 + rate) ** periods;
      const annuity = annuityFactor(rate, periods);
      if (annuity === null) return { kind: 'error', code: 6, text: '#NUM!' };
      if (fn === 'FV') {
        const presentValue = fourth;
        return numericResult(-(presentValue * factor + payment * (1 + rate * type) * annuity));
      }
      const futureValue = fourth;
      if (factor === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
      return numericResult(-(futureValue + payment * (1 + rate * type) * annuity) / factor);
    }
    if (fn === 'NORMSDIST') {
      const [zOperand] = args;
      if (!zOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const z = readNumber(readOperand(zOperand, rowOffset, colOffset));
      if (z === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      return numericResult(standardNormalCdf(z));
    }
    if (fn === 'NORM.S.DIST') {
      const [zOperand, cumulativeOperand] = args;
      if (!zOperand || !cumulativeOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const z = readNumber(readOperand(zOperand, rowOffset, colOffset));
      const cumulative = readLogical(readOperand(cumulativeOperand, rowOffset, colOffset));
      if (z === null || cumulative === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      return numericResult(cumulative ? standardNormalCdf(z) : standardNormalPdf(z));
    }
    if (fn === 'CONFIDENCE' || fn === 'CONFIDENCE.NORM' || fn === 'CONFIDENCE.T') {
      const [alphaOperand, standardDeviationOperand, sizeOperand] = args;
      if (!alphaOperand || !standardDeviationOperand || !sizeOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const alpha = readNumber(readOperand(alphaOperand, rowOffset, colOffset));
      const standardDeviation = readNumber(
        readOperand(standardDeviationOperand, rowOffset, colOffset),
      );
      const size = readNumber(readOperand(sizeOperand, rowOffset, colOffset));
      if (alpha === null || standardDeviation === null || size === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const sampleSize = Math.trunc(size);
      if (alpha <= 0 || alpha >= 1 || standardDeviation <= 0 || sampleSize < 1) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      if (fn === 'CONFIDENCE.T') {
        if (sampleSize < 2) return { kind: 'error', code: 6, text: '#NUM!' };
        const critical = inverseStudentTCdf(1 - alpha / 2, sampleSize - 1);
        return critical === null
          ? { kind: 'error', code: 6, text: '#NUM!' }
          : numericResult((critical * standardDeviation) / Math.sqrt(sampleSize));
      }
      return numericResult(
        (inverseStandardNormal(1 - alpha / 2) * standardDeviation) / Math.sqrt(sampleSize),
      );
    }
    if (fn === 'NORMDIST' || fn === 'NORM.DIST') {
      const [xOperand, meanOperand, standardDeviationOperand, cumulativeOperand] = args;
      if (!xOperand || !meanOperand || !standardDeviationOperand || !cumulativeOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const x = readNumber(readOperand(xOperand, rowOffset, colOffset));
      const mean = readNumber(readOperand(meanOperand, rowOffset, colOffset));
      const standardDeviation = readNumber(
        readOperand(standardDeviationOperand, rowOffset, colOffset),
      );
      const cumulative = readLogical(readOperand(cumulativeOperand, rowOffset, colOffset));
      if (x === null || mean === null || standardDeviation === null || cumulative === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (standardDeviation <= 0) return { kind: 'error', code: 6, text: '#NUM!' };
      const z = (x - mean) / standardDeviation;
      return numericResult(
        cumulative ? standardNormalCdf(z) : standardNormalPdf(z) / standardDeviation,
      );
    }
    if (fn === 'NORMSINV' || fn === 'NORM.S.INV') {
      const [probabilityOperand] = args;
      if (!probabilityOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const probability = readNumber(readOperand(probabilityOperand, rowOffset, colOffset));
      if (probability === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      if (probability <= 0 || probability >= 1) return { kind: 'error', code: 6, text: '#NUM!' };
      return numericResult(inverseStandardNormal(probability));
    }
    if (fn === 'NORMINV' || fn === 'NORM.INV' || fn === 'LOGINV' || fn === 'LOGNORM.INV') {
      const [probabilityOperand, meanOperand, standardDeviationOperand] = args;
      if (!probabilityOperand || !meanOperand || !standardDeviationOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const probability = readNumber(readOperand(probabilityOperand, rowOffset, colOffset));
      const mean = readNumber(readOperand(meanOperand, rowOffset, colOffset));
      const standardDeviation = readNumber(
        readOperand(standardDeviationOperand, rowOffset, colOffset),
      );
      if (probability === null || mean === null || standardDeviation === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (probability <= 0 || probability >= 1 || standardDeviation <= 0) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const value = mean + standardDeviation * inverseStandardNormal(probability);
      return numericResult(fn === 'LOGINV' || fn === 'LOGNORM.INV' ? Math.exp(value) : value);
    }
    if (fn === 'LOGNORMDIST' || fn === 'LOGNORM.DIST') {
      const [xOperand, meanOperand, standardDeviationOperand, cumulativeOperand] = args;
      if (!xOperand || !meanOperand || !standardDeviationOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const x = readNumber(readOperand(xOperand, rowOffset, colOffset));
      const mean = readNumber(readOperand(meanOperand, rowOffset, colOffset));
      const standardDeviation = readNumber(
        readOperand(standardDeviationOperand, rowOffset, colOffset),
      );
      const cumulative =
        fn === 'LOGNORMDIST'
          ? true
          : cumulativeOperand
            ? readLogical(readOperand(cumulativeOperand, rowOffset, colOffset))
            : null;
      if (x === null || mean === null || standardDeviation === null || cumulative === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (x <= 0 || standardDeviation <= 0) return { kind: 'error', code: 6, text: '#NUM!' };
      const z = (Math.log(x) - mean) / standardDeviation;
      return numericResult(
        cumulative ? standardNormalCdf(z) : standardNormalPdf(z) / (x * standardDeviation),
      );
    }
    if (fn === 'GAMMADIST' || fn === 'GAMMA.DIST') {
      const [xOperand, alphaOperand, betaOperand, cumulativeOperand] = args;
      if (!xOperand || !alphaOperand || !betaOperand || !cumulativeOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const x = readNumber(readOperand(xOperand, rowOffset, colOffset));
      const alpha = readNumber(readOperand(alphaOperand, rowOffset, colOffset));
      const beta = readNumber(readOperand(betaOperand, rowOffset, colOffset));
      const cumulative = readLogical(readOperand(cumulativeOperand, rowOffset, colOffset));
      if (x === null || alpha === null || beta === null || cumulative === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (x < 0 || alpha <= 0 || beta <= 0) return { kind: 'error', code: 6, text: '#NUM!' };
      if (!cumulative) {
        if (x === 0) {
          if (alpha === 1) return numericResult(1 / beta);
          return alpha > 1
            ? { kind: 'number', value: 0 }
            : { kind: 'error', code: 6, text: '#NUM!' };
        }
        return numericResult(
          Math.exp((alpha - 1) * Math.log(x) - x / beta - alpha * Math.log(beta) - logGamma(alpha)),
        );
      }
      const result = regularizedGammaP(alpha, x / beta);
      return result === null ? { kind: 'error', code: 6, text: '#NUM!' } : numericResult(result);
    }
    if (fn === 'GAMMAINV' || fn === 'GAMMA.INV') {
      const [probabilityOperand, alphaOperand, betaOperand] = args;
      if (!probabilityOperand || !alphaOperand || !betaOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const probability = readNumber(readOperand(probabilityOperand, rowOffset, colOffset));
      const alpha = readNumber(readOperand(alphaOperand, rowOffset, colOffset));
      const beta = readNumber(readOperand(betaOperand, rowOffset, colOffset));
      if (probability === null || alpha === null || beta === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (probability <= 0 || probability >= 1 || alpha <= 0 || beta <= 0) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const result = inverseRegularizedGammaP(alpha, probability);
      return result === null
        ? { kind: 'error', code: 6, text: '#NUM!' }
        : numericResult(result * beta);
    }
    if (fn === 'BETADIST' || fn === 'BETA.DIST') {
      const [xOperand, alphaOperand, betaOperand, cumulativeOperand, lowerOperand, upperOperand] =
        args;
      if (
        !xOperand ||
        !alphaOperand ||
        !betaOperand ||
        (fn === 'BETA.DIST' && !cumulativeOperand)
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const x = readNumber(readOperand(xOperand, rowOffset, colOffset));
      const alpha = readNumber(readOperand(alphaOperand, rowOffset, colOffset));
      const beta = readNumber(readOperand(betaOperand, rowOffset, colOffset));
      const cumulative =
        fn === 'BETADIST'
          ? true
          : readLogical(readOperand(cumulativeOperand as FormulaOperand, rowOffset, colOffset));
      const lower = lowerOperand ? readNumber(readOperand(lowerOperand, rowOffset, colOffset)) : 0;
      const upper = upperOperand ? readNumber(readOperand(upperOperand, rowOffset, colOffset)) : 1;
      if (
        x === null ||
        alpha === null ||
        beta === null ||
        cumulative === null ||
        lower === null ||
        upper === null
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (alpha <= 0 || beta <= 0 || lower >= upper || x < lower || x > upper) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const normalized = (x - lower) / (upper - lower);
      if (!cumulative) {
        if (normalized === 0 || normalized === 1) {
          const edgeDensity =
            normalized === 0 && alpha === 1
              ? Math.exp(logGamma(alpha + beta) - logGamma(alpha) - logGamma(beta)) /
                (upper - lower)
              : normalized === 1 && beta === 1
                ? Math.exp(logGamma(alpha + beta) - logGamma(alpha) - logGamma(beta)) /
                  (upper - lower)
                : null;
          return edgeDensity === null
            ? { kind: 'error', code: 6, text: '#NUM!' }
            : numericResult(edgeDensity);
        }
        return numericResult(
          Math.exp(
            (alpha - 1) * Math.log(normalized) +
              (beta - 1) * Math.log(1 - normalized) +
              logGamma(alpha + beta) -
              logGamma(alpha) -
              logGamma(beta),
          ) /
            (upper - lower),
        );
      }
      const result = regularizedBeta(normalized, alpha, beta);
      return result === null ? { kind: 'error', code: 6, text: '#NUM!' } : numericResult(result);
    }
    if (fn === 'BETAINV' || fn === 'BETA.INV') {
      const [probabilityOperand, alphaOperand, betaOperand, lowerOperand, upperOperand] = args;
      if (!probabilityOperand || !alphaOperand || !betaOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const probability = readNumber(readOperand(probabilityOperand, rowOffset, colOffset));
      const alpha = readNumber(readOperand(alphaOperand, rowOffset, colOffset));
      const beta = readNumber(readOperand(betaOperand, rowOffset, colOffset));
      const lower = lowerOperand ? readNumber(readOperand(lowerOperand, rowOffset, colOffset)) : 0;
      const upper = upperOperand ? readNumber(readOperand(upperOperand, rowOffset, colOffset)) : 1;
      if (
        probability === null ||
        alpha === null ||
        beta === null ||
        lower === null ||
        upper === null
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (probability <= 0 || probability >= 1 || alpha <= 0 || beta <= 0 || lower >= upper) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const result = inverseRegularizedBeta(probability, alpha, beta);
      return result === null
        ? { kind: 'error', code: 6, text: '#NUM!' }
        : numericResult(lower + result * (upper - lower));
    }
    if (fn === 'FDIST' || fn === 'F.DIST' || fn === 'F.DIST.RT') {
      const [xOperand, degrees1Operand, degrees2Operand, cumulativeOperand] = args;
      if (
        !xOperand ||
        !degrees1Operand ||
        !degrees2Operand ||
        (fn === 'F.DIST' && !cumulativeOperand)
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const x = readNumber(readOperand(xOperand, rowOffset, colOffset));
      const degrees1 = readNumber(readOperand(degrees1Operand, rowOffset, colOffset));
      const degrees2 = readNumber(readOperand(degrees2Operand, rowOffset, colOffset));
      const cumulative =
        fn === 'F.DIST'
          ? readLogical(readOperand(cumulativeOperand as FormulaOperand, rowOffset, colOffset))
          : true;
      if (x === null || degrees1 === null || degrees2 === null || cumulative === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const d1 = Math.trunc(degrees1);
      const d2 = Math.trunc(degrees2);
      if (x < 0 || d1 < 1 || d2 < 1) return { kind: 'error', code: 6, text: '#NUM!' };
      const transformed = d1 * x === 0 ? 0 : (d1 * x) / (d1 * x + d2);
      if (!cumulative) {
        if (x === 0) {
          return d1 === 2
            ? { kind: 'number', value: 1 }
            : { kind: 'error', code: 6, text: '#NUM!' };
        }
        const halfD1 = d1 / 2;
        const halfD2 = d2 / 2;
        const logDensity =
          halfD1 * Math.log(d1 / d2) +
          (halfD1 - 1) * Math.log(x) -
          (halfD1 + halfD2) * Math.log(1 + (d1 * x) / d2) +
          logGamma(halfD1 + halfD2) -
          logGamma(halfD1) -
          logGamma(halfD2);
        return numericResult(Math.exp(logDensity));
      }
      const leftTail = regularizedBeta(transformed, d1 / 2, d2 / 2);
      if (leftTail === null) return { kind: 'error', code: 6, text: '#NUM!' };
      return numericResult(fn === 'F.DIST' ? leftTail : 1 - leftTail);
    }
    if (fn === 'FINV' || fn === 'F.INV' || fn === 'F.INV.RT') {
      const [probabilityOperand, degrees1Operand, degrees2Operand] = args;
      if (!probabilityOperand || !degrees1Operand || !degrees2Operand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const probability = readNumber(readOperand(probabilityOperand, rowOffset, colOffset));
      const degrees1 = readNumber(readOperand(degrees1Operand, rowOffset, colOffset));
      const degrees2 = readNumber(readOperand(degrees2Operand, rowOffset, colOffset));
      if (probability === null || degrees1 === null || degrees2 === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const d1 = Math.trunc(degrees1);
      const d2 = Math.trunc(degrees2);
      if (probability <= 0 || probability >= 1 || d1 < 1 || d2 < 1) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const leftTailProbability = fn === 'F.INV' ? probability : 1 - probability;
      const transformed = inverseRegularizedBeta(leftTailProbability, d1 / 2, d2 / 2);
      if (transformed === null || transformed >= 1) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      return numericResult((d2 * transformed) / (d1 * (1 - transformed)));
    }
    if (fn === 'TDIST' || fn === 'T.DIST' || fn === 'T.DIST.2T' || fn === 'T.DIST.RT') {
      const [xOperand, degreesOperand, cumulativeOrTailsOperand] = args;
      if (
        !xOperand ||
        !degreesOperand ||
        ((fn === 'TDIST' || fn === 'T.DIST') && !cumulativeOrTailsOperand)
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const x = readNumber(readOperand(xOperand, rowOffset, colOffset));
      const degrees = readNumber(readOperand(degreesOperand, rowOffset, colOffset));
      if (x === null || degrees === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      const d = Math.trunc(degrees);
      if (d < 1 || ((fn === 'TDIST' || fn === 'T.DIST.2T' || fn === 'T.DIST.RT') && x < 0)) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      if (fn === 'T.DIST') {
        const cumulative = readLogical(
          readOperand(cumulativeOrTailsOperand as FormulaOperand, rowOffset, colOffset),
        );
        if (cumulative === null) return { kind: 'error', code: 15, text: '#VALUE!' };
        if (!cumulative) return numericResult(studentTPdf(x, d));
        const leftTail = studentTCdf(x, d);
        return leftTail === null
          ? { kind: 'error', code: 6, text: '#NUM!' }
          : numericResult(leftTail);
      }
      const leftTail = studentTCdf(x, d);
      if (leftTail === null) return { kind: 'error', code: 6, text: '#NUM!' };
      if (fn === 'T.DIST.RT') return numericResult(1 - leftTail);
      if (fn === 'T.DIST.2T') return numericResult(2 * (1 - leftTail));
      const tails = readNumber(
        readOperand(cumulativeOrTailsOperand as FormulaOperand, rowOffset, colOffset),
      );
      if (tails === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      const tailCount = Math.trunc(tails);
      if (tailCount !== 1 && tailCount !== 2) return { kind: 'error', code: 6, text: '#NUM!' };
      return numericResult(tailCount === 1 ? 1 - leftTail : 2 * (1 - leftTail));
    }
    if (fn === 'TINV' || fn === 'T.INV' || fn === 'T.INV.2T') {
      const [probabilityOperand, degreesOperand] = args;
      if (!probabilityOperand || !degreesOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const probability = readNumber(readOperand(probabilityOperand, rowOffset, colOffset));
      const degrees = readNumber(readOperand(degreesOperand, rowOffset, colOffset));
      if (probability === null || degrees === null)
        return { kind: 'error', code: 15, text: '#VALUE!' };
      const d = Math.trunc(degrees);
      if (probability <= 0 || probability >= 1 || d < 1) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const leftTailProbability = fn === 'T.INV' ? probability : 1 - probability / 2;
      const result = inverseStudentTCdf(leftTailProbability, d);
      return result === null ? { kind: 'error', code: 6, text: '#NUM!' } : numericResult(result);
    }
    if (fn === 'CHIDIST' || fn === 'CHISQ.DIST' || fn === 'CHISQ.DIST.RT') {
      const [xOperand, degreesOperand, cumulativeOperand] = args;
      if (!xOperand || !degreesOperand || (fn === 'CHISQ.DIST' && !cumulativeOperand)) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const x = readNumber(readOperand(xOperand, rowOffset, colOffset));
      const degrees = readNumber(readOperand(degreesOperand, rowOffset, colOffset));
      const cumulative =
        fn === 'CHISQ.DIST'
          ? readLogical(readOperand(cumulativeOperand as FormulaOperand, rowOffset, colOffset))
          : true;
      if (x === null || degrees === null || cumulative === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const df = Math.trunc(degrees);
      if (x < 0 || df < 1) return { kind: 'error', code: 6, text: '#NUM!' };
      const alpha = df / 2;
      if (!cumulative) {
        if (x === 0) {
          if (df === 2) return { kind: 'number', value: 0.5 };
          return df > 2 ? { kind: 'number', value: 0 } : { kind: 'error', code: 6, text: '#NUM!' };
        }
        return numericResult(Math.exp((alpha - 1) * Math.log(x / 2) - x / 2 - logGamma(alpha)) / 2);
      }
      const leftTail = regularizedGammaP(alpha, x / 2);
      if (leftTail === null) return { kind: 'error', code: 6, text: '#NUM!' };
      return numericResult(fn === 'CHISQ.DIST' ? leftTail : 1 - leftTail);
    }
    if (fn === 'CHIINV' || fn === 'CHISQ.INV' || fn === 'CHISQ.INV.RT') {
      const [probabilityOperand, degreesOperand] = args;
      if (!probabilityOperand || !degreesOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const probability = readNumber(readOperand(probabilityOperand, rowOffset, colOffset));
      const degrees = readNumber(readOperand(degreesOperand, rowOffset, colOffset));
      if (probability === null || degrees === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const df = Math.trunc(degrees);
      if (probability <= 0 || probability >= 1 || df < 1) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const leftTailProbability = fn === 'CHISQ.INV' ? probability : 1 - probability;
      const result = inverseRegularizedGammaP(df / 2, leftTailProbability);
      return result === null
        ? { kind: 'error', code: 6, text: '#NUM!' }
        : numericResult(result * 2);
    }
    if (fn === 'BINOMDIST' || fn === 'BINOM.DIST') {
      const [successesOperand, trialsOperand, probabilityOperand, cumulativeOperand] = args;
      if (!successesOperand || !trialsOperand || !probabilityOperand || !cumulativeOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const successes = readNumber(readOperand(successesOperand, rowOffset, colOffset));
      const trials = readNumber(readOperand(trialsOperand, rowOffset, colOffset));
      const probability = readNumber(readOperand(probabilityOperand, rowOffset, colOffset));
      const cumulative = readLogical(readOperand(cumulativeOperand, rowOffset, colOffset));
      if (successes === null || trials === null || probability === null || cumulative === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const successCount = Math.trunc(successes);
      const trialCount = Math.trunc(trials);
      if (
        successCount < 0 ||
        trialCount < 0 ||
        successCount > trialCount ||
        probability < 0 ||
        probability > 1
      ) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      if (!cumulative) {
        return numericResult(binomialProbability(successCount, trialCount, probability));
      }
      let total = 0;
      for (let k = 0; k <= successCount; k += 1) {
        total += binomialProbability(k, trialCount, probability);
      }
      return numericResult(total);
    }
    if (fn === 'CRITBINOM' || fn === 'BINOM.INV') {
      const [trialsOperand, probabilityOperand, alphaOperand] = args;
      if (!trialsOperand || !probabilityOperand || !alphaOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const trials = readNumber(readOperand(trialsOperand, rowOffset, colOffset));
      const probability = readNumber(readOperand(probabilityOperand, rowOffset, colOffset));
      const alpha = readNumber(readOperand(alphaOperand, rowOffset, colOffset));
      if (trials === null || probability === null || alpha === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const trialCount = Math.trunc(trials);
      if (trialCount < 0 || probability < 0 || probability > 1 || alpha <= 0 || alpha >= 1) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      let total = 0;
      for (let k = 0; k <= trialCount; k += 1) {
        total += binomialProbability(k, trialCount, probability);
        if (total >= alpha) return { kind: 'number', value: k };
      }
      return { kind: 'number', value: trialCount };
    }
    if (fn === 'NEGBINOMDIST' || fn === 'NEGBINOM.DIST') {
      const [failuresOperand, successesOperand, probabilityOperand, cumulativeOperand] = args;
      if (!failuresOperand || !successesOperand || !probabilityOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const failures = readNumber(readOperand(failuresOperand, rowOffset, colOffset));
      const successes = readNumber(readOperand(successesOperand, rowOffset, colOffset));
      const probability = readNumber(readOperand(probabilityOperand, rowOffset, colOffset));
      const cumulative =
        fn === 'NEGBINOMDIST'
          ? false
          : cumulativeOperand
            ? readLogical(readOperand(cumulativeOperand, rowOffset, colOffset))
            : null;
      if (failures === null || successes === null || probability === null || cumulative === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const failureCount = Math.trunc(failures);
      const successCount = Math.trunc(successes);
      if (failureCount < 0 || successCount < 1 || probability < 0 || probability > 1) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      if (!cumulative) {
        return numericResult(negativeBinomialProbability(failureCount, successCount, probability));
      }
      let total = 0;
      for (let k = 0; k <= failureCount; k += 1) {
        total += negativeBinomialProbability(k, successCount, probability);
      }
      return numericResult(total);
    }
    if (fn === 'HYPGEOMDIST' || fn === 'HYPGEOM.DIST') {
      const [
        sampleSuccessesOperand,
        sampleSizeOperand,
        populationSuccessesOperand,
        populationSizeOperand,
        cumulativeOperand,
      ] = args;
      if (
        !sampleSuccessesOperand ||
        !sampleSizeOperand ||
        !populationSuccessesOperand ||
        !populationSizeOperand
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const sampleSuccesses = readNumber(readOperand(sampleSuccessesOperand, rowOffset, colOffset));
      const sampleSize = readNumber(readOperand(sampleSizeOperand, rowOffset, colOffset));
      const populationSuccesses = readNumber(
        readOperand(populationSuccessesOperand, rowOffset, colOffset),
      );
      const populationSize = readNumber(readOperand(populationSizeOperand, rowOffset, colOffset));
      const cumulative =
        fn === 'HYPGEOMDIST'
          ? false
          : cumulativeOperand
            ? readLogical(readOperand(cumulativeOperand, rowOffset, colOffset))
            : null;
      if (
        sampleSuccesses === null ||
        sampleSize === null ||
        populationSuccesses === null ||
        populationSize === null ||
        cumulative === null
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const x = Math.trunc(sampleSuccesses);
      const n = Math.trunc(sampleSize);
      const m = Math.trunc(populationSuccesses);
      const bigN = Math.trunc(populationSize);
      const valid =
        x >= 0 && n >= 0 && m >= 0 && bigN >= 0 && x <= n && x <= m && n <= bigN && m <= bigN;
      if (!valid || n - x > bigN - m) return { kind: 'error', code: 6, text: '#NUM!' };
      if (!cumulative) return numericResult(hypergeometricProbability(x, n, m, bigN));
      let total = 0;
      const minSuccess = Math.max(0, n - (bigN - m));
      for (let k = minSuccess; k <= x; k += 1) {
        total += hypergeometricProbability(k, n, m, bigN);
      }
      return numericResult(total);
    }
    if (fn === 'POISSON' || fn === 'POISSON.DIST') {
      const [xOperand, meanOperand, cumulativeOperand] = args;
      if (!xOperand || !meanOperand || !cumulativeOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const x = readNumber(readOperand(xOperand, rowOffset, colOffset));
      const mean = readNumber(readOperand(meanOperand, rowOffset, colOffset));
      const cumulative = readLogical(readOperand(cumulativeOperand, rowOffset, colOffset));
      if (x === null || mean === null || cumulative === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const count = Math.trunc(x);
      if (count < 0 || mean < 0) return { kind: 'error', code: 6, text: '#NUM!' };
      if (!cumulative) return numericResult(poissonProbability(count, mean));
      let total = 0;
      for (let k = 0; k <= count; k += 1) total += poissonProbability(k, mean);
      return numericResult(total);
    }
    if (fn === 'EXPONDIST' || fn === 'EXPON.DIST') {
      const [xOperand, lambdaOperand, cumulativeOperand] = args;
      if (!xOperand || !lambdaOperand || !cumulativeOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const x = readNumber(readOperand(xOperand, rowOffset, colOffset));
      const lambda = readNumber(readOperand(lambdaOperand, rowOffset, colOffset));
      const cumulative = readLogical(readOperand(cumulativeOperand, rowOffset, colOffset));
      if (x === null || lambda === null || cumulative === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (x < 0 || lambda <= 0) return { kind: 'error', code: 6, text: '#NUM!' };
      return numericResult(cumulative ? 1 - Math.exp(-lambda * x) : lambda * Math.exp(-lambda * x));
    }
    if (fn === 'WEIBULL' || fn === 'WEIBULL.DIST') {
      const [xOperand, alphaOperand, betaOperand, cumulativeOperand] = args;
      if (!xOperand || !alphaOperand || !betaOperand || !cumulativeOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const x = readNumber(readOperand(xOperand, rowOffset, colOffset));
      const alpha = readNumber(readOperand(alphaOperand, rowOffset, colOffset));
      const beta = readNumber(readOperand(betaOperand, rowOffset, colOffset));
      const cumulative = readLogical(readOperand(cumulativeOperand, rowOffset, colOffset));
      if (x === null || alpha === null || beta === null || cumulative === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (x < 0 || alpha <= 0 || beta <= 0) return { kind: 'error', code: 6, text: '#NUM!' };
      const scaled = (x / beta) ** alpha;
      return numericResult(
        cumulative
          ? 1 - Math.exp(-scaled)
          : (alpha / beta) * (x / beta) ** (alpha - 1) * Math.exp(-scaled),
      );
    }
    if (fn === 'ADDRESS') {
      const [rowOperand, colOperand, absOperand, a1Operand, sheetOperand] = args;
      if (!rowOperand || !colOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const rowValue = readNumber(readOperand(rowOperand, rowOffset, colOffset));
      const colValue = readNumber(readOperand(colOperand, rowOffset, colOffset));
      const absValue = absOperand ? readNumber(readOperand(absOperand, rowOffset, colOffset)) : 1;
      const a1Value = a1Operand ? readLogical(readOperand(a1Operand, rowOffset, colOffset)) : true;
      const sheetValue = sheetOperand
        ? textValue(readOperand(sheetOperand, rowOffset, colOffset))
        : '';
      if (
        rowValue === null ||
        colValue === null ||
        absValue === null ||
        a1Value === null ||
        sheetValue === null
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const row = Math.trunc(rowValue);
      const col = Math.trunc(colValue);
      const abs = Math.trunc(absValue);
      if (row < 1 || row > 1048576 || col < 1 || col > 16384 || abs < 1 || abs > 4) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const absoluteCol = abs === 1 || abs === 3;
      const absoluteRow = abs === 1 || abs === 2;
      const ref = a1Value
        ? `${absoluteCol ? '$' : ''}${colToLetters(col - 1)}${absoluteRow ? '$' : ''}${row}`
        : `${absoluteRow ? `R${row}` : `R[${row}]`}${absoluteCol ? `C${col}` : `C[${col}]`}`;
      return {
        kind: 'text',
        value: sheetValue === '' ? ref : `${quoteAddressSheet(sheetValue)}!${ref}`,
      };
    }
    if (fn === 'DECIMAL') {
      const [textOperand, radixOperand] = args;
      if (!textOperand || !radixOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const text = textValue(readOperand(textOperand, rowOffset, colOffset));
      const radix = readNumber(readOperand(radixOperand, rowOffset, colOffset));
      if (text === null || radix === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      const base = Math.trunc(radix);
      if (base < 2 || base > 36) return { kind: 'error', code: 6, text: '#NUM!' };
      const normalized = text.trim().toUpperCase();
      if (normalized === '') return { kind: 'error', code: 6, text: '#NUM!' };
      let result = 0;
      for (const char of normalized) {
        const digit = baseDigits.indexOf(char);
        if (digit < 0 || digit >= base) return { kind: 'error', code: 6, text: '#NUM!' };
        result = result * base + digit;
      }
      return numericResult(result);
    }
    if (fn === 'BIN2DEC') {
      const [textOperand] = args;
      if (!textOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const text = textValue(readOperand(textOperand, rowOffset, colOffset));
      if (text === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      const value = engineeringBaseValue(text, 2);
      return value === null ? { kind: 'error', code: 6, text: '#NUM!' } : numericResult(value);
    }
    if (fn === 'HEX2DEC' || fn === 'OCT2DEC') {
      const [textOperand] = args;
      if (!textOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const text = textValue(readOperand(textOperand, rowOffset, colOffset));
      if (text === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      const base = fn === 'HEX2DEC' ? 16 : 8;
      const value = engineeringBaseValue(text, base);
      return value === null ? { kind: 'error', code: 6, text: '#NUM!' } : numericResult(value);
    }
    if (
      fn === 'BIN2HEX' ||
      fn === 'HEX2BIN' ||
      fn === 'BIN2OCT' ||
      fn === 'OCT2BIN' ||
      fn === 'HEX2OCT' ||
      fn === 'OCT2HEX'
    ) {
      const [textOperand, placesOperand] = args;
      if (!textOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const text = textValue(readOperand(textOperand, rowOffset, colOffset));
      const rawPlaces =
        placesOperand === undefined
          ? undefined
          : readNumber(readOperand(placesOperand, rowOffset, colOffset));
      if (text === null || rawPlaces === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      const fromBase = fn.startsWith('BIN') ? 2 : fn.startsWith('HEX') ? 16 : 8;
      const toBase = fn.endsWith('BIN') ? 2 : fn.endsWith('HEX') ? 16 : 8;
      const value = engineeringBaseValue(text, fromBase);
      const output =
        value === null
          ? null
          : engineeringBaseText(
              value,
              toBase,
              rawPlaces === undefined ? null : Math.trunc(rawPlaces),
            );
      return output === null
        ? { kind: 'error', code: 6, text: '#NUM!' }
        : { kind: 'text', value: output };
    }
    if (fn === 'ARABIC') {
      const [textOperand] = args;
      if (!textOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const text = textValue(readOperand(textOperand, rowOffset, colOffset));
      if (text === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      const value = romanValue(text);
      return value === null
        ? { kind: 'error', code: 15, text: '#VALUE!' }
        : { kind: 'number', value };
    }
    if (fn === 'CODE' || fn === 'UNICODE') {
      const [textOperand] = args;
      if (!textOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const text = textValue(readOperand(textOperand, rowOffset, colOffset));
      const value = text?.codePointAt(0);
      return value === undefined
        ? { kind: 'error', code: 15, text: '#VALUE!' }
        : { kind: 'number', value };
    }
    if (fn === 'TYPE') {
      const [valueOperand] = args;
      if (!valueOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const value = readOperand(valueOperand, rowOffset, colOffset);
      const type =
        value.kind === 'text' ? 2 : value.kind === 'bool' ? 4 : value.kind === 'error' ? 16 : 1;
      return { kind: 'number', value: type };
    }
    if (fn === 'ERROR.TYPE') {
      const [valueOperand] = args;
      if (!valueOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const value = readOperand(valueOperand, rowOffset, colOffset);
      if (value.kind !== 'error') return { kind: 'error', code: 6, text: '#N/A' };
      const errorTypes: Record<string, number> = {
        '#NULL!': 1,
        '#DIV/0!': 2,
        '#VALUE!': 3,
        '#REF!': 4,
        '#NAME?': 5,
        '#NUM!': 6,
        '#N/A': 7,
        '#GETTING_DATA': 8,
        '#SPILL!': 9,
        '#CALC!': 14,
      };
      const type = errorTypes[value.text];
      return type === undefined
        ? { kind: 'error', code: 6, text: '#N/A' }
        : { kind: 'number', value: type };
    }
    const values = args.map((arg) => readNumber(readOperand(arg, rowOffset, colOffset)));
    if (values.some((value) => value === null)) return { kind: 'error', code: 15, text: '#VALUE!' };
    const first = values[0] as number;
    if (fn === 'PI') return { kind: 'number', value: Math.PI };
    if (fn === 'CHAR' || fn === 'UNICHAR') {
      const codePoint = Math.trunc(first);
      const max = fn === 'CHAR' ? 255 : 0x10ffff;
      if (
        codePoint < 1 ||
        codePoint > max ||
        (fn === 'UNICHAR' && codePoint >= 0xd800 && codePoint <= 0xdfff)
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      return { kind: 'text', value: String.fromCodePoint(codePoint) };
    }
    if (fn === 'ABS') return { kind: 'number', value: Math.abs(first) };
    if (fn === 'RADIANS') return { kind: 'number', value: (first * Math.PI) / 180 };
    if (fn === 'DEGREES') return { kind: 'number', value: (first * 180) / Math.PI };
    if (fn === 'SIN') return { kind: 'number', value: Math.sin(first) };
    if (fn === 'COS') return { kind: 'number', value: Math.cos(first) };
    if (fn === 'TAN') return numericResult(Math.tan(first));
    if (fn === 'SEC') return numericResult(1 / Math.cos(first));
    if (fn === 'CSC') {
      const sine = Math.sin(first);
      if (sine === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
      return numericResult(1 / sine);
    }
    if (fn === 'COT') {
      const tangent = Math.tan(first);
      if (tangent === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
      return numericResult(1 / tangent);
    }
    if (fn === 'ASIN' || fn === 'ACOS') {
      if (first < -1 || first > 1) return { kind: 'error', code: 6, text: '#NUM!' };
      return { kind: 'number', value: fn === 'ASIN' ? Math.asin(first) : Math.acos(first) };
    }
    if (fn === 'ATAN') return { kind: 'number', value: Math.atan(first) };
    if (fn === 'ACOT') {
      if (first === 0) return { kind: 'number', value: Math.PI / 2 };
      const value = Math.atan(1 / first);
      return { kind: 'number', value: value < 0 ? value + Math.PI : value };
    }
    if (fn === 'ATAN2') {
      const second = values[1] as number;
      if (first === 0 && second === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
      return { kind: 'number', value: Math.atan2(second, first) };
    }
    if (fn === 'SINH') return numericResult(Math.sinh(first));
    if (fn === 'COSH') return numericResult(Math.cosh(first));
    if (fn === 'TANH') return { kind: 'number', value: Math.tanh(first) };
    if (fn === 'COTH') {
      const tangent = Math.tanh(first);
      if (tangent === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
      return numericResult(1 / tangent);
    }
    if (fn === 'SECH') return numericResult(1 / Math.cosh(first));
    if (fn === 'CSCH') {
      const sine = Math.sinh(first);
      if (sine === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
      return numericResult(1 / sine);
    }
    if (fn === 'ASINH') return numericResult(Math.asinh(first));
    if (fn === 'ACOSH') {
      if (first < 1) return { kind: 'error', code: 6, text: '#NUM!' };
      return numericResult(Math.acosh(first));
    }
    if (fn === 'ATANH') {
      if (first <= -1 || first >= 1) return { kind: 'error', code: 6, text: '#NUM!' };
      return numericResult(Math.atanh(first));
    }
    if (fn === 'ACOTH') {
      if (Math.abs(first) <= 1) return { kind: 'error', code: 6, text: '#NUM!' };
      return numericResult(0.5 * Math.log((first + 1) / (first - 1)));
    }
    if (fn === 'EXP') return numericResult(Math.exp(first));
    if (fn === 'LN' || fn === 'LOG10') {
      if (first <= 0) return { kind: 'error', code: 6, text: '#NUM!' };
      return { kind: 'number', value: fn === 'LN' ? Math.log(first) : Math.log10(first) };
    }
    if (fn === 'FISHER') {
      if (first <= -1 || first >= 1) return { kind: 'error', code: 6, text: '#NUM!' };
      return numericResult(0.5 * Math.log((1 + first) / (1 - first)));
    }
    if (fn === 'FISHERINV') {
      const exponent = Math.exp(2 * first);
      return numericResult((exponent - 1) / (exponent + 1));
    }
    if (fn === 'ERF') {
      const upper = values[1] as number | undefined;
      return numericResult(upper === undefined ? erf(first) : erf(upper) - erf(first));
    }
    if (fn === 'ERF.PRECISE') return numericResult(erf(first));
    if (fn === 'ERFC' || fn === 'ERFC.PRECISE') return numericResult(1 - erf(first));
    if (fn === 'GAUSS') return numericResult(standardNormalCdf(first) - 0.5);
    if (fn === 'BASE') {
      const number = Math.trunc(first);
      const radix = Math.trunc(values[1] as number);
      const minLength = Math.trunc((values[2] as number | undefined) ?? 0);
      if (number < 0 || radix < 2 || radix > 36 || minLength < 0) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      let remaining = number;
      let text = '';
      do {
        text = `${baseDigits[remaining % radix] ?? ''}${text}`;
        remaining = Math.trunc(remaining / radix);
      } while (remaining > 0);
      return { kind: 'text', value: text.padStart(minLength, '0') };
    }
    if (fn === 'DEC2BIN') {
      const number = Math.trunc(first);
      const places = values[1] === undefined ? null : Math.trunc(values[1] as number);
      if (number < -512 || number > 511 || (places !== null && places < 0)) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      if (number < 0) return { kind: 'text', value: (number + 1024).toString(2) };
      const text = number.toString(2);
      if (places !== null && text.length > places) return { kind: 'error', code: 6, text: '#NUM!' };
      return { kind: 'text', value: places === null ? text : text.padStart(places, '0') };
    }
    if (fn === 'DEC2HEX' || fn === 'DEC2OCT') {
      const number = Math.trunc(first);
      const places = values[1] === undefined ? null : Math.trunc(values[1] as number);
      const base = fn === 'DEC2HEX' ? 16 : 8;
      const negativeLimit = fn === 'DEC2HEX' ? -(16 ** 9) : -(8 ** 9);
      const positiveLimit = fn === 'DEC2HEX' ? 16 ** 9 - 1 : 8 ** 9 - 1;
      if (number < negativeLimit || number > positiveLimit || (places !== null && places < 0)) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      if (number < 0)
        return { kind: 'text', value: (number + base ** 10).toString(base).toUpperCase() };
      const text = number.toString(base).toUpperCase();
      if (places !== null && text.length > places) return { kind: 'error', code: 6, text: '#NUM!' };
      return { kind: 'text', value: places === null ? text : text.padStart(places, '0') };
    }
    if (fn === 'ROMAN') {
      const number = Math.trunc(first);
      const form = Math.trunc((values[1] as number | undefined) ?? 0);
      if (number < 1 || number > 3999 || form < 0 || form > 4) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      if (form !== 0) return { kind: 'error', code: 15, text: '#VALUE!' };
      return { kind: 'text', value: romanText(number) };
    }
    if (fn === 'DELTA' || fn === 'GESTEP') {
      const second = (values[1] as number | undefined) ?? 0;
      return {
        kind: 'number',
        value: fn === 'DELTA' ? (first === second ? 1 : 0) : first >= second ? 1 : 0,
      };
    }
    if (
      fn === 'BITAND' ||
      fn === 'BITOR' ||
      fn === 'BITXOR' ||
      fn === 'BITLSHIFT' ||
      fn === 'BITRSHIFT'
    ) {
      const left = bitOperand(first);
      const second = values[1] as number;
      if (left === null) return { kind: 'error', code: 6, text: '#NUM!' };
      let result: bigint;
      if (fn === 'BITAND' || fn === 'BITOR' || fn === 'BITXOR') {
        const right = bitOperand(second);
        if (right === null) return { kind: 'error', code: 6, text: '#NUM!' };
        result = fn === 'BITAND' ? left & right : fn === 'BITOR' ? left | right : left ^ right;
      } else {
        const shift = Math.trunc(second);
        if (Math.abs(shift) > 53) return { kind: 'error', code: 6, text: '#NUM!' };
        const amount = BigInt(Math.abs(shift));
        const shiftLeft = fn === 'BITLSHIFT' ? shift >= 0 : shift < 0;
        result = shiftLeft ? left << amount : left >> amount;
      }
      const number = Number(result);
      return number < 0 || number > maxBitValue || !Number.isFinite(number)
        ? { kind: 'error', code: 6, text: '#NUM!' }
        : { kind: 'number', value: number };
    }
    if (fn === 'SQRTPI') {
      if (first < 0) return { kind: 'error', code: 6, text: '#NUM!' };
      return { kind: 'number', value: Math.sqrt(first * Math.PI) };
    }
    if (fn === 'SUMSQ') {
      return {
        kind: 'number',
        value: values.reduce<number>((sum, value) => sum + (value as number) ** 2, 0),
      };
    }
    if (fn === 'INT') return { kind: 'number', value: Math.floor(first) };
    if (fn === 'SIGN') return { kind: 'number', value: Math.sign(first) };
    if (fn === 'GAMMA') {
      const result = gamma(first);
      return result === null ? { kind: 'error', code: 6, text: '#NUM!' } : numericResult(result);
    }
    if (fn === 'GAMMALN' || fn === 'GAMMALN.PRECISE') {
      if (first <= 0) return { kind: 'error', code: 6, text: '#NUM!' };
      return numericResult(logGamma(first));
    }
    if (fn === 'FACT' || fn === 'FACTDOUBLE') {
      const integer = Math.trunc(first);
      if (integer < 0) return { kind: 'error', code: 6, text: '#NUM!' };
      return numericResult(fn === 'FACT' ? factorial(integer) : doubleFactorial(integer));
    }
    if (fn === 'STANDARDIZE') {
      const mean = values[1] as number;
      const standardDeviation = values[2] as number;
      if (standardDeviation <= 0) return { kind: 'error', code: 6, text: '#NUM!' };
      return { kind: 'number', value: (first - mean) / standardDeviation };
    }
    if (fn === 'PHI') return numericResult(standardNormalPdf(first));
    if (fn === 'COMBIN' || fn === 'COMBINA' || fn === 'PERMUT' || fn === 'PERMUTATIONA') {
      const n = Math.trunc(first);
      const k = Math.trunc(values[1] as number);
      if (n < 0 || k < 0) return { kind: 'error', code: 6, text: '#NUM!' };
      if (fn === 'COMBINA') {
        if (n === 0 && k > 0) return { kind: 'error', code: 6, text: '#NUM!' };
        return numericResult(k === 0 ? 1 : combination(n + k - 1, k));
      }
      if (fn === 'PERMUTATIONA') return numericResult(n ** k);
      if (k > n) return { kind: 'error', code: 6, text: '#NUM!' };
      if (fn === 'COMBIN') return numericResult(combination(n, k));
      return numericResult(factorial(n) / factorial(n - k));
    }
    if (fn === 'MULTINOMIAL') {
      const integers = values.map((value) => Math.trunc(value as number));
      if (integers.some((value) => value < 0)) return { kind: 'error', code: 6, text: '#NUM!' };
      const total = integers.reduce((sum, value) => sum + value, 0);
      return numericResult(
        factorial(total) / integers.reduce((product, value) => product * factorial(value), 1),
      );
    }
    if (fn === 'GCD' || fn === 'LCM') {
      const integers = values.map((value) => Math.trunc(value as number));
      if (integers.some((value) => value < 0)) return { kind: 'error', code: 6, text: '#NUM!' };
      if (fn === 'GCD') {
        return { kind: 'number', value: integers.reduce((acc, value) => gcdPair(acc, value), 0) };
      }
      if (integers.some((value) => value === 0)) return { kind: 'number', value: 0 };
      return {
        kind: 'number',
        value: integers.reduce((acc, value) => Math.abs(acc * value) / gcdPair(acc, value), 1),
      };
    }
    if (fn === 'EVEN' || fn === 'ODD') {
      const magnitude = Math.ceil(Math.abs(first));
      const isEven = magnitude % 2 === 0;
      const rounded =
        fn === 'EVEN' ? (isEven ? magnitude : magnitude + 1) : isEven ? magnitude + 1 : magnitude;
      return { kind: 'number', value: Math.sign(first) * rounded };
    }
    if (
      fn === 'CEILING.MATH' ||
      fn === 'FLOOR.MATH' ||
      fn === 'CEILING.PRECISE' ||
      fn === 'FLOOR.PRECISE' ||
      fn === 'ISO.CEILING'
    ) {
      const significance = Math.abs((values[1] as number | undefined) ?? 1);
      const mode =
        fn === 'CEILING.PRECISE' || fn === 'FLOOR.PRECISE' || fn === 'ISO.CEILING'
          ? 0
          : Math.trunc((values[2] as number | undefined) ?? 0);
      if (significance === 0) return { kind: 'number', value: 0 };
      const scaled = Math.abs(first) / significance;
      const rounded =
        fn === 'CEILING.MATH' || fn === 'CEILING.PRECISE' || fn === 'ISO.CEILING'
          ? first < 0 && mode === 0
            ? Math.floor(scaled)
            : Math.ceil(scaled)
          : first < 0 && mode === 0
            ? Math.ceil(scaled)
            : Math.floor(scaled);
      return { kind: 'number', value: Math.sign(first) * rounded * significance };
    }
    if (fn === 'CEILING' || fn === 'FLOOR') {
      const significance = values[1] as number;
      if (significance === 0) return { kind: 'number', value: 0 };
      if (Math.sign(first) !== Math.sign(significance)) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const scaled = first / significance;
      const rounded = fn === 'CEILING' ? Math.ceil(scaled) : Math.floor(scaled);
      return { kind: 'number', value: rounded * significance };
    }
    if (fn === 'MROUND') {
      const multiple = values[1] as number;
      if (multiple === 0) return { kind: 'number', value: 0 };
      if (Math.sign(first) !== Math.sign(multiple)) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      return {
        kind: 'number',
        value:
          Math.sign(first) * Math.round(Math.abs(first) / Math.abs(multiple)) * Math.abs(multiple),
      };
    }
    if (fn === 'SQRT') {
      return first < 0
        ? { kind: 'error', code: 6, text: '#NUM!' }
        : { kind: 'number', value: Math.sqrt(first) };
    }
    if (fn === 'TRUNC') {
      const digits = Math.trunc((values[1] as number | undefined) ?? 0);
      if (digits >= 0) {
        const factor = 10 ** digits;
        return { kind: 'number', value: Math.trunc(first * factor) / factor };
      }
      const factor = 10 ** -digits;
      return { kind: 'number', value: Math.trunc(first / factor) * factor };
    }
    const second = values[1] as number;
    if (fn === 'QUOTIENT') {
      if (second === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
      return { kind: 'number', value: Math.trunc(first / second) };
    }
    if (fn === 'MOD') {
      if (second === 0) return { kind: 'error', code: 1, text: '#DIV/0!' };
      return { kind: 'number', value: first - second * Math.floor(first / second) };
    }
    if (fn === 'POWER') {
      const value = first ** second;
      return Number.isFinite(value)
        ? { kind: 'number', value }
        : { kind: 'error', code: 6, text: '#NUM!' };
    }
    if (fn === 'LOG') {
      const base = second ?? 10;
      if (first <= 0 || base <= 0 || base === 1) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      return { kind: 'number', value: Math.log(first) / Math.log(base) };
    }
    const digits = Math.trunc(second);
    if (fn === 'ROUNDDOWN') {
      if (digits >= 0) {
        const factor = 10 ** digits;
        return { kind: 'number', value: Math.trunc(first * factor) / factor };
      }
      const factor = 10 ** -digits;
      return { kind: 'number', value: Math.trunc(first / factor) * factor };
    }
    if (fn === 'ROUNDUP') {
      if (digits >= 0) {
        const factor = 10 ** digits;
        return { kind: 'number', value: roundUpAwayFromZero(first * factor) / factor };
      }
      const factor = 10 ** -digits;
      return { kind: 'number', value: roundUpAwayFromZero(first / factor) * factor };
    }
    if (digits >= 0) {
      const factor = 10 ** digits;
      return { kind: 'number', value: roundAwayFromZero(first * factor) / factor };
    }
    const factor = 10 ** -digits;
    return { kind: 'number', value: roundAwayFromZero(first / factor) * factor };
  };
  const numericPredicate = (
    fn: 'ISEVEN' | 'ISODD',
    value: FormulaOperand,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const number = readNumber(readOperand(value, rowOffset, colOffset));
    if (number === null) return { kind: 'error', code: 15, text: '#VALUE!' };
    const even = Math.abs(Math.trunc(number)) % 2 === 0;
    return { kind: 'bool', value: fn === 'ISEVEN' ? even : !even };
  };
  const dateFunction = (
    fn:
      | 'DATE'
      | 'YEAR'
      | 'MONTH'
      | 'DAY'
      | 'WEEKDAY'
      | 'WEEKNUM'
      | 'ISOWEEKNUM'
      | 'TODAY'
      | 'NOW'
      | 'TIME'
      | 'EDATE'
      | 'EOMONTH'
      | 'DAYS'
      | 'DAYS360'
      | 'DATEDIF'
      | 'YEARFRAC'
      | 'DATEVALUE'
      | 'TIMEVALUE'
      | 'NETWORKDAYS'
      | 'NETWORKDAYS.INTL'
      | 'WORKDAY'
      | 'WORKDAY.INTL'
      | 'HOUR'
      | 'MINUTE'
      | 'SECOND',
    args: FormulaDateArg[],
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    if (fn === 'TODAY') return { kind: 'number', value: todaySerial() };
    if (fn === 'NOW') return { kind: 'number', value: Date.now() / 86_400_000 + 25569 };
    const asOperand = (arg: FormulaDateArg): FormulaOperand | null =>
      arg.kind === 'range' || arg.kind === 'dynamic-range' ? null : arg;
    if (fn === 'DATEVALUE' || fn === 'TIMEVALUE') {
      const [arg] = args;
      if (!arg) return { kind: 'error', code: 15, text: '#VALUE!' };
      const operand = asOperand(arg);
      if (!operand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const value = readOperand(operand, rowOffset, colOffset);
      return fn === 'DATEVALUE' ? dateValueText(value) : timeValueText(value);
    }
    if (fn === 'DATEDIF') {
      const [startArg, endArg, unitArg] = args;
      if (!startArg || !endArg || !unitArg) return { kind: 'error', code: 15, text: '#VALUE!' };
      const startOperand = asOperand(startArg);
      const endOperand = asOperand(endArg);
      const unitOperand = asOperand(unitArg);
      if (!startOperand || !endOperand || !unitOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const start = readNumber(readOperand(startOperand, rowOffset, colOffset));
      const end = readNumber(readOperand(endOperand, rowOffset, colOffset));
      if (start === null || end === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      return datedif(start, end, readOperand(unitOperand, rowOffset, colOffset));
    }
    if (fn === 'DAYS360') {
      const [startArg, endArg, methodArg] = args;
      if (!startArg || !endArg) return { kind: 'error', code: 15, text: '#VALUE!' };
      const startOperand = asOperand(startArg);
      const endOperand = asOperand(endArg);
      const methodOperand = methodArg ? asOperand(methodArg) : undefined;
      if (!startOperand || !endOperand || methodOperand === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const start = readNumber(readOperand(startOperand, rowOffset, colOffset));
      const end = readNumber(readOperand(endOperand, rowOffset, colOffset));
      if (start === null || end === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      const method =
        methodOperand === undefined
          ? false
          : readLogical(readOperand(methodOperand, rowOffset, colOffset));
      if (method === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      return days360(start, end, method);
    }
    if (fn === 'YEARFRAC') {
      const [startArg, endArg, basisArg] = args;
      if (!startArg || !endArg) return { kind: 'error', code: 15, text: '#VALUE!' };
      const startOperand = asOperand(startArg);
      const endOperand = asOperand(endArg);
      const basisOperand = basisArg ? asOperand(basisArg) : undefined;
      if (!startOperand || !endOperand || basisOperand === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const start = readNumber(readOperand(startOperand, rowOffset, colOffset));
      const end = readNumber(readOperand(endOperand, rowOffset, colOffset));
      if (start === null || end === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      const basis =
        basisOperand === undefined
          ? 0
          : readNumber(readOperand(basisOperand, rowOffset, colOffset));
      if (basis === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      return yearFrac(start, end, basis);
    }
    if (
      fn === 'NETWORKDAYS' ||
      fn === 'WORKDAY' ||
      fn === 'NETWORKDAYS.INTL' ||
      fn === 'WORKDAY.INTL'
    ) {
      const [startArg, endOrDaysArg, thirdArg, fourthArg] = args;
      const startOperand = startArg ? asOperand(startArg) : null;
      const endOrDaysOperand = endOrDaysArg ? asOperand(endOrDaysArg) : null;
      if (!startOperand || !endOrDaysOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const start = readNumber(readOperand(startOperand, rowOffset, colOffset));
      const endOrDays = readNumber(readOperand(endOrDaysOperand, rowOffset, colOffset));
      if (start === null || endOrDays === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      const isIntl = fn === 'NETWORKDAYS.INTL' || fn === 'WORKDAY.INTL';
      let weekends = defaultWeekendDays;
      if (isIntl && thirdArg) {
        const weekendOperand = asOperand(thirdArg);
        if (!weekendOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
        const parsedWeekends = weekendDaysFromValue(
          readOperand(weekendOperand, rowOffset, colOffset),
        );
        if (parsedWeekends === null) return { kind: 'error', code: 6, text: '#NUM!' };
        weekends = parsedWeekends;
      }
      const holidaysArg = isIntl ? fourthArg : thirdArg;
      const holidays = new Set<number>();
      if (holidaysArg) {
        if (holidaysArg.kind === 'range' || holidaysArg.kind === 'dynamic-range') {
          const values = numericValuesInFormulaRangeArg(holidaysArg, rowOffset, colOffset);
          if (values === null) return { kind: 'error', code: 15, text: '#VALUE!' };
          for (const value of values) holidays.add(Math.trunc(value));
        } else {
          const value = readNumber(readOperand(holidaysArg, rowOffset, colOffset));
          if (value === null) return { kind: 'error', code: 15, text: '#VALUE!' };
          holidays.add(Math.trunc(value));
        }
      }
      return {
        kind: 'number',
        value:
          fn === 'NETWORKDAYS' || fn === 'NETWORKDAYS.INTL'
            ? networkDays(start, endOrDays, holidays, weekends)
            : workday(start, endOrDays, holidays, weekends),
      };
    }
    const operandArgs = args as FormulaOperand[];
    const values = operandArgs.map((arg) => readNumber(readOperand(arg, rowOffset, colOffset)));
    if (values.some((value) => value === null)) return { kind: 'error', code: 15, text: '#VALUE!' };
    if (fn === 'TIME') {
      const [hour, minute, second] = values.map((value) => Math.trunc(value as number));
      if ((hour as number) < 0 || (minute as number) < 0 || (second as number) < 0) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const totalSeconds = (hour as number) * 3600 + (minute as number) * 60 + (second as number);
      return { kind: 'number', value: (totalSeconds % 86_400) / 86_400 };
    }
    if (fn === 'DATE') {
      const [year, month, day] = values.map((value) => Math.trunc(value as number));
      return {
        kind: 'number',
        value: dateSerialFromParts(year as number, month as number, day as number),
      };
    }
    if (fn === 'HOUR' || fn === 'MINUTE' || fn === 'SECOND') {
      const totalSeconds = Math.round(serialTimeFraction(values[0] as number) * 86_400) % 86_400;
      if (fn === 'HOUR') return { kind: 'number', value: Math.floor(totalSeconds / 3600) };
      if (fn === 'MINUTE') return { kind: 'number', value: Math.floor(totalSeconds / 60) % 60 };
      return { kind: 'number', value: totalSeconds % 60 };
    }
    if (fn === 'DAYS') {
      return { kind: 'number', value: (values[0] as number) - (values[1] as number) };
    }
    const date = dateFromSerial(values[0] as number);
    if (date === null) return { kind: 'error', code: 15, text: '#VALUE!' };
    if (fn === 'EDATE' || fn === 'EOMONTH') {
      const monthOffset = Math.trunc(values[1] as number);
      const year = date.getUTCFullYear();
      const month = date.getUTCMonth() + monthOffset;
      const lastDay = new Date(Date.UTC(year, month + 1, 0)).getUTCDate();
      const day = fn === 'EOMONTH' ? lastDay : Math.min(date.getUTCDate(), lastDay);
      return { kind: 'number', value: dateSerialFromParts(year, month + 1, day) };
    }
    if (fn === 'YEAR') return { kind: 'number', value: date.getUTCFullYear() };
    if (fn === 'MONTH') return { kind: 'number', value: date.getUTCMonth() + 1 };
    if (fn === 'DAY') return { kind: 'number', value: date.getUTCDate() };
    if (fn === 'ISOWEEKNUM') return { kind: 'number', value: isoWeekNumber(date) };
    const returnType = values.length === 2 ? Math.trunc(values[1] as number) : 1;
    if (fn === 'WEEKNUM') {
      if (returnType === 21) return { kind: 'number', value: isoWeekNumber(date) };
      const firstDay = weekStartForReturnType(returnType);
      if (firstDay === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      const janFirst = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
      const offset = (janFirst.getUTCDay() - firstDay + 7) % 7;
      return { kind: 'number', value: Math.floor((dayOfYear(date) + offset - 1) / 7) + 1 };
    }
    const weekday = weekdayValue(date, returnType);
    return weekday === null
      ? { kind: 'error', code: 15, text: '#VALUE!' }
      : { kind: 'number', value: weekday };
  };
  const oneDimensionalValues = (
    range: FormulaRangeArg | ParsedA1Range,
    rowOffset: number,
    colOffset: number,
  ): CellValue[] | null => {
    const bounds = formulaRangeArgBounds(range, rowOffset, colOffset);
    if (!bounds) return null;
    if (!validRangeBounds(bounds) || (bounds.width !== 1 && bounds.height !== 1)) return null;
    const vertical = bounds.width === 1;
    const count = vertical ? bounds.height : bounds.width;
    const values: CellValue[] = [];
    for (let i = 0; i < count; i += 1) {
      const row = vertical ? bounds.r0 + i : bounds.r0;
      const col = vertical ? bounds.c0 : bounds.c0 + i;
      values.push(state.data.cells.get(addrKey({ sheet, row, col }))?.value ?? { kind: 'blank' });
    }
    return values;
  };
  const matchExactRange = (
    lookup: CellValue,
    range: FormulaRangeArg,
    matchType: CellValue | null,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const matchTypeValue = matchType === null ? 1 : readNumber(matchType);
    if (matchTypeValue === null) return { kind: 'error', code: 15, text: '#VALUE!' };
    const matchTypeInt = Math.trunc(matchTypeValue);
    if (matchTypeInt === 1 || matchTypeInt === -1) {
      const values = oneDimensionalValues(range, rowOffset, colOffset);
      if (!values) return { kind: 'error', code: 15, text: '#VALUE!' };
      const index = approximateMatchIndex(lookup, values, matchTypeInt);
      return index === null
        ? { kind: 'error', code: 6, text: '#N/A' }
        : { kind: 'number', value: index + 1 };
    }
    if (matchType !== null) {
      const value = readNumber(matchType);
      if (value === null || value !== 0) return { kind: 'error', code: 6, text: '#N/A' };
    }
    const bounds = formulaRangeArgBounds(range, rowOffset, colOffset);
    if (!bounds || !validRangeBounds(bounds) || (bounds.width !== 1 && bounds.height !== 1)) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    let index = 1;
    for (let r = bounds.r0; r <= bounds.r1; r += 1) {
      for (let c = bounds.c0; c <= bounds.c1; c += 1) {
        const value = state.data.cells.get(addrKey({ sheet, row: r, col: c }))?.value ?? {
          kind: 'blank' as const,
        };
        if (exactMatchValues(lookup, value)) return { kind: 'number', value: index };
        index += 1;
      }
    }
    return { kind: 'error', code: 6, text: '#N/A' };
  };
  const xmatchRange = (
    lookup: CellValue,
    range: FormulaRangeArg,
    matchMode: CellValue | null,
    searchMode: CellValue | null,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const matchModeValue = matchMode === null ? 0 : readNumber(matchMode);
    const searchModeValue = searchMode === null ? 1 : readNumber(searchMode);
    if (matchModeValue === null || searchModeValue === null) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const matchModeInt = Math.trunc(matchModeValue);
    const searchModeInt = Math.trunc(searchModeValue);
    if (
      (matchModeInt !== 0 && matchModeInt !== 2 && matchModeInt !== -1 && matchModeInt !== 1) ||
      (searchModeInt !== 1 && searchModeInt !== -1)
    ) {
      return { kind: 'error', code: 6, text: '#N/A' };
    }
    const values = oneDimensionalValues(range, rowOffset, colOffset);
    if (!values) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    if (matchModeInt === -1 || matchModeInt === 1) {
      const index = approximateXmatchIndex(lookup, values, matchModeInt);
      return index === null
        ? { kind: 'error', code: 6, text: '#N/A' }
        : { kind: 'number', value: index + 1 };
    }
    for (
      let i = searchModeInt === -1 ? values.length - 1 : 0;
      i >= 0 && i < values.length;
      i += searchModeInt
    ) {
      const candidate = values[i] as CellValue;
      if (exactMatchValues(lookup, candidate, matchModeInt === 2)) {
        return { kind: 'number', value: i + 1 };
      }
    }
    return { kind: 'error', code: 6, text: '#N/A' };
  };
  const indexRange = (
    range: FormulaRangeArg,
    rowValue: CellValue,
    colValue: CellValue | null,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const rowNumber = readNumber(rowValue);
    const colNumber = colValue === null ? null : readNumber(colValue);
    if (rowNumber === null || (colValue !== null && colNumber === null)) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const bounds = formulaRangeArgBounds(range, rowOffset, colOffset);
    if (!bounds || !validRangeBounds(bounds)) return { kind: 'error', code: 15, text: '#VALUE!' };
    const rowIndex = Math.trunc(rowNumber);
    const colIndex = colNumber === null ? null : Math.trunc(colNumber);
    if (rowIndex < 1 || (colIndex !== null && colIndex < 1)) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    let targetRow: number;
    let targetCol: number;
    if (colIndex === null) {
      if (bounds.width === 1) {
        targetRow = bounds.r0 + rowIndex - 1;
        targetCol = bounds.c0;
      } else if (bounds.height === 1) {
        targetRow = bounds.r0;
        targetCol = bounds.c0 + rowIndex - 1;
      } else {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
    } else {
      targetRow = bounds.r0 + rowIndex - 1;
      targetCol = bounds.c0 + colIndex - 1;
    }
    if (targetRow > bounds.r1 || targetCol > bounds.c1) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    return (
      state.data.cells.get(addrKey({ sheet, row: targetRow, col: targetCol }))?.value ?? {
        kind: 'blank',
      }
    );
  };
  const offsetValue = (
    reference: FormulaRangeArg,
    rowsValue: CellValue,
    colsValue: CellValue,
    heightValue: CellValue | null,
    widthValue: CellValue | null,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const rows = readNumber(rowsValue);
    const cols = readNumber(colsValue);
    const rawHeight = heightValue === null ? null : readNumber(heightValue);
    const rawWidth = widthValue === null ? null : readNumber(widthValue);
    if (
      rows === null ||
      cols === null ||
      (heightValue !== null && rawHeight === null) ||
      (widthValue !== null && rawWidth === null)
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const bounds = formulaRangeArgBounds(reference, rowOffset, colOffset);
    if (!bounds || !validRangeBounds(bounds)) return { kind: 'error', code: 15, text: '#VALUE!' };
    const height = rawHeight === null ? bounds.height : Math.trunc(rawHeight);
    const width = rawWidth === null ? bounds.width : Math.trunc(rawWidth);
    if (height !== 1 || width !== 1) return { kind: 'error', code: 15, text: '#VALUE!' };
    const targetRow = bounds.r0 + Math.trunc(rows);
    const targetCol = bounds.c0 + Math.trunc(cols);
    if (targetRow < 0 || targetRow > 1048575 || targetCol < 0 || targetCol > 16383) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    return (
      state.data.cells.get(addrKey({ sheet, row: targetRow, col: targetCol }))?.value ?? {
        kind: 'blank',
      }
    );
  };
  const indirectValue = (
    refTextValue: CellValue,
    a1Value: CellValue | null,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const refText = textValue(refTextValue);
    const a1 = a1Value === null ? true : readLogical(a1Value);
    if (refText === null || a1 === null) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const ref = a1
      ? parseA1Ref(refText, sheet)
      : parseR1C1Ref(refText, sheet, anchorRow + rowOffset, anchorCol + colOffset);
    if (!ref) return { kind: 'error', code: 15, text: '#VALUE!' };
    const row = ref.row;
    const col = ref.col;
    if (row < 0 || row > 1048575 || col < 0 || col > 16383) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    return (
      state.data.cells.get(addrKey({ sheet, row, col }))?.value ?? {
        kind: 'blank',
      }
    );
  };
  const tableLookup = (
    fn: 'VLOOKUP' | 'HLOOKUP',
    lookup: CellValue,
    range: FormulaRangeArg,
    indexValue: CellValue,
    rangeLookup: CellValue,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    if (!isExactLookupMode(rangeLookup) && !isApproximateLookupMode(rangeLookup)) {
      return { kind: 'error', code: 6, text: '#N/A' };
    }
    const indexNumber = readNumber(indexValue);
    if (indexNumber === null) return { kind: 'error', code: 15, text: '#VALUE!' };
    const index = Math.trunc(indexNumber);
    if (index < 1) return { kind: 'error', code: 15, text: '#VALUE!' };
    const bounds = formulaRangeArgBounds(range, rowOffset, colOffset);
    if (!bounds || !validRangeBounds(bounds)) return { kind: 'error', code: 15, text: '#VALUE!' };
    const approximate = isApproximateLookupMode(rangeLookup);
    if (fn === 'VLOOKUP') {
      if (index > bounds.width) return { kind: 'error', code: 15, text: '#VALUE!' };
      if (approximate) {
        const values: CellValue[] = [];
        for (let r = bounds.r0; r <= bounds.r1; r += 1) {
          values.push(
            state.data.cells.get(addrKey({ sheet, row: r, col: bounds.c0 }))?.value ?? {
              kind: 'blank',
            },
          );
        }
        const matchIndex = approximateMatchIndex(lookup, values, 1);
        if (matchIndex === null) return { kind: 'error', code: 6, text: '#N/A' };
        return (
          state.data.cells.get(
            addrKey({ sheet, row: bounds.r0 + matchIndex, col: bounds.c0 + index - 1 }),
          )?.value ?? { kind: 'blank' }
        );
      }
      for (let r = bounds.r0; r <= bounds.r1; r += 1) {
        const candidate = state.data.cells.get(addrKey({ sheet, row: r, col: bounds.c0 }))
          ?.value ?? {
          kind: 'blank' as const,
        };
        if (exactMatchValues(lookup, candidate)) {
          return (
            state.data.cells.get(addrKey({ sheet, row: r, col: bounds.c0 + index - 1 }))?.value ?? {
              kind: 'blank',
            }
          );
        }
      }
    } else {
      if (index > bounds.height) return { kind: 'error', code: 15, text: '#VALUE!' };
      if (approximate) {
        const values: CellValue[] = [];
        for (let c = bounds.c0; c <= bounds.c1; c += 1) {
          values.push(
            state.data.cells.get(addrKey({ sheet, row: bounds.r0, col: c }))?.value ?? {
              kind: 'blank',
            },
          );
        }
        const matchIndex = approximateMatchIndex(lookup, values, 1);
        if (matchIndex === null) return { kind: 'error', code: 6, text: '#N/A' };
        return (
          state.data.cells.get(
            addrKey({ sheet, row: bounds.r0 + index - 1, col: bounds.c0 + matchIndex }),
          )?.value ?? { kind: 'blank' }
        );
      }
      for (let c = bounds.c0; c <= bounds.c1; c += 1) {
        const candidate = state.data.cells.get(addrKey({ sheet, row: bounds.r0, col: c }))
          ?.value ?? {
          kind: 'blank' as const,
        };
        if (exactMatchValues(lookup, candidate)) {
          return (
            state.data.cells.get(addrKey({ sheet, row: bounds.r0 + index - 1, col: c }))?.value ?? {
              kind: 'blank',
            }
          );
        }
      }
    }
    return { kind: 'error', code: 6, text: '#N/A' };
  };
  const xlookupRange = (
    lookup: CellValue,
    lookupRange: FormulaRangeArg,
    returnRange: FormulaRangeArg,
    ifNotFound: CellValue | null,
    matchMode: CellValue | null,
    searchMode: CellValue | null,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const matchModeValue = matchMode === null ? 0 : readNumber(matchMode);
    const searchModeValue = searchMode === null ? 1 : readNumber(searchMode);
    if (matchModeValue === null || searchModeValue === null) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const matchModeInt = Math.trunc(matchModeValue);
    const searchModeInt = Math.trunc(searchModeValue);
    if (
      (matchModeInt !== 0 && matchModeInt !== 2 && matchModeInt !== -1 && matchModeInt !== 1) ||
      (searchModeInt !== 1 && searchModeInt !== -1)
    ) {
      return { kind: 'error', code: 6, text: '#N/A' };
    }
    const lookupBounds = formulaRangeArgBounds(lookupRange, rowOffset, colOffset);
    const returnBounds = formulaRangeArgBounds(returnRange, rowOffset, colOffset);
    if (
      !lookupBounds ||
      !returnBounds ||
      !validRangeBounds(lookupBounds) ||
      !validRangeBounds(returnBounds) ||
      (lookupBounds.width !== 1 && lookupBounds.height !== 1)
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const vertical = lookupBounds.width === 1;
    const count = vertical ? lookupBounds.height : lookupBounds.width;
    if ((vertical && returnBounds.height < count) || (!vertical && returnBounds.width < count)) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const returnAt = (index: number): CellValue =>
      state.data.cells.get(
        addrKey({
          sheet,
          row: vertical ? returnBounds.r0 + index : returnBounds.r0,
          col: vertical ? returnBounds.c0 : returnBounds.c0 + index,
        }),
      )?.value ?? { kind: 'blank' };
    if (matchModeInt === -1 || matchModeInt === 1) {
      const values: CellValue[] = [];
      for (let i = 0; i < count; i += 1) {
        const row = vertical ? lookupBounds.r0 + i : lookupBounds.r0;
        const col = vertical ? lookupBounds.c0 : lookupBounds.c0 + i;
        values.push(state.data.cells.get(addrKey({ sheet, row, col }))?.value ?? { kind: 'blank' });
      }
      const matchIndex = approximateXmatchIndex(lookup, values, matchModeInt);
      return matchIndex === null
        ? (ifNotFound ?? { kind: 'error', code: 6, text: '#N/A' })
        : returnAt(matchIndex);
    }
    for (let i = searchModeInt === -1 ? count - 1 : 0; i >= 0 && i < count; i += searchModeInt) {
      const lookupRow = vertical ? lookupBounds.r0 + i : lookupBounds.r0;
      const lookupCol = vertical ? lookupBounds.c0 : lookupBounds.c0 + i;
      const candidate = state.data.cells.get(addrKey({ sheet, row: lookupRow, col: lookupCol }))
        ?.value ?? {
        kind: 'blank' as const,
      };
      if (exactMatchValues(lookup, candidate, matchModeInt === 2)) {
        return returnAt(i);
      }
    }
    return ifNotFound ?? { kind: 'error', code: 6, text: '#N/A' };
  };
  const vectorLookup = (
    lookup: CellValue,
    lookupRange: FormulaRangeArg,
    resultRange: FormulaRangeArg | undefined,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const lookupValues = oneDimensionalValues(lookupRange, rowOffset, colOffset);
    if (!lookupValues) return { kind: 'error', code: 15, text: '#VALUE!' };
    const resultValues = resultRange
      ? oneDimensionalValues(resultRange, rowOffset, colOffset)
      : lookupValues;
    if (!resultValues || resultValues.length !== lookupValues.length) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const matchIndex = approximateMatchIndex(lookup, lookupValues, 1);
    return matchIndex === null
      ? { kind: 'error', code: 6, text: '#N/A' }
      : (resultValues[matchIndex] as CellValue);
  };
  const cellInfo = (
    infoType: CellValue,
    ref: FormulaRangeArg | undefined,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const type = textValue(infoType)?.trim().toLowerCase();
    if (!type) return { kind: 'error', code: 15, text: '#VALUE!' };
    const position = ref ? singleCellRefPosition(ref, rowOffset, colOffset) : null;
    if (ref && !position) return { kind: 'error', code: 15, text: '#VALUE!' };
    const [row, col] = position
      ? [position.row, position.col]
      : [anchorRow + rowOffset, anchorCol + colOffset];
    if (row < 0 || row > 1048575 || col < 0 || col > 16383) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const value = state.data.cells.get(addrKey({ sheet, row, col }))?.value ?? {
      kind: 'blank' as const,
    };
    if (type === 'address') {
      return { kind: 'text', value: `$${colToLetters(col)}$${row + 1}` };
    }
    if (type === 'row') return { kind: 'number', value: row + 1 };
    if (type === 'col') return { kind: 'number', value: col + 1 };
    if (type === 'contents') return value;
    if (type === 'type') {
      return {
        kind: 'text',
        value: value.kind === 'blank' ? 'b' : value.kind === 'text' ? 'l' : 'v',
      };
    }
    return { kind: 'error', code: 15, text: '#VALUE!' };
  };
  const sheetInfo = (
    fn: 'SHEET' | 'SHEETS',
    range: FormulaRangeArg | undefined,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    if (!range) return { kind: 'number', value: sheet + 1 };
    const bounds = formulaRangeArgBounds(range, rowOffset, colOffset);
    if (!bounds) return { kind: 'error', code: 15, text: '#VALUE!' };
    if (!validRangeBounds(bounds)) return { kind: 'error', code: 15, text: '#VALUE!' };
    return { kind: 'number', value: fn === 'SHEET' ? sheet + 1 : 1 };
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
  const readOperand = (
    operand: FormulaOperand,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
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
      const rate = readNumber(readOperand(operand.rate, rowOffset, colOffset));
      if (rate === null || rate === -1) {
        return {
          kind: 'error',
          code: rate === -1 ? 1 : 15,
          text: rate === -1 ? '#DIV/0!' : '#VALUE!',
        };
      }
      let period = 1;
      let result = 0;
      for (const arg of operand.values) {
        if (arg.kind === 'range' || arg.kind === 'dynamic-range') {
          const bounds = formulaRangeArgBounds(arg, rowOffset, colOffset);
          const values = bounds ? numericValuesInBounds(bounds) : null;
          if (values === null) return { kind: 'error', code: 15, text: '#VALUE!' };
          for (const value of values) {
            result += value / (1 + rate) ** period;
            period += 1;
          }
          continue;
        }
        const value = readNumber(readOperand(arg.operand, rowOffset, colOffset));
        if (value === null) return { kind: 'error', code: 15, text: '#VALUE!' };
        result += value / (1 + rate) ** period;
        period += 1;
      }
      return Number.isFinite(result)
        ? { kind: 'number', value: result }
        : { kind: 'error', code: 6, text: '#NUM!' };
    }
    if (operand.kind === 'mirr') {
      const values = numericValuesInFormulaRangeArg(operand.values, rowOffset, colOffset);
      const financeRate = readNumber(readOperand(operand.financeRate, rowOffset, colOffset));
      const reinvestRate = readNumber(readOperand(operand.reinvestRate, rowOffset, colOffset));
      if (values === null || financeRate === null || reinvestRate === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (values.length < 2 || financeRate === -1 || reinvestRate === -1) {
        return {
          kind: 'error',
          code: financeRate === -1 || reinvestRate === -1 ? 1 : 6,
          text: financeRate === -1 || reinvestRate === -1 ? '#DIV/0!' : '#NUM!',
        };
      }
      let presentValueNegative = 0;
      let futureValuePositive = 0;
      for (let index = 0; index < values.length; index += 1) {
        const value = values[index] ?? 0;
        if (value < 0) presentValueNegative += value / (1 + financeRate) ** index;
        else if (value > 0)
          futureValuePositive += value * (1 + reinvestRate) ** (values.length - 1 - index);
      }
      if (presentValueNegative === 0 || futureValuePositive === 0) {
        return { kind: 'error', code: 1, text: '#DIV/0!' };
      }
      const result = (-futureValuePositive / presentValueNegative) ** (1 / (values.length - 1)) - 1;
      return Number.isFinite(result)
        ? { kind: 'number', value: result }
        : { kind: 'error', code: 6, text: '#NUM!' };
    }
    if (operand.kind === 'xnpv') {
      const rate = readNumber(readOperand(operand.rate, rowOffset, colOffset));
      const values = numericValuesInRangeWithShape(operand.values, rowOffset, colOffset);
      const dates = numericValuesInRangeWithShape(operand.dates, rowOffset, colOffset);
      if (rate === null || values === null || dates === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (
        values.width !== dates.width ||
        values.height !== dates.height ||
        values.values.length === 0 ||
        rate === -1
      ) {
        return {
          kind: 'error',
          code: rate === -1 ? 1 : 15,
          text: rate === -1 ? '#DIV/0!' : '#VALUE!',
        };
      }
      const firstDate = dates.values[0] as number;
      let result = 0;
      for (let index = 0; index < values.values.length; index += 1) {
        const value = values.values[index] as number;
        const date = dates.values[index] as number;
        if (date < firstDate) return { kind: 'error', code: 6, text: '#NUM!' };
        result += value / (1 + rate) ** ((date - firstDate) / 365);
      }
      return Number.isFinite(result)
        ? { kind: 'number', value: result }
        : { kind: 'error', code: 6, text: '#NUM!' };
    }
    if (operand.kind === 'irr') {
      const values = numericValuesInFormulaRangeArg(operand.values, rowOffset, colOffset);
      const rawGuess = operand.guess
        ? readNumber(readOperand(operand.guess, rowOffset, colOffset))
        : 0.1;
      if (values === null || rawGuess === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (values.length === 0 || rawGuess <= -1) return { kind: 'error', code: 6, text: '#NUM!' };
      let hasPositive = false;
      let hasNegative = false;
      for (const value of values) {
        if (value > 0) hasPositive = true;
        if (value < 0) hasNegative = true;
      }
      if (!hasPositive || !hasNegative) return { kind: 'error', code: 1, text: '#DIV/0!' };
      const evaluatePeriodicNpv = (rate: number): number => {
        let total = 0;
        for (let index = 0; index < values.length; index += 1) {
          total += (values[index] as number) / (1 + rate) ** index;
        }
        return total;
      };
      let current = rawGuess;
      for (let iteration = 0; iteration < 100; iteration += 1) {
        if (current <= -1) return { kind: 'error', code: 6, text: '#NUM!' };
        const value = evaluatePeriodicNpv(current);
        if (!Number.isFinite(value)) return { kind: 'error', code: 6, text: '#NUM!' };
        if (Math.abs(value) < 1e-7) return { kind: 'number', value: current };
        const step = Math.max(Math.abs(current) * 1e-6, 1e-7);
        const high = evaluatePeriodicNpv(current + step);
        const low = evaluatePeriodicNpv(current - step);
        const derivative = (high - low) / (2 * step);
        if (!Number.isFinite(derivative) || derivative === 0) {
          return { kind: 'error', code: 6, text: '#NUM!' };
        }
        const next = current - value / derivative;
        if (!Number.isFinite(next)) return { kind: 'error', code: 6, text: '#NUM!' };
        if (Math.abs(next - current) < 1e-10) return { kind: 'number', value: next };
        current = next;
      }
      return { kind: 'error', code: 6, text: '#NUM!' };
    }
    if (operand.kind === 'xirr') {
      const values = numericValuesInRangeWithShape(operand.values, rowOffset, colOffset);
      const dates = numericValuesInRangeWithShape(operand.dates, rowOffset, colOffset);
      const rawGuess = operand.guess
        ? readNumber(readOperand(operand.guess, rowOffset, colOffset))
        : 0.1;
      if (values === null || dates === null || rawGuess === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (
        values.width !== dates.width ||
        values.height !== dates.height ||
        values.values.length === 0 ||
        rawGuess <= -1
      ) {
        return {
          kind: 'error',
          code: rawGuess <= -1 ? 6 : 15,
          text: rawGuess <= -1 ? '#NUM!' : '#VALUE!',
        };
      }
      const firstDate = dates.values[0] as number;
      let hasPositive = false;
      let hasNegative = false;
      for (let index = 0; index < values.values.length; index += 1) {
        const value = values.values[index] as number;
        const date = dates.values[index] as number;
        if (date < firstDate) return { kind: 'error', code: 6, text: '#NUM!' };
        if (value > 0) hasPositive = true;
        if (value < 0) hasNegative = true;
      }
      if (!hasPositive || !hasNegative) return { kind: 'error', code: 1, text: '#DIV/0!' };
      const evaluateXnpv = (rate: number): number => {
        let total = 0;
        for (let index = 0; index < values.values.length; index += 1) {
          total +=
            (values.values[index] as number) /
            (1 + rate) ** (((dates.values[index] as number) - firstDate) / 365);
        }
        return total;
      };
      let current = rawGuess;
      for (let iteration = 0; iteration < 100; iteration += 1) {
        if (current <= -1) return { kind: 'error', code: 6, text: '#NUM!' };
        const value = evaluateXnpv(current);
        if (!Number.isFinite(value)) return { kind: 'error', code: 6, text: '#NUM!' };
        if (Math.abs(value) < 1e-7) return { kind: 'number', value: current };
        const step = Math.max(Math.abs(current) * 1e-6, 1e-7);
        const high = evaluateXnpv(current + step);
        const low = evaluateXnpv(current - step);
        const derivative = (high - low) / (2 * step);
        if (!Number.isFinite(derivative) || derivative === 0) {
          return { kind: 'error', code: 6, text: '#NUM!' };
        }
        const next = current - value / derivative;
        if (!Number.isFinite(next)) return { kind: 'error', code: 6, text: '#NUM!' };
        if (Math.abs(next - current) < 1e-10) return { kind: 'number', value: next };
        current = next;
      }
      return { kind: 'error', code: 6, text: '#NUM!' };
    }
    if (operand.kind === 'fv-schedule') {
      const principal = readNumber(readOperand(operand.principal, rowOffset, colOffset));
      const schedule = numericValuesInFormulaRangeArg(operand.schedule, rowOffset, colOffset);
      if (principal === null || schedule === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      let value = principal;
      for (const rate of schedule) value *= 1 + rate;
      return Number.isFinite(value)
        ? { kind: 'number', value }
        : { kind: 'error', code: 6, text: '#NUM!' };
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
  };
  return readOperand;
}
