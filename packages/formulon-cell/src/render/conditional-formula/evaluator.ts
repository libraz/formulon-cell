import { formatNumber } from '../../commands/format.js';
import { addrKey } from '../../engine/address.js';
import type { CellValue } from '../../engine/types.js';
import type { ConditionalRule, State } from '../../store/store.js';
import {
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
} from './parser.js';
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

function compareValues(
  left: CellValue,
  op: '>' | '<' | '>=' | '<=' | '=' | '<>',
  right: CellValue,
): boolean {
  if (left.kind === 'number' && right.kind === 'number') {
    switch (op) {
      case '>':
        return left.value > right.value;
      case '<':
        return left.value < right.value;
      case '>=':
        return left.value >= right.value;
      case '<=':
        return left.value <= right.value;
      case '=':
        return left.value === right.value;
      case '<>':
        return left.value !== right.value;
    }
  }
  if (left.kind === 'error' && right.kind === 'error') {
    return op === '=' ? left.text === right.text : op === '<>' ? left.text !== right.text : false;
  }
  const leftText =
    left.kind === 'text'
      ? left.value
      : left.kind === 'bool'
        ? String(left.value).toUpperCase()
        : null;
  const rightText =
    right.kind === 'text'
      ? right.value
      : right.kind === 'bool'
        ? String(right.value).toUpperCase()
        : null;
  if (leftText === null || rightText === null) return false;
  const leftComparable = leftText.toLocaleLowerCase();
  const rightComparable = rightText.toLocaleLowerCase();
  return op === '='
    ? leftComparable === rightComparable
    : op === '<>'
      ? leftComparable !== rightComparable
      : false;
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

function countIfWildcardPattern(criteria: string): RegExp | null {
  let pattern = '^';
  let hasWildcard = false;
  for (let i = 0; i < criteria.length; i += 1) {
    const ch = criteria[i] ?? '';
    if (ch === '~') {
      const next = criteria[i + 1];
      if (next === '*' || next === '?' || next === '~') {
        pattern += escapeRegExp(next);
        hasWildcard = true;
        i += 1;
      } else {
        pattern += escapeRegExp(ch);
      }
      continue;
    }
    if (ch === '*') {
      pattern += '.*';
      hasWildcard = true;
      continue;
    }
    if (ch === '?') {
      pattern += '.';
      hasWildcard = true;
      continue;
    }
    pattern += escapeRegExp(ch);
  }
  return hasWildcard ? new RegExp(`${pattern}$`, 'iu') : null;
}

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
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
  const matchesCountIfCriteria = (value: CellValue, criteria: CellValue): boolean => {
    if (criteria.kind === 'text') {
      const raw = criteria.value.trim();
      const m = raw.match(/^(>=|<=|<>|>|<|=)?\s*(.*)$/);
      const op = (m?.[1] ?? '=') as '>' | '<' | '>=' | '<=' | '=' | '<>';
      const rhs = m?.[2] ?? raw;
      if (FORMULA_NUMBER_LITERAL.test(rhs)) {
        return value.kind === 'number'
          ? compareValues(value, op, { kind: 'number', value: Number(rhs) })
          : false;
      }
      if (/^true$/i.test(rhs) || /^false$/i.test(rhs)) {
        return compareValues(value, op, { kind: 'bool', value: /^true$/i.test(rhs) });
      }
      if (rhs === '') {
        const blankLike = value.kind === 'blank' || (value.kind === 'text' && value.value === '');
        return op === '=' ? blankLike : op === '<>' ? !blankLike : false;
      }
      const leftText =
        value.kind === 'text'
          ? value.value
          : value.kind === 'bool'
            ? String(value.value).toUpperCase()
            : value.kind === 'error'
              ? value.text
              : null;
      if (leftText === null) return false;
      const wildcard = op === '=' || op === '<>' ? countIfWildcardPattern(rhs) : null;
      if (wildcard) {
        const matched = wildcard.test(leftText);
        return op === '<>' ? !matched : matched;
      }
      const left = leftText.toLocaleLowerCase();
      const right = rhs.toLocaleLowerCase();
      switch (op) {
        case '=':
          return left === right;
        case '<>':
          return left !== right;
        case '>':
          return left > right;
        case '<':
          return left < right;
        case '>=':
          return left >= right;
        case '<=':
          return left <= right;
      }
    }
    if (criteria.kind === 'number') {
      return value.kind === 'number' && value.value === criteria.value;
    }
    if (criteria.kind === 'bool') {
      return value.kind === 'bool' && value.value === criteria.value;
    }
    if (criteria.kind === 'blank') {
      return value.kind === 'blank' || (value.kind === 'text' && value.value === '');
    }
    return value.kind === 'error' && value.text === criteria.text;
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
  const textValue = (value: CellValue): string | null => {
    if (value.kind === 'text') return value.value;
    if (value.kind === 'number') return String(value.value);
    if (value.kind === 'bool') return value.value ? 'TRUE' : 'FALSE';
    if (value.kind === 'blank') return '';
    return null;
  };
  const concatTextValue = (value: CellValue): string =>
    value.kind === 'error' ? value.text : (textValue(value) ?? '');
  const booleanValue = (value: CellValue): boolean | null => {
    if (value.kind === 'bool') return value.value;
    if (value.kind === 'number' && Number.isFinite(value.value)) return value.value !== 0;
    return null;
  };
  const searchText = (
    fn: 'SEARCH' | 'FIND',
    needleValue: CellValue,
    haystackValue: CellValue,
    startValue: CellValue | null,
  ): CellValue => {
    const needle = textValue(needleValue);
    const haystack = textValue(haystackValue);
    if (needle === null || haystack === null) return { kind: 'error', code: 15, text: '#VALUE!' };
    let start = 0;
    if (startValue !== null) {
      if (
        startValue.kind !== 'number' ||
        !Number.isFinite(startValue.value) ||
        startValue.value < 1
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      start = Math.floor(startValue.value) - 1;
    }
    if (start > haystack.length) return { kind: 'error', code: 15, text: '#VALUE!' };
    const searchNeedle = fn === 'SEARCH' ? needle.toLocaleLowerCase() : needle;
    const searchHaystack = fn === 'SEARCH' ? haystack.toLocaleLowerCase() : haystack;
    const literalNeedle = fn === 'SEARCH' ? searchLiteralPattern(searchNeedle) : null;
    if (literalNeedle !== null) {
      const index = searchHaystack.indexOf(literalNeedle, start);
      return index >= 0
        ? { kind: 'number', value: index + 1 }
        : { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const wildcard = fn === 'SEARCH' ? searchWildcardPattern(searchNeedle) : null;
    if (wildcard) {
      for (let index = start; index <= searchHaystack.length; index += 1) {
        if (wildcard.test(searchHaystack.slice(index))) return { kind: 'number', value: index + 1 };
      }
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const index = searchHaystack.indexOf(searchNeedle, start);
    return index >= 0
      ? { kind: 'number', value: index + 1 }
      : { kind: 'error', code: 15, text: '#VALUE!' };
  };
  const searchLiteralPattern = (criteria: string): string | null => {
    let out = '';
    let hasEscape = false;
    for (let i = 0; i < criteria.length; i += 1) {
      const ch = criteria[i] ?? '';
      if (ch === '~') {
        const next = criteria[i + 1];
        if (next === '*' || next === '?' || next === '~') {
          out += next;
          hasEscape = true;
          i += 1;
          continue;
        }
      }
      if (ch === '*' || ch === '?') return null;
      out += ch;
    }
    return hasEscape ? out : null;
  };
  const searchWildcardPattern = (criteria: string): RegExp | null => {
    let pattern = '^';
    let hasWildcard = false;
    for (let i = 0; i < criteria.length; i += 1) {
      const ch = criteria[i] ?? '';
      if (ch === '~') {
        const next = criteria[i + 1];
        if (next === '*' || next === '?' || next === '~') {
          pattern += escapeRegExp(next);
          i += 1;
        } else {
          pattern += escapeRegExp(ch);
        }
        continue;
      }
      if (ch === '*') {
        pattern += '.*';
        hasWildcard = true;
        continue;
      }
      if (ch === '?') {
        pattern += '.';
        hasWildcard = true;
        continue;
      }
      pattern += escapeRegExp(ch);
    }
    return hasWildcard ? new RegExp(pattern, 'u') : null;
  };
  const nonNegativeInteger = (value: CellValue): number | null => {
    if (value.kind !== 'number' || !Number.isFinite(value.value) || value.value < 0) return null;
    return Math.floor(value.value);
  };
  const positiveInteger = (value: CellValue): number | null => {
    if (value.kind !== 'number' || !Number.isFinite(value.value) || value.value < 1) return null;
    return Math.floor(value.value);
  };
  const sliceText = (
    fn: 'LEFT' | 'RIGHT' | 'MID',
    textValueCell: CellValue,
    startValue: CellValue | null,
    countValue: CellValue,
  ): CellValue => {
    const text = textValue(textValueCell);
    const count = nonNegativeInteger(countValue);
    if (text === null || count === null) return { kind: 'error', code: 15, text: '#VALUE!' };
    if (fn === 'LEFT') return { kind: 'text', value: text.slice(0, count) };
    if (fn === 'RIGHT') return { kind: 'text', value: count === 0 ? '' : text.slice(-count) };
    if (startValue === null) return { kind: 'error', code: 15, text: '#VALUE!' };
    const start = positiveInteger(startValue);
    if (start === null) return { kind: 'error', code: 15, text: '#VALUE!' };
    return { kind: 'text', value: text.slice(start - 1, start - 1 + count) };
  };
  const transformText = (
    fn: 'LOWER' | 'UPPER' | 'TRIM' | 'CLEAN' | 'PROPER' | 'ENCODEURL',
    textValueCell: CellValue,
  ): CellValue => {
    const text = textValue(textValueCell);
    if (text === null) return { kind: 'error', code: 15, text: '#VALUE!' };
    if (fn === 'LOWER') return { kind: 'text', value: text.toLocaleLowerCase() };
    if (fn === 'UPPER') return { kind: 'text', value: text.toLocaleUpperCase() };
    if (fn === 'TRIM') return { kind: 'text', value: text.trim().replace(/ +/g, ' ') };
    if (fn === 'ENCODEURL') return { kind: 'text', value: encodeURIComponent(text) };
    if (fn === 'CLEAN') {
      return {
        kind: 'text',
        value: [...text].filter((char) => char.charCodeAt(0) > 31).join(''),
      };
    }
    return {
      kind: 'text',
      value: text
        .toLocaleLowerCase()
        .replace(
          /(^|[^A-Za-z0-9])([A-Za-z])/g,
          (_match, prefix: string, letter: string) => `${prefix}${letter.toLocaleUpperCase()}`,
        ),
    };
  };
  const substituteText = (
    textValueCell: CellValue,
    oldTextValue: CellValue,
    newTextValue: CellValue,
    instanceValue: CellValue | null,
  ): CellValue => {
    const text = textValue(textValueCell);
    const oldText = textValue(oldTextValue);
    const newText = textValue(newTextValue);
    if (text === null || oldText === null || newText === null) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    if (oldText === '') return { kind: 'text', value: text };
    if (instanceValue === null) return { kind: 'text', value: text.split(oldText).join(newText) };
    const instance = positiveInteger(instanceValue);
    if (instance === null) return { kind: 'error', code: 15, text: '#VALUE!' };
    let seen = 0;
    let offset = 0;
    for (;;) {
      const index = text.indexOf(oldText, offset);
      if (index < 0) return { kind: 'text', value: text };
      seen += 1;
      if (seen === instance) {
        return {
          kind: 'text',
          value: `${text.slice(0, index)}${newText}${text.slice(index + oldText.length)}`,
        };
      }
      offset = index + oldText.length;
    }
  };
  const replaceText = (
    textValueCell: CellValue,
    startValue: CellValue,
    countValue: CellValue,
    newTextValue: CellValue,
  ): CellValue => {
    const text = textValue(textValueCell);
    const newText = textValue(newTextValue);
    const start = positiveInteger(startValue);
    const count = nonNegativeInteger(countValue);
    if (text === null || newText === null || start === null || count === null) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const index = start - 1;
    return { kind: 'text', value: `${text.slice(0, index)}${newText}${text.slice(index + count)}` };
  };
  const repeatText = (textValueCell: CellValue, countValue: CellValue): CellValue => {
    const text = textValue(textValueCell);
    const count = nonNegativeInteger(countValue);
    if (text === null || count === null || text.length * count > 32767) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    return { kind: 'text', value: text.repeat(count) };
  };
  const beforeAfterText = (
    fn: 'TEXTBEFORE' | 'TEXTAFTER',
    textValueCell: CellValue,
    delimiterValue: CellValue,
    instanceValue: CellValue | null,
    matchModeValue: CellValue | null,
    matchEndValue: CellValue | null,
    ifNotFoundValue: CellValue | null,
  ): CellValue => {
    const text = textValue(textValueCell);
    const delimiter = textValue(delimiterValue);
    if (text === null || delimiter === null || delimiter === '') {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const instance = instanceValue === null ? 1 : readNumber(instanceValue);
    const matchMode = matchModeValue === null ? 0 : readNumber(matchModeValue);
    const matchEnd = matchEndValue === null ? 0 : readNumber(matchEndValue);
    if (
      instance === null ||
      matchMode === null ||
      matchEnd === null ||
      Math.trunc(instance) === 0
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const nth = Math.trunc(instance);
    const ignoreCase = Math.trunc(matchMode) === 1;
    if (Math.trunc(matchMode) !== 0 && !ignoreCase) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    if (Math.trunc(matchEnd) !== 0 && Math.trunc(matchEnd) !== 1) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const haystack = ignoreCase ? text.toLocaleLowerCase() : text;
    const needle = ignoreCase ? delimiter.toLocaleLowerCase() : delimiter;
    const matches: number[] = [];
    let offset = 0;
    for (;;) {
      const index = haystack.indexOf(needle, offset);
      if (index < 0) break;
      matches.push(index);
      offset = index + needle.length;
    }
    if (Math.trunc(matchEnd) === 1) {
      if (fn === 'TEXTBEFORE' && nth > 0) matches.push(text.length);
      if (fn === 'TEXTAFTER' && nth < 0) matches.unshift(-delimiter.length);
    }
    const index = nth > 0 ? matches[nth - 1] : matches[matches.length + nth];
    if (index === undefined) {
      return ifNotFoundValue ?? { kind: 'error', code: 6, text: '#N/A' };
    }
    return {
      kind: 'text',
      value:
        fn === 'TEXTBEFORE'
          ? text.slice(0, index)
          : text.slice(Math.max(0, index + delimiter.length)),
    };
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
  const exactText = (leftValue: CellValue, rightValue: CellValue): CellValue => {
    const left = textValue(leftValue);
    const right = textValue(rightValue);
    if (left === null || right === null) return { kind: 'error', code: 15, text: '#VALUE!' };
    return { kind: 'bool', value: left === right };
  };
  const formatText = (valueCell: CellValue, patternCell: CellValue): CellValue => {
    const value = readNumber(valueCell);
    const pattern = textValue(patternCell);
    if (value === null || pattern === null || pattern === '') {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    return { kind: 'text', value: formatNumber(value, { kind: 'custom', pattern }) };
  };
  const fixedFormatText = (
    fn: 'DOLLAR' | 'FIXED',
    valueCell: CellValue,
    decimalsCell: CellValue | null,
    noCommasCell: CellValue | null,
  ): CellValue => {
    const value = readNumber(valueCell);
    const decimalsValue = decimalsCell === null ? 2 : readNumber(decimalsCell);
    const noCommas = noCommasCell === null ? false : booleanValue(noCommasCell);
    if (value === null || decimalsValue === null || noCommas === null) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const decimals = Math.trunc(decimalsValue);
    const visibleDecimals = Math.max(0, decimals);
    const roundedValue =
      decimals >= 0 ? value : Math.round(value / 10 ** -decimals) * 10 ** -decimals;
    return {
      kind: 'text',
      value: formatNumber(
        roundedValue,
        fn === 'DOLLAR'
          ? { kind: 'currency', decimals: visibleDecimals, symbol: '$' }
          : { kind: 'fixed', decimals: visibleDecimals, thousands: !noCommas },
      ),
    };
  };
  const parseNumberText = (
    text: string,
    decimalSeparator = '.',
    groupSeparator = ',',
  ): CellValue => {
    if (
      decimalSeparator.length !== 1 ||
      groupSeparator.length !== 1 ||
      decimalSeparator === groupSeparator
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const isPercent = text.endsWith('%');
    const body = isPercent ? text.slice(0, -1) : text;
    const decimal = escapeRegExp(decimalSeparator);
    const group = escapeRegExp(groupSeparator);
    const pattern = new RegExp(
      `^[+-]?(?:(?:\\d{1,3}(?:${group}\\d{3})+|\\d+)(?:${decimal}\\d*)?|${decimal}\\d+)(?:[eE][+-]?\\d+)?$`,
      'u',
    );
    if (!pattern.test(body)) return { kind: 'error', code: 15, text: '#VALUE!' };
    const normalized = body.split(groupSeparator).join('').replace(decimalSeparator, '.');
    const number = Number(normalized);
    if (!Number.isFinite(number)) return { kind: 'error', code: 15, text: '#VALUE!' };
    return { kind: 'number', value: isPercent ? number / 100 : number };
  };
  const valueText = (value: CellValue): CellValue => {
    if (value.kind === 'number') return value;
    const text = value.kind === 'text' ? value.value.trim() : value.kind === 'blank' ? '' : null;
    if (text === null || text === '') return { kind: 'error', code: 15, text: '#VALUE!' };
    if (!FORMULA_VALUE_NUMBER_LITERAL.test(text)) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    return parseNumberText(text);
  };
  const numberValueText = (
    value: CellValue,
    decimalSeparatorValue: CellValue | null,
    groupSeparatorValue: CellValue | null,
  ): CellValue => {
    const text = value.kind === 'number' ? String(value.value) : (textValue(value)?.trim() ?? null);
    const decimalSeparator =
      decimalSeparatorValue === null ? '.' : textValue(decimalSeparatorValue);
    const groupSeparator = groupSeparatorValue === null ? ',' : textValue(groupSeparatorValue);
    if (text === null || text === '' || !decimalSeparator || !groupSeparator) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    return parseNumberText(text, decimalSeparator, groupSeparator);
  };
  const valueToText = (value: CellValue, formatValue: CellValue | null): CellValue => {
    const format = formatValue === null ? 0 : readNumber(formatValue);
    if (format === null) return { kind: 'error', code: 15, text: '#VALUE!' };
    const mode = Math.trunc(format);
    if (mode !== 0 && mode !== 1) return { kind: 'error', code: 15, text: '#VALUE!' };
    if (value.kind === 'error') return { kind: 'text', value: value.text };
    if (value.kind === 'blank') return { kind: 'text', value: '' };
    const text = textValue(value);
    if (text === null) return { kind: 'error', code: 15, text: '#VALUE!' };
    return {
      kind: 'text',
      value: mode === 1 && value.kind === 'text' ? `"${text.replace(/"/g, '""')}"` : text,
    };
  };
  const coerceScalar = (fn: 'N' | 'T', value: CellValue): CellValue => {
    if (value.kind === 'error') return value;
    if (fn === 'N') {
      if (value.kind === 'number') return value;
      if (value.kind === 'bool') return { kind: 'number', value: value.value ? 1 : 0 };
      return { kind: 'number', value: 0 };
    }
    return value.kind === 'text' ? value : { kind: 'text', value: '' };
  };
  const readNumber = (value: CellValue): number | null =>
    value.kind === 'number' && Number.isFinite(value.value) ? value.value : null;
  const readLogical = (value: CellValue): boolean | null => {
    if (value.kind === 'bool') return value.value;
    if (value.kind === 'number' && Number.isFinite(value.value)) return value.value !== 0;
    return null;
  };
  const erf = (value: number): number => {
    const sign = value < 0 ? -1 : 1;
    const x = Math.abs(value);
    const t = 1 / (1 + 0.5 * x);
    let polynomial = 0.17087277;
    polynomial = -0.82215223 + t * polynomial;
    polynomial = 1.48851587 + t * polynomial;
    polynomial = -1.13520398 + t * polynomial;
    polynomial = 0.27886807 + t * polynomial;
    polynomial = -0.18628806 + t * polynomial;
    polynomial = 0.09678418 + t * polynomial;
    polynomial = 0.37409196 + t * polynomial;
    polynomial = 1.00002368 + t * polynomial;
    const tau = t * Math.exp(-x * x - 1.26551223 + t * polynomial);
    return sign * (1 - tau);
  };
  const standardNormalCdf = (z: number): number => 0.5 * (1 + erf(z / Math.SQRT2));
  const standardNormalPdf = (z: number): number => Math.exp(-0.5 * z * z) / Math.sqrt(2 * Math.PI);
  const inverseStandardNormal = (probability: number): number => {
    const pick = (values: number[], index: number): number => values[index] ?? 0;
    const a = [
      -39.69683028665376, 220.9460984245205, -275.9285104469687, 138.357751867269,
      -30.66479806614716, 2.506628277459239,
    ];
    const b = [
      -54.47609879822406, 161.5858368580409, -155.6989798598866, 66.80131188771972,
      -13.28068155288572,
    ];
    const c = [
      -0.007784894002430293, -0.3223964580411365, -2.400758277161838, -2.549732539343734,
      4.374664141464968, 2.938163982698783,
    ];
    const d = [0.007784695709041462, 0.3224671290700398, 2.445134137142996, 3.754408661907416];
    const low = 0.02425;
    const high = 1 - low;
    if (probability < low) {
      const q = Math.sqrt(-2 * Math.log(probability));
      const numerator =
        ((((pick(c, 0) * q + pick(c, 1)) * q + pick(c, 2)) * q + pick(c, 3)) * q + pick(c, 4)) * q +
        pick(c, 5);
      const denominator =
        (((pick(d, 0) * q + pick(d, 1)) * q + pick(d, 2)) * q + pick(d, 3)) * q + 1;
      return numerator / denominator;
    }
    if (probability > high) {
      const q = Math.sqrt(-2 * Math.log(1 - probability));
      const numerator =
        ((((pick(c, 0) * q + pick(c, 1)) * q + pick(c, 2)) * q + pick(c, 3)) * q + pick(c, 4)) * q +
        pick(c, 5);
      const denominator =
        (((pick(d, 0) * q + pick(d, 1)) * q + pick(d, 2)) * q + pick(d, 3)) * q + 1;
      return -(numerator / denominator);
    }
    const q = probability - 0.5;
    const r = q * q;
    const numerator =
      (((((pick(a, 0) * r + pick(a, 1)) * r + pick(a, 2)) * r + pick(a, 3)) * r + pick(a, 4)) * r +
        pick(a, 5)) *
      q;
    const denominator =
      ((((pick(b, 0) * r + pick(b, 1)) * r + pick(b, 2)) * r + pick(b, 3)) * r + pick(b, 4)) * r +
      1;
    return numerator / denominator;
  };
  const binomialProbability = (successes: number, trials: number, probability: number): number => {
    if (probability === 0) return successes === 0 ? 1 : 0;
    if (probability === 1) return successes === trials ? 1 : 0;
    let coefficient = 1;
    const choose = Math.min(successes, trials - successes);
    for (let i = 1; i <= choose; i += 1) {
      coefficient *= (trials - choose + i) / i;
    }
    return coefficient * probability ** successes * (1 - probability) ** (trials - successes);
  };
  const poissonProbability = (x: number, mean: number): number => {
    if (mean === 0) return x === 0 ? 1 : 0;
    let factorial = 1;
    for (let i = 2; i <= x; i += 1) factorial *= i;
    return (Math.exp(-mean) * mean ** x) / factorial;
  };
  const factorial = (value: number): number => {
    let result = 1;
    for (let i = 2; i <= value; i += 1) result *= i;
    return result;
  };
  const doubleFactorial = (value: number): number => {
    let result = 1;
    for (let i = value; i > 1; i -= 2) result *= i;
    return result;
  };
  const logGamma = (value: number): number => {
    const coefficients = [
      676.5203681218851, -1259.1392167224028, 771.3234287776531, -176.6150291621406,
      12.507343278686905, -0.13857109526572012, 0.000009984369578019572, 0.00000015056327351493116,
    ];
    if (value < 0.5) {
      return Math.log(Math.PI) - Math.log(Math.sin(Math.PI * value)) - logGamma(1 - value);
    }
    const z = value - 1;
    let x = 0.9999999999998099;
    for (let i = 0; i < coefficients.length; i += 1) {
      x += (coefficients[i] as number) / (z + i + 1);
    }
    const t = z + coefficients.length - 0.5;
    return Math.log(Math.sqrt(2 * Math.PI)) + (z + 0.5) * Math.log(t) - t + Math.log(x);
  };
  const gamma = (value: number): number | null => {
    if (value === 0 || (value < 0 && Number.isInteger(value))) return null;
    if (value < 0.5) {
      return Math.PI / (Math.sin(Math.PI * value) * Math.exp(logGamma(1 - value)));
    }
    return Math.exp(logGamma(value));
  };
  const regularizedGammaP = (alpha: number, x: number): number | null => {
    if (x <= 0) return 0;
    const epsilon = 1e-12;
    const maxIterations = 100;
    const tiny = 1e-300;
    const logTerm = alpha * Math.log(x) - x - logGamma(alpha);
    if (x < alpha + 1) {
      let sum = 1 / alpha;
      let term = sum;
      for (let n = 1; n <= maxIterations; n += 1) {
        term *= x / (alpha + n);
        sum += term;
        if (Math.abs(term) < Math.abs(sum) * epsilon) {
          return Math.exp(logTerm) * sum;
        }
      }
      return null;
    }
    let b = x + 1 - alpha;
    let c = 1 / tiny;
    let d = 1 / Math.max(b, tiny);
    let h = d;
    for (let i = 1; i <= maxIterations; i += 1) {
      const an = -i * (i - alpha);
      b += 2;
      d = an * d + b;
      if (Math.abs(d) < tiny) d = tiny;
      c = b + an / c;
      if (Math.abs(c) < tiny) c = tiny;
      d = 1 / d;
      const delta = d * c;
      h *= delta;
      if (Math.abs(delta - 1) < epsilon) {
        return 1 - Math.exp(logTerm) * h;
      }
    }
    return null;
  };
  const inverseRegularizedGammaP = (alpha: number, probability: number): number | null => {
    let low = 0;
    let high = Math.max(1, alpha);
    for (let i = 0; i < 100; i += 1) {
      const value = regularizedGammaP(alpha, high);
      if (value === null) return null;
      if (value >= probability) break;
      high *= 2;
      if (!Number.isFinite(high)) return null;
    }
    for (let i = 0; i < 100; i += 1) {
      const mid = (low + high) / 2;
      const value = regularizedGammaP(alpha, mid);
      if (value === null) return null;
      if (value < probability) low = mid;
      else high = mid;
    }
    return (low + high) / 2;
  };
  const betaContinuedFraction = (x: number, alpha: number, beta: number): number | null => {
    const maxIterations = 100;
    const epsilon = 3e-14;
    const tiny = 1e-300;
    const qab = alpha + beta;
    const qap = alpha + 1;
    const qam = alpha - 1;
    let c = 1;
    let d = 1 - (qab * x) / qap;
    if (Math.abs(d) < tiny) d = tiny;
    d = 1 / d;
    let h = d;
    for (let m = 1; m <= maxIterations; m += 1) {
      const m2 = 2 * m;
      let aa = (m * (beta - m) * x) / ((qam + m2) * (alpha + m2));
      d = 1 + aa * d;
      if (Math.abs(d) < tiny) d = tiny;
      c = 1 + aa / c;
      if (Math.abs(c) < tiny) c = tiny;
      d = 1 / d;
      h *= d * c;
      aa = (-(alpha + m) * (qab + m) * x) / ((alpha + m2) * (qap + m2));
      d = 1 + aa * d;
      if (Math.abs(d) < tiny) d = tiny;
      c = 1 + aa / c;
      if (Math.abs(c) < tiny) c = tiny;
      d = 1 / d;
      const delta = d * c;
      h *= delta;
      if (Math.abs(delta - 1) < epsilon) return h;
    }
    return null;
  };
  const regularizedBeta = (x: number, alpha: number, beta: number): number | null => {
    if (x <= 0) return 0;
    if (x >= 1) return 1;
    const logBt =
      logGamma(alpha + beta) -
      logGamma(alpha) -
      logGamma(beta) +
      alpha * Math.log(x) +
      beta * Math.log(1 - x);
    if (x < (alpha + 1) / (alpha + beta + 2)) {
      const fraction = betaContinuedFraction(x, alpha, beta);
      return fraction === null ? null : (Math.exp(logBt) * fraction) / alpha;
    }
    const fraction = betaContinuedFraction(1 - x, beta, alpha);
    return fraction === null ? null : 1 - (Math.exp(logBt) * fraction) / beta;
  };
  const inverseRegularizedBeta = (
    probability: number,
    alpha: number,
    beta: number,
  ): number | null => {
    let low = 0;
    let high = 1;
    for (let i = 0; i < 100; i += 1) {
      const mid = (low + high) / 2;
      const value = regularizedBeta(mid, alpha, beta);
      if (value === null) return null;
      if (value < probability) low = mid;
      else high = mid;
    }
    return (low + high) / 2;
  };
  const studentTCdf = (x: number, degrees: number): number | null => {
    if (x === 0) return 0.5;
    const betaInput = degrees / (degrees + x * x);
    const betaValue = regularizedBeta(betaInput, degrees / 2, 0.5);
    if (betaValue === null) return null;
    return x > 0 ? 1 - betaValue / 2 : betaValue / 2;
  };
  const studentTPdf = (x: number, degrees: number): number =>
    Math.exp(
      logGamma((degrees + 1) / 2) -
        logGamma(degrees / 2) -
        0.5 * Math.log(degrees * Math.PI) -
        ((degrees + 1) / 2) * Math.log(1 + (x * x) / degrees),
    );
  const inverseStudentTCdf = (probability: number, degrees: number): number | null => {
    let low = -1;
    let high = 1;
    for (let i = 0; i < 100; i += 1) {
      const lowValue = studentTCdf(low, degrees);
      const highValue = studentTCdf(high, degrees);
      if (lowValue === null || highValue === null) return null;
      if (lowValue <= probability && highValue >= probability) break;
      low *= 2;
      high *= 2;
      if (!Number.isFinite(low) || !Number.isFinite(high)) return null;
    }
    for (let i = 0; i < 100; i += 1) {
      const mid = (low + high) / 2;
      const value = studentTCdf(mid, degrees);
      if (value === null) return null;
      if (value < probability) low = mid;
      else high = mid;
    }
    return (low + high) / 2;
  };
  const combination = (n: number, k: number): number => {
    const choose = Math.min(k, n - k);
    let result = 1;
    for (let i = 1; i <= choose; i += 1) result *= (n - choose + i) / i;
    return result;
  };
  const negativeBinomialProbability = (
    failures: number,
    successes: number,
    probability: number,
  ): number =>
    combination(failures + successes - 1, failures) *
    probability ** successes *
    (1 - probability) ** failures;
  const hypergeometricProbability = (
    sampleSuccesses: number,
    sampleSize: number,
    populationSuccesses: number,
    populationSize: number,
  ): number =>
    (combination(populationSuccesses, sampleSuccesses) *
      combination(populationSize - populationSuccesses, sampleSize - sampleSuccesses)) /
    combination(populationSize, sampleSize);
  const baseDigits = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  const engineeringBaseValue = (text: string, base: 2 | 8 | 16): number | null => {
    const normalized = text.trim().toUpperCase();
    if (normalized === '' || normalized.length > 10) return null;
    let unsigned = 0;
    for (const char of normalized) {
      const digit = baseDigits.indexOf(char);
      if (digit < 0 || digit >= base) return null;
      unsigned = unsigned * base + digit;
    }
    const signThreshold = base ** 9;
    const modulus = base ** 10;
    return normalized.length === 10 && unsigned >= signThreshold ? unsigned - modulus : unsigned;
  };
  const engineeringBaseText = (
    value: number,
    base: 2 | 8 | 16,
    places: number | null,
  ): string | null => {
    const negativeLimit = -(base ** 9);
    const positiveLimit = base ** 9 - 1;
    if (value < negativeLimit || value > positiveLimit || (places !== null && places < 0)) {
      return null;
    }
    if (value < 0)
      return Math.trunc(value + base ** 10)
        .toString(base)
        .toUpperCase();
    const text = Math.trunc(value).toString(base).toUpperCase();
    if (places !== null && text.length > places) return null;
    return places === null ? text : text.padStart(places, '0');
  };
  const romanNumerals: [number, string][] = [
    [1000, 'M'],
    [900, 'CM'],
    [500, 'D'],
    [400, 'CD'],
    [100, 'C'],
    [90, 'XC'],
    [50, 'L'],
    [40, 'XL'],
    [10, 'X'],
    [9, 'IX'],
    [5, 'V'],
    [4, 'IV'],
    [1, 'I'],
  ];
  const romanText = (value: number): string => {
    let remaining = value;
    let result = '';
    for (const [amount, symbol] of romanNumerals) {
      while (remaining >= amount) {
        result += symbol;
        remaining -= amount;
      }
    }
    return result;
  };
  const romanValue = (value: string): number | null => {
    const normalized = value.trim().toUpperCase();
    if (normalized === '') return null;
    let index = 0;
    let result = 0;
    while (index < normalized.length) {
      const match = romanNumerals.find(([, symbol]) => normalized.startsWith(symbol, index));
      if (!match) return null;
      result += match[0];
      index += match[1].length;
    }
    return romanText(result) === normalized ? result : null;
  };
  const maxBitValue = 281_474_976_710_655;
  const bitOperand = (value: number): bigint | null => {
    const integer = Math.trunc(value);
    return integer < 0 || integer > maxBitValue ? null : BigInt(integer);
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
  const dateSerialFromParts = (year: number, month: number, day: number): number => {
    const normalizedYear = year >= 0 && year < 1900 ? year + 1900 : year;
    const date = new Date(Date.UTC(normalizedYear, month - 1, day));
    if (normalizedYear >= 0 && normalizedYear < 100) date.setUTCFullYear(normalizedYear);
    return date.getTime() / 86_400_000 + 25569;
  };
  const validatedDateSerial = (year: number, month: number, day: number): number | null => {
    const normalizedYear = year >= 0 && year < 1900 ? year + 1900 : year;
    const date = new Date(Date.UTC(normalizedYear, month - 1, day));
    if (normalizedYear >= 0 && normalizedYear < 100) date.setUTCFullYear(normalizedYear);
    if (
      date.getUTCFullYear() !== normalizedYear ||
      date.getUTCMonth() !== month - 1 ||
      date.getUTCDate() !== day
    ) {
      return null;
    }
    return dateSerialFromParts(year, month, day);
  };
  const dateValueText = (value: CellValue): CellValue => {
    const text = textValue(value)?.trim();
    if (!text) return { kind: 'error', code: 15, text: '#VALUE!' };
    let match = /^(\d{4})-(\d{1,2})-(\d{1,2})$/u.exec(text);
    if (match) {
      const serial = validatedDateSerial(Number(match[1]), Number(match[2]), Number(match[3]));
      return serial === null
        ? { kind: 'error', code: 15, text: '#VALUE!' }
        : { kind: 'number', value: serial };
    }
    match = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/u.exec(text);
    if (match) {
      const rawYear = Number(match[3]);
      const year = rawYear < 100 ? (rawYear < 30 ? rawYear + 2000 : rawYear + 1900) : rawYear;
      const serial = validatedDateSerial(year, Number(match[1]), Number(match[2]));
      return serial === null
        ? { kind: 'error', code: 15, text: '#VALUE!' }
        : { kind: 'number', value: serial };
    }
    return { kind: 'error', code: 15, text: '#VALUE!' };
  };
  const timeValueText = (value: CellValue): CellValue => {
    const text = textValue(value)?.trim();
    if (!text) return { kind: 'error', code: 15, text: '#VALUE!' };
    const match = /^(\d{1,2})(?::(\d{1,2}))(?::(\d{1,2}))?\s*(AM|PM)?$/iu.exec(text);
    if (!match) return { kind: 'error', code: 15, text: '#VALUE!' };
    let hour = Number(match[1]);
    const minute = Number(match[2]);
    const second = match[3] === undefined ? 0 : Number(match[3]);
    const meridiem = match[4]?.toUpperCase();
    if (meridiem) {
      if (hour < 1 || hour > 12) return { kind: 'error', code: 15, text: '#VALUE!' };
      hour = (hour % 12) + (meridiem === 'PM' ? 12 : 0);
    }
    if (hour > 23 || minute > 59 || second > 59) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    return { kind: 'number', value: (hour * 3600 + minute * 60 + second) / 86_400 };
  };
  const todaySerial = (): number => {
    const now = new Date(Date.now());
    return dateSerialFromParts(now.getUTCFullYear(), now.getUTCMonth() + 1, now.getUTCDate());
  };
  const serialTimeFraction = (serial: number): number => ((serial % 1) + 1) % 1;
  const dateFromSerial = (serial: number): Date | null => {
    if (!Number.isFinite(serial)) return null;
    return new Date(Math.trunc(serial) * 86_400_000 - 25569 * 86_400_000);
  };
  const serialDateParts = (serial: number): { year: number; month: number; day: number } | null => {
    const date = dateFromSerial(serial);
    if (date === null) return null;
    return {
      year: date.getUTCFullYear(),
      month: date.getUTCMonth() + 1,
      day: date.getUTCDate(),
    };
  };
  const isLastDayOfFebruary = (year: number, month: number, day: number): boolean =>
    month === 2 && day === new Date(Date.UTC(year, 2, 0)).getUTCDate();
  const nextMonthStart = (year: number, month: number): { year: number; month: number; day: 1 } =>
    month === 12 ? { year: year + 1, month: 1, day: 1 } : { year, month: month + 1, day: 1 };
  const previousMonthParts = (year: number, month: number): { year: number; month: number } =>
    month === 1 ? { year: year - 1, month: 12 } : { year, month: month - 1 };
  const daysInMonth = (year: number, month: number): number =>
    new Date(Date.UTC(year, month, 0)).getUTCDate();
  const clampedDateSerial = (year: number, month: number, day: number): number =>
    dateSerialFromParts(year, month, Math.min(day, daysInMonth(year, month)));
  const defaultWeekendDays = new Set<number>([0, 6]);
  const isWeekendSerial = (serial: number, weekends = defaultWeekendDays): boolean => {
    const date = dateFromSerial(serial);
    if (date === null) return false;
    return weekends.has(date.getUTCDay());
  };
  const weekendDaysFromCode = (code: number): Set<number> | null => {
    const normalized = Math.trunc(code);
    if (normalized >= 1 && normalized <= 7) {
      const first = (normalized + 5) % 7;
      return new Set<number>([first, (first + 1) % 7]);
    }
    if (normalized >= 11 && normalized <= 17) {
      return new Set<number>([normalized - 11]);
    }
    return null;
  };
  const weekendDaysFromValue = (value: CellValue): Set<number> | null => {
    if (value.kind === 'number') return weekendDaysFromCode(value.value);
    const text = textValue(value)?.trim();
    if (!text) return null;
    if (/^[01]{7}$/.test(text)) {
      const days = new Set<number>();
      for (let index = 0; index < text.length; index += 1) {
        if (text[index] === '1') days.add((index + 1) % 7);
      }
      return days.size === 7 ? null : days;
    }
    if (FORMULA_NUMBER_LITERAL.test(text)) return weekendDaysFromCode(Number(text));
    return null;
  };
  const isBusinessDay = (
    serial: number,
    holidays: Set<number>,
    weekends = defaultWeekendDays,
  ): boolean => !isWeekendSerial(serial, weekends) && !holidays.has(Math.trunc(serial));
  const networkDays = (
    start: number,
    end: number,
    holidays = new Set<number>(),
    weekends = defaultWeekendDays,
  ): number => {
    const first = Math.trunc(start);
    const last = Math.trunc(end);
    const direction = first <= last ? 1 : -1;
    let count = 0;
    for (let serial = first; direction > 0 ? serial <= last : serial >= last; serial += direction) {
      if (isBusinessDay(serial, holidays, weekends)) count += direction;
    }
    return count;
  };
  const workday = (
    start: number,
    days: number,
    holidays = new Set<number>(),
    weekends = defaultWeekendDays,
  ): number => {
    let remaining = Math.trunc(days);
    let serial = Math.trunc(start);
    const direction = remaining >= 0 ? 1 : -1;
    while (remaining !== 0) {
      serial += direction;
      if (!isBusinessDay(serial, holidays, weekends)) continue;
      remaining -= direction;
    }
    return serial;
  };
  const days360 = (start: number, end: number, european: boolean): CellValue => {
    const startParts = serialDateParts(start);
    const endParts = serialDateParts(end);
    if (startParts === null || endParts === null) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    let { year: y1, month: m1, day: d1 } = startParts;
    let { year: y2, month: m2, day: d2 } = endParts;
    if (european) {
      if (d1 === 31) d1 = 30;
      if (d2 === 31) d2 = 30;
    } else {
      if (d1 === 31 || isLastDayOfFebruary(y1, m1, d1)) d1 = 30;
      if (isLastDayOfFebruary(y2, m2, d2)) {
        if (d1 < 30) {
          ({ year: y2, month: m2, day: d2 } = nextMonthStart(y2, m2));
        } else {
          d2 = 30;
        }
      } else if (d2 === 31) {
        if (d1 < 30) {
          ({ year: y2, month: m2, day: d2 } = nextMonthStart(y2, m2));
        } else {
          d2 = 30;
        }
      }
    }
    return { kind: 'number', value: (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1) };
  };
  const isLeapYear = (year: number): boolean =>
    (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
  const daysInYear = (year: number): number => (isLeapYear(year) ? 366 : 365);
  const actualActualYearFrac = (start: number, end: number): CellValue => {
    const startDate = dateFromSerial(start);
    const endDate = dateFromSerial(end);
    if (startDate === null || endDate === null) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const first = Math.trunc(start);
    const last = Math.trunc(end);
    if (first === last) return { kind: 'number', value: 0 };
    if (first > last) {
      const value = actualActualYearFrac(end, start);
      return value.kind === 'number' ? { kind: 'number', value: -value.value } : value;
    }
    const startYear = startDate.getUTCFullYear();
    const endYear = endDate.getUTCFullYear();
    if (startYear === endYear) {
      return { kind: 'number', value: (last - first) / daysInYear(startYear) };
    }
    const nextYearStart = dateSerialFromParts(startYear + 1, 1, 1);
    const endYearStart = dateSerialFromParts(endYear, 1, 1);
    let value = (nextYearStart - first) / daysInYear(startYear);
    for (let year = startYear + 1; year < endYear; year += 1) {
      value += 1;
    }
    value += (last - endYearStart) / daysInYear(endYear);
    return { kind: 'number', value };
  };
  const yearFrac = (start: number, end: number, basis: number): CellValue => {
    const normalizedBasis = Math.trunc(basis);
    if (normalizedBasis < 0 || normalizedBasis > 4) {
      return { kind: 'error', code: 6, text: '#NUM!' };
    }
    if (normalizedBasis === 0 || normalizedBasis === 4) {
      const value = days360(start, end, normalizedBasis === 4);
      return value.kind === 'number' ? { kind: 'number', value: value.value / 360 } : value;
    }
    const days = Math.trunc(end) - Math.trunc(start);
    if (normalizedBasis === 1) return actualActualYearFrac(start, end);
    return { kind: 'number', value: days / (normalizedBasis === 2 ? 360 : 365) };
  };
  const datedif = (start: number, end: number, unitValue: CellValue): CellValue => {
    const startParts = serialDateParts(start);
    const endParts = serialDateParts(end);
    const unit = textValue(unitValue)?.trim().toUpperCase();
    if (startParts === null || endParts === null || !unit) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const first = Math.trunc(start);
    const last = Math.trunc(end);
    if (first > last) return { kind: 'error', code: 6, text: '#NUM!' };
    const { year: y1, month: m1, day: d1 } = startParts;
    const { year: y2, month: m2, day: d2 } = endParts;
    const anniversaryThisYear = clampedDateSerial(y2, m1, d1);
    const fullYears = y2 - y1 - (anniversaryThisYear > last ? 1 : 0);
    const fullMonths = (y2 - y1) * 12 + (m2 - m1) - (d2 < d1 ? 1 : 0);
    if (unit === 'D') return { kind: 'number', value: last - first };
    if (unit === 'Y') return { kind: 'number', value: fullYears };
    if (unit === 'M') return { kind: 'number', value: fullMonths };
    if (unit === 'YM') return { kind: 'number', value: ((fullMonths % 12) + 12) % 12 };
    if (unit === 'YD') {
      const anniversary =
        anniversaryThisYear <= last ? anniversaryThisYear : clampedDateSerial(y2 - 1, m1, d1);
      return { kind: 'number', value: last - anniversary };
    }
    if (unit === 'MD') {
      if (d2 >= d1) return { kind: 'number', value: d2 - d1 };
      const previous = previousMonthParts(y2, m2);
      return { kind: 'number', value: last - clampedDateSerial(previous.year, previous.month, d1) };
    }
    return { kind: 'error', code: 6, text: '#NUM!' };
  };
  const dayOfYear = (date: Date): number => {
    const start = Date.UTC(date.getUTCFullYear(), 0, 1);
    return Math.floor((date.getTime() - start) / 86_400_000) + 1;
  };
  const isoWeekNumber = (date: Date): number => {
    const normalized = new Date(
      Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()),
    );
    const day = normalized.getUTCDay() || 7;
    normalized.setUTCDate(normalized.getUTCDate() + 4 - day);
    const yearStart = new Date(Date.UTC(normalized.getUTCFullYear(), 0, 1));
    return Math.ceil(((normalized.getTime() - yearStart.getTime()) / 86_400_000 + 1) / 7);
  };
  const weekStartForReturnType = (returnType: number): number | null => {
    if (returnType === 1 || returnType === 17) return 0;
    if (returnType === 2 || returnType === 11) return 1;
    if (returnType >= 12 && returnType <= 16) return returnType - 10;
    return null;
  };
  const weekdayValue = (date: Date, returnType: number): number | null => {
    const day = date.getUTCDay();
    if (returnType === 1) return day + 1;
    if (returnType === 2) return ((day + 6) % 7) + 1;
    if (returnType === 3) return (day + 6) % 7;
    if (returnType >= 11 && returnType <= 17) {
      const firstDay = returnType === 17 ? 0 : returnType - 10;
      return ((day - firstDay + 7) % 7) + 1;
    }
    return null;
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
  const exactMatchValues = (left: CellValue, right: CellValue, allowWildcard = true): boolean => {
    if (left.kind === 'blank' && right.kind === 'blank') return true;
    if (left.kind === 'number' && right.kind === 'number') return left.value === right.value;
    if (left.kind === 'bool' && right.kind === 'bool') return left.value === right.value;
    if (left.kind === 'text' && right.kind === 'text') {
      const wildcard = allowWildcard ? countIfWildcardPattern(left.value) : null;
      if (wildcard) return wildcard.test(right.value);
      return left.value.toLocaleLowerCase() === right.value.toLocaleLowerCase();
    }
    if (left.kind === 'error' && right.kind === 'error') return left.text === right.text;
    return false;
  };
  const compareApproxValues = (lookup: CellValue, candidate: CellValue): number | null => {
    if (lookup.kind === 'number' && candidate.kind === 'number') {
      if (!Number.isFinite(lookup.value) || !Number.isFinite(candidate.value)) return null;
      return candidate.value === lookup.value ? 0 : candidate.value < lookup.value ? -1 : 1;
    }
    if (lookup.kind === 'text' && candidate.kind === 'text') {
      const left = candidate.value.toLocaleLowerCase();
      const right = lookup.value.toLocaleLowerCase();
      return left === right ? 0 : left < right ? -1 : 1;
    }
    return null;
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
  const approximateMatchIndex = (
    lookup: CellValue,
    values: CellValue[],
    mode: -1 | 1,
  ): number | null => {
    let bestIndex: number | null = null;
    let previous: CellValue | null = null;
    for (let i = 0; i < values.length; i += 1) {
      const candidate = values[i] as CellValue;
      const comparison = compareApproxValues(lookup, candidate);
      if (comparison === null) return null;
      if (previous) {
        const order = compareApproxValues(previous, candidate);
        if (order === null || (mode === 1 ? order < 0 : order > 0)) return null;
      }
      previous = candidate;
      if (mode === 1 ? comparison <= 0 : comparison >= 0) {
        bestIndex = i;
      }
    }
    return bestIndex;
  };
  const approximateXmatchIndex = (
    lookup: CellValue,
    values: CellValue[],
    mode: -1 | 1,
  ): number | null => {
    for (let i = 0; i < values.length; i += 1) {
      const candidate = values[i] as CellValue;
      if (compareApproxValues(lookup, candidate) === null) return null;
      if (i > 0) {
        const previous = values[i - 1] as CellValue;
        const order = compareApproxValues(previous, candidate);
        if (order === null || order < 0) return null;
      }
    }
    let nextSmaller: number | null = null;
    for (let i = 0; i < values.length; i += 1) {
      const candidate = values[i] as CellValue;
      const comparison = compareApproxValues(lookup, candidate) as number;
      if (comparison === 0) return i;
      if (mode === -1 && comparison < 0) nextSmaller = i;
      if (mode === 1 && comparison > 0) return i;
    }
    return mode === -1 ? nextSmaller : null;
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
  const isExactLookupMode = (value: CellValue): boolean =>
    (value.kind === 'bool' && !value.value) || (value.kind === 'number' && value.value === 0);
  const isApproximateLookupMode = (value: CellValue): boolean =>
    (value.kind === 'bool' && value.value) || (value.kind === 'number' && value.value !== 0);
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
