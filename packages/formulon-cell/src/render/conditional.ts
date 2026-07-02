import { addrKey } from '../engine/address.js';
import type { CellValue, Range } from '../engine/types.js';
import type {
  CellFormat,
  ConditionalIconSet,
  ConditionalRule,
  ConditionalScalePoint,
  State,
} from '../store/store.js';

const inRange = (sheet: number, row: number, col: number, r: Range): boolean =>
  r.sheet === sheet && row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1;

/** Per-cell visual outputs derived from the active conditional rules. The
 *  renderer consults this for each painted cell to overlay fills, bars, and
 *  font tweaks. */
export interface ConditionalCellOverlay {
  fill?: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** Width fraction (0..1) for a horizontal data bar drawn behind the text.
   *  When set, `barColor` is also defined. */
  bar?: number;
  /** Zero-axis position (0..1) for signed data bars. Defaults to 0. */
  barAxis?: number;
  /** Direction from the zero-axis. Defaults to right. */
  barDirection?: 'left' | 'right';
  barColor?: string;
  barGradient?: boolean;
  /** Icon-set artwork + slot index. When set, the painter draws a small
   *  glyph in a left gutter inside the cell. `slot` is 0-based and
   *  bounded by the icon family (3 or 5). */
  iconKind?: ConditionalIconSet;
  iconSlot?: number;
  /** False when conditional formatting should hide the underlying cell value. */
  showValue?: boolean;
}

// Single-slot identity cache. zustand replaces conditional.rules /
// data.cells by reference on every mutation, so a triple reference match
// means the previous evaluation is still valid. Pan, scroll, and selection
// changes leave these references untouched and hit the cache.
let cachedRulesRef: State['conditional']['rules'] | null = null;
let cachedCellsRef: State['data']['cells'] | null = null;
let cachedSheet: number | null = null;
let cachedOverlay: Map<string, ConditionalCellOverlay> | null = null;

/** Test hook — drop the cached overlay so the next call recomputes. */
export function _resetConditionalCache(): void {
  cachedRulesRef = null;
  cachedCellsRef = null;
  cachedSheet = null;
  cachedOverlay = null;
}

/** Number of slots per icon family. `arrows5` is the only 5-slot family;
 *  the rest land on 3 slots with thresholds at 0.33 / 0.67. */
export function iconSetSlotCount(set: ConditionalIconSet): 3 | 5 {
  return set === 'arrows5' ||
    set === 'quarters5' ||
    set === 'ratings5' ||
    set === 'bars5' ||
    set === 'boxes5'
    ? 5
    : 3;
}

/** Classify `t` (a 0..1 percentile) into a slot index for the icon family.
 *  Uses the spreadsheet's default thresholds — [0.33, 0.67] for 3-slot families and
 *  [0.20, 0.40, 0.60, 0.80] for 5-slot families. */
export function iconSetSlotFor(set: ConditionalIconSet, t: number): number {
  if (iconSetSlotCount(set) === 5) {
    if (t < 0.2) return 0;
    if (t < 0.4) return 1;
    if (t < 0.6) return 2;
    if (t < 0.8) return 3;
    return 4;
  }
  if (t < 0.33) return 0;
  if (t < 0.67) return 1;
  return 2;
}

/** Pick the cells whose values land in the top-N (or bottom-N) of `values`.
 *  Ties at the threshold all qualify so the result count can exceed `n` when
 *  the input has duplicates — spreadsheet parity. Returns the inclusive cutoff. */
export function topBottomThreshold(
  values: readonly number[],
  mode: 'top' | 'bottom',
  n: number,
  percent: boolean,
): number | null {
  if (values.length === 0 || !Number.isFinite(n) || n <= 0) return null;
  const k = percent
    ? Math.max(1, Math.ceil((values.length * n) / 100))
    : Math.min(values.length, Math.floor(n));
  if (k <= 0) return null;
  const sorted = values.slice().sort((a, b) => (mode === 'top' ? b - a : a - b));
  // The k-th element (1-indexed) is the threshold; ties at the threshold
  // still qualify so `Math.min(k, sorted.length) - 1` is the index.
  const idx = Math.min(k, sorted.length) - 1;
  return sorted[idx] ?? null;
}

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

interface ParsedRef {
  row: number;
  col: number;
  absRow: boolean;
  absCol: boolean;
}

interface ParsedA1Range {
  start: ParsedRef;
  end: ParsedRef;
}

type FormulaAggregateName = 'SUM' | 'AVERAGE' | 'MIN' | 'MAX' | 'COUNT';

type FormulaOperand =
  | { kind: 'ref'; ref: ParsedRef }
  | { kind: 'range-aggregate'; fn: FormulaAggregateName; range: ParsedA1Range }
  | { kind: 'countif'; range: ParsedA1Range; criteria: FormulaOperand }
  | { kind: 'countifs'; pairs: { range: ParsedA1Range; criteria: FormulaOperand }[] }
  | { kind: 'literal'; value: CellValue }
  | {
      kind: 'binary';
      op: '+' | '-' | '*' | '/' | '^';
      left: FormulaOperand;
      right: FormulaOperand;
    };

const MAX_FORMULA_AGGREGATE_CELLS = 10000;
const FORMULA_NUMBER_LITERAL = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;

const lettersToCol = (letters: string): number => {
  let col = 0;
  for (let i = 0; i < letters.length; i += 1) {
    col = col * 26 + (letters.toUpperCase().charCodeAt(i) - 64);
  }
  return col - 1;
};

function sheetNameMatchesIndex(name: string, sheetIndex: number): boolean {
  return name.trim().toLowerCase() === `sheet${sheetIndex + 1}`.toLowerCase();
}

function stripSupportedSheetQualifier(raw: string, sheetIndex: number): string | null {
  const body = raw.trim();
  if (!body.includes('!')) return body;
  if (body.startsWith("'")) {
    let name = '';
    for (let i = 1; i < body.length; i += 1) {
      const ch = body[i];
      if (ch === "'") {
        if (body[i + 1] === "'") {
          name += "'";
          i += 1;
          continue;
        }
        if (body[i + 1] !== '!') return null;
        return sheetNameMatchesIndex(name, sheetIndex) ? body.slice(i + 2).trim() : null;
      }
      name += ch;
    }
    return null;
  }
  const bang = body.indexOf('!');
  const sheetName = body.slice(0, bang);
  if (!/^[A-Za-z_][A-Za-z0-9_. ]*$/.test(sheetName)) return null;
  return sheetNameMatchesIndex(sheetName, sheetIndex) ? body.slice(bang + 1).trim() : null;
}

function parseA1Ref(raw: string, sheetIndex: number): ParsedRef | null {
  const body = stripSupportedSheetQualifier(raw, sheetIndex);
  if (body === null) return null;
  const m = body.match(/^(\$?)([A-Za-z]+)(\$?)(\d+)$/);
  if (!m) return null;
  const col = lettersToCol(m[2] ?? '');
  const row = Number.parseInt(m[4] ?? '', 10) - 1;
  if (row < 0 || col < 0 || row > 1048575 || col > 16383) return null;
  return { row, col, absCol: m[1] === '$', absRow: m[3] === '$' };
}

function parseA1Range(raw: string, sheetIndex: number): ParsedA1Range | null {
  const body = stripSupportedSheetQualifier(raw, sheetIndex);
  if (body === null) return null;
  const parts = body.split(':');
  if (parts.length === 1) {
    const ref = parseA1Ref(parts[0] ?? '', sheetIndex);
    return ref ? { start: ref, end: ref } : null;
  }
  if (parts.length !== 2) return null;
  const start = parseA1Ref(parts[0] ?? '', sheetIndex);
  const end = parseA1Ref(parts[1] ?? '', sheetIndex);
  return start && end ? { start, end } : null;
}

function parseFormulaOperand(raw: string, sheetIndex: number): FormulaOperand | null {
  const body = stripOuterParens(raw.trim());
  const ref = parseA1Ref(body, sheetIndex);
  if (ref) return { kind: 'ref', ref };
  const aggregate = body.match(/^([A-Za-z]+)\s*\((.*)\)$/);
  if (aggregate) {
    const fn = (aggregate[1] ?? '').toUpperCase();
    if (fn === 'SUM' || fn === 'AVERAGE' || fn === 'MIN' || fn === 'MAX' || fn === 'COUNT') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 1) {
        const range = parseA1Range(args[0] ?? '', sheetIndex);
        if (range) return { kind: 'range-aggregate', fn, range };
      }
    }
    if (fn === 'COUNTIF') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2) {
        const range = parseA1Range(args[0] ?? '', sheetIndex);
        const criteria = parseFormulaOperand(args[1] ?? '', sheetIndex);
        if (range && criteria) return { kind: 'countif', range, criteria };
      }
    }
    if (fn === 'COUNTIFS') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args && args.length >= 2 && args.length % 2 === 0) {
        const pairs: { range: ParsedA1Range; criteria: FormulaOperand }[] = [];
        for (let i = 0; i < args.length; i += 2) {
          const range = parseA1Range(args[i] ?? '', sheetIndex);
          const criteria = parseFormulaOperand(args[i + 1] ?? '', sheetIndex);
          if (!range || !criteria) return null;
          pairs.push({ range, criteria });
        }
        return { kind: 'countifs', pairs };
      }
    }
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
  const arithmetic = splitFormulaArithmetic(body);
  if (arithmetic) {
    const left = parseFormulaOperand(arithmetic.left, sheetIndex);
    const right = parseFormulaOperand(arithmetic.right, sheetIndex);
    if (left && right) return { kind: 'binary', op: arithmetic.op, left, right };
  }
  return null;
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
  return op === '=' ? leftText === rightText : op === '<>' ? leftText !== rightText : false;
}

function parseFormulaCellPredicate(
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
  const logical = inner.match(/^([A-Za-z]+)\s*\((.*)\)$/);
  if (logical) {
    const name = (logical[1] ?? '').toUpperCase();
    if (name === 'AND' || name === 'OR' || name === 'NOT') {
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
          return name === 'AND'
            ? parsed.every((predicate) => predicate.test(row, col))
            : parsed.some((predicate) => predicate.test(row, col));
        },
      };
    }
    if (name === 'IF') {
      const args = splitFormulaArgs(logical[2] ?? '');
      if (args?.length !== 3) return null;
      const condition = parseFormulaBooleanExpression(state, rule, args[0] ?? '');
      const whenTrue = parseFormulaBooleanExpression(state, rule, args[1] ?? '');
      const whenFalse = parseFormulaBooleanExpression(state, rule, args[2] ?? '');
      if (!condition || !whenTrue || !whenFalse) return null;
      return {
        test(row, col): boolean {
          return condition.test(row, col) ? whenTrue.test(row, col) : whenFalse.test(row, col);
        },
      };
    }
    if (name === 'ISBLANK' || name === 'ISERROR' || name === 'ISNUMBER' || name === 'ISTEXT') {
      const args = splitFormulaArgs(logical[2] ?? '');
      if (args?.length !== 1) return null;
      const operand = parseFormulaOperand(args[0] ?? '', state.data.sheetIndex);
      if (!operand) return null;
      const readOperand = makeFormulaOperandReader(state, state.data.sheetIndex);
      return {
        test(row, col): boolean {
          const rowOffset = row - rule.range.r0;
          const colOffset = col - rule.range.c0;
          const value = readOperand(operand, rowOffset, colOffset);
          if (name === 'ISBLANK') return value.kind === 'blank';
          if (name === 'ISERROR') return value.kind === 'error';
          if (name === 'ISNUMBER') return value.kind === 'number';
          return value.kind === 'text';
        },
      };
    }
  }
  return parseFormulaComparisonPredicate(state, rule, inner);
}

function stripOuterParens(body: string): string {
  let out = body;
  for (;;) {
    if (!out.startsWith('(') || !out.endsWith(')')) return out;
    const inner = out.slice(1, -1);
    if (splitFormulaArgs(inner) === null) return out;
    out = inner.trim();
  }
}

function splitFormulaArgs(raw: string): string[] | null {
  const args: string[] = [];
  let depth = 0;
  let quote: '"' | "'" | null = null;
  let start = 0;
  for (let i = 0; i < raw.length; i += 1) {
    const ch = raw[i];
    if (quote) {
      if (ch === quote) quote = null;
      continue;
    }
    if (ch === '"' || ch === "'") {
      quote = ch;
      continue;
    }
    if (ch === '(') {
      depth += 1;
      continue;
    }
    if (ch === ')') {
      depth -= 1;
      if (depth < 0) return null;
      continue;
    }
    if (ch === ',' && depth === 0) {
      args.push(raw.slice(start, i).trim());
      start = i + 1;
    }
  }
  if (quote || depth !== 0) return null;
  args.push(raw.slice(start).trim());
  return args.every((arg) => arg.length > 0) ? args : null;
}

type FormulaArithmeticOp = '+' | '-' | '*' | '/' | '^';

function splitFormulaArithmetic(body: string): {
  left: string;
  op: FormulaArithmeticOp;
  right: string;
} | null {
  const operatorsByPrecedence: FormulaArithmeticOp[][] = [['+', '-'], ['*', '/'], ['^']];
  for (const ops of operatorsByPrecedence) {
    let depth = 0;
    let quote: '"' | "'" | null = null;
    const start = ops.includes('^') ? 0 : body.length - 1;
    const end = ops.includes('^') ? body.length : -1;
    const step = ops.includes('^') ? 1 : -1;
    for (let i = start; i !== end; i += step) {
      const ch = body[i];
      if (quote) {
        if (ch === quote) quote = null;
        continue;
      }
      if (ch === '"' || ch === "'") {
        quote = ch;
        continue;
      }
      if (ch === ')') {
        depth += step < 0 ? 1 : -1;
        if (depth < 0) return null;
        continue;
      }
      if (ch === '(') {
        depth += step < 0 ? -1 : 1;
        if (depth < 0) return null;
        continue;
      }
      if (depth !== 0 || !ops.includes(ch as FormulaArithmeticOp)) continue;
      if ((ch === '+' || ch === '-') && isUnaryArithmeticSign(body, i)) continue;
      const left = body.slice(0, i).trim();
      const right = body.slice(i + 1).trim();
      if (left.length === 0 || right.length === 0) continue;
      return { left, op: ch as FormulaArithmeticOp, right };
    }
  }
  return null;
}

function isUnaryArithmeticSign(body: string, index: number): boolean {
  for (let i = index - 1; i >= 0; i -= 1) {
    const ch = body[i];
    if (ch === ' ') continue;
    return ch === '(' || ch === '+' || ch === '-' || ch === '*' || ch === '/';
  }
  return true;
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
  const readOperand = makeFormulaOperandReader(state, sheet);
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
): (operand: FormulaOperand, rowOffset: number, colOffset: number) => CellValue {
  const resolveRef = (ref: ParsedRef, rowOffset: number, colOffset: number): [number, number] => [
    ref.absRow ? ref.row : ref.row + rowOffset,
    ref.absCol ? ref.col : ref.col + colOffset,
  ];
  const aggregateRange = (
    range: ParsedA1Range,
    fn: FormulaAggregateName,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const bounds = rangeBounds(range, rowOffset, colOffset);
    const { r0, r1, c0, c1 } = bounds;
    if (!validRangeBounds(bounds)) return { kind: 'error', code: 15, text: '#VALUE!' };
    const values: number[] = [];
    for (let r = r0; r <= r1; r += 1) {
      for (let c = c0; c <= c1; c += 1) {
        const value = state.data.cells.get(addrKey({ sheet, row: r, col: c }))?.value;
        if (value?.kind === 'number' && Number.isFinite(value.value)) values.push(value.value);
      }
    }
    if (fn === 'COUNT') return { kind: 'number', value: values.length };
    if (fn === 'SUM') {
      return { kind: 'number', value: values.reduce((sum, value) => sum + value, 0) };
    }
    if (values.length === 0) return { kind: 'error', code: 15, text: '#VALUE!' };
    if (fn === 'AVERAGE') {
      return {
        kind: 'number',
        value: values.reduce((sum, value) => sum + value, 0) / values.length,
      };
    }
    return {
      kind: 'number',
      value: fn === 'MIN' ? Math.min(...values) : Math.max(...values),
    };
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
  const validRangeBounds = (bounds: { width: number; height: number }): boolean =>
    bounds.width > 0 &&
    bounds.height > 0 &&
    bounds.width * bounds.height <= MAX_FORMULA_AGGREGATE_CELLS;
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
    range: ParsedA1Range,
    criteria: CellValue,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const bounds = rangeBounds(range, rowOffset, colOffset);
    if (!validRangeBounds(bounds)) {
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
    pairs: { range: ParsedA1Range; criteria: CellValue }[],
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const bounds = pairs.map((pair) => rangeBounds(pair.range, rowOffset, colOffset));
    const first = bounds[0];
    if (!first || !bounds.every(validRangeBounds)) {
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
  const readOperand = (
    operand: FormulaOperand,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    if (operand.kind === 'literal') return operand.value;
    if (operand.kind === 'range-aggregate') {
      return aggregateRange(operand.range, operand.fn, rowOffset, colOffset);
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
    if (operand.kind === 'binary') {
      const leftValue = readOperand(operand.left, rowOffset, colOffset);
      const rightValue = readOperand(operand.right, rowOffset, colOffset);
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

function splitFormulaComparison(
  body: string,
): { left: string; op: '>' | '<' | '>=' | '<=' | '=' | '<>'; right: string } | null {
  let depth = 0;
  let quote: '"' | "'" | null = null;
  for (let i = 0; i < body.length; i += 1) {
    const ch = body[i];
    if (quote) {
      if (ch === quote) quote = null;
      continue;
    }
    if (ch === '"' || ch === "'") {
      quote = ch;
      continue;
    }
    if (ch === '(') {
      depth += 1;
      continue;
    }
    if (ch === ')') {
      depth -= 1;
      if (depth < 0) return null;
      continue;
    }
    if (depth !== 0) continue;
    const two = body.slice(i, i + 2);
    const op =
      two === '>=' || two === '<=' || two === '<>'
        ? two
        : ch === '>' || ch === '<' || ch === '='
          ? ch
          : null;
    if (!op) continue;
    const left = body.slice(0, i).trim();
    const right = body.slice(i + op.length).trim();
    return left && right ? { left, op, right } : null;
  }
  return null;
}

/** Stable canonical key for a cell value, used by the duplicates / unique
 *  predicates. Blank cells are skipped (returns null). */
function valueKey(v: CellValue): string | null {
  switch (v.kind) {
    case 'blank':
      return null;
    case 'number':
      return `n:${v.value}`;
    case 'bool':
      return v.value ? 'b:1' : 'b:0';
    case 'text':
      return `t:${v.value}`;
    case 'error':
      return `e:${v.text}`;
  }
}

const isErrorValue = (v: CellValue): boolean => v.kind === 'error';
const isBlankValue = (v: CellValue): boolean => v.kind === 'blank';

/**
 * Evaluate conditional formatting rules for the active sheet's cells. We
 * compute per-rule numeric extremes for color-scale / data-bar rules once,
 * then walk the cell entries assigning overlays.
 */
export function evaluateConditional(state: State): Map<string, ConditionalCellOverlay> {
  if (
    cachedOverlay !== null &&
    cachedRulesRef === state.conditional.rules &&
    cachedCellsRef === state.data.cells &&
    cachedSheet === state.data.sheetIndex
  ) {
    return cachedOverlay;
  }
  const out = new Map<string, ConditionalCellOverlay>();
  const rules = state.conditional.rules;
  if (rules.length === 0) {
    cachedRulesRef = rules;
    cachedCellsRef = state.data.cells;
    cachedSheet = state.data.sheetIndex;
    cachedOverlay = out;
    return out;
  }
  const sheet = state.data.sheetIndex;
  const stopped = new Set<string>();

  for (let ri = 0; ri < rules.length; ri += 1) {
    const rule = rules[ri];
    if (!rule) continue;
    if (rule.range.sheet !== sheet) continue;
    const ruleOverlay = new Map<string, ConditionalCellOverlay>();

    if (rule.kind === 'cell-value') {
      paintCellValue(state, rule, ruleOverlay);
    } else if (rule.kind === 'color-scale') {
      paintColorScale(state, rule, ruleOverlay);
    } else if (rule.kind === 'data-bar') {
      paintDataBar(state, rule, ruleOverlay);
    } else if (rule.kind === 'icon-set') {
      paintIconSet(state, rule, ruleOverlay);
    } else if (rule.kind === 'top-bottom') {
      paintTopBottom(state, rule, ruleOverlay);
    } else if (rule.kind === 'average') {
      paintAverage(state, rule, ruleOverlay);
    } else if (rule.kind === 'text-contains') {
      paintTextContains(state, rule, ruleOverlay);
    } else if (rule.kind === 'date-occurring') {
      paintDateOccurring(state, rule, ruleOverlay);
    } else if (rule.kind === 'formula') {
      paintFormula(state, rule, ruleOverlay);
    } else if (rule.kind === 'duplicates' || rule.kind === 'unique') {
      paintDupsUnique(state, rule, ruleOverlay);
    } else if (
      rule.kind === 'blanks' ||
      rule.kind === 'non-blanks' ||
      rule.kind === 'errors' ||
      rule.kind === 'no-errors'
    ) {
      paintBlankErrorPredicate(state, rule, ruleOverlay);
    }

    for (const [key, overlay] of ruleOverlay) {
      if (stopped.has(key)) continue;
      const target = out.get(key) ?? {};
      mergeOverlayByPriority(target, overlay);
      out.set(key, target);
      if (rule.stopIfTrue === true) stopped.add(key);
    }
  }

  cachedRulesRef = state.conditional.rules;
  cachedCellsRef = state.data.cells;
  cachedSheet = state.data.sheetIndex;
  cachedOverlay = out;
  return out;
}

function paintCellValue(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'cell-value' }>,
  out: Map<string, ConditionalCellOverlay>,
): void {
  const sheet = state.data.sheetIndex;
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      if (!cell) continue;
      if (!inRange(sheet, r, c, rule.range)) continue;
      if (testCellValue(cell.value, rule.op, rule.a, rule.b)) {
        const overlay = out.get(key) ?? {};
        mergeApply(overlay, rule.apply);
        out.set(key, overlay);
      }
    }
  }
}

function paintColorScale(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'color-scale' }>,
  out: Map<string, ConditionalCellOverlay>,
): void {
  const sheet = state.data.sheetIndex;
  const values: number[] = [];
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const cell = state.data.cells.get(addrKey({ sheet, row: r, col: c }));
      if (cell?.value.kind !== 'number') continue;
      values.push(cell.value.value);
    }
  }
  if (values.length === 0) return;
  const scale = colorScaleThresholds(rule, values);
  if (!scale) return;
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      if (cell?.value.kind !== 'number') continue;
      const v = cell.value.value;
      const t = colorScalePosition(v, scale);
      const overlay = out.get(key) ?? {};
      overlay.fill = pickStop(rule.stops, t);
      out.set(key, overlay);
    }
  }
}

interface ColorScaleThresholds {
  low: number;
  mid?: number;
  high: number;
}

function colorScaleThresholds(
  rule: Extract<ConditionalRule, { kind: 'color-scale' }>,
  values: readonly number[],
): ColorScaleThresholds | null {
  const sorted = values
    .filter((value) => Number.isFinite(value))
    .slice()
    .sort((a, b) => a - b);
  if (sorted.length === 0) return null;
  const defaultThresholds =
    rule.stops.length === 2
      ? ([{ kind: 'min' }, { kind: 'max' }] as const)
      : ([{ kind: 'min' }, { kind: 'percentile', value: 50 }, { kind: 'max' }] as const);
  const thresholds = rule.thresholds ?? defaultThresholds;
  const low = resolveScalePoint(thresholds[0] ?? { kind: 'min' }, sorted);
  const high = resolveScalePoint(thresholds[thresholds.length - 1] ?? { kind: 'max' }, sorted);
  if (rule.stops.length === 2) return { low, high };
  const mid = resolveScalePoint(thresholds[1] ?? { kind: 'percentile', value: 50 }, sorted);
  return { low, mid, high };
}

function resolveScalePoint(point: ConditionalScalePoint, sorted: readonly number[]): number {
  const min = sorted[0] ?? 0;
  const max = sorted[sorted.length - 1] ?? min;
  if (point.kind === 'min') return min;
  if (point.kind === 'max') return max;
  if (point.kind === 'number') return point.value;
  if (point.kind !== 'percent' && point.kind !== 'percentile') return min;
  const pct = Math.max(0, Math.min(100, point.value));
  if (point.kind === 'percent') return min + ((max - min) * pct) / 100;
  const rank = ((sorted.length - 1) * pct) / 100;
  const lo = Math.floor(rank);
  const hi = Math.ceil(rank);
  const a = sorted[lo] ?? min;
  const b = sorted[hi] ?? a;
  return a + (b - a) * (rank - lo);
}

function colorScalePosition(value: number, thresholds: ColorScaleThresholds): number {
  const low = thresholds.low;
  const high = thresholds.high;
  const mid = thresholds.mid;
  if (mid === undefined) {
    if (high === low) return 0.5;
    return Math.max(0, Math.min(1, (value - low) / (high - low)));
  }
  if (high === low) return 0.5;
  if (value <= mid) {
    if (mid === low) return 0.5;
    return Math.max(0, Math.min(0.5, ((value - low) / (mid - low)) * 0.5));
  }
  if (high === mid) return 0.5;
  return Math.max(0.5, Math.min(1, 0.5 + ((value - mid) / (high - mid)) * 0.5));
}

function paintDataBar(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'data-bar' }>,
  out: Map<string, ConditionalCellOverlay>,
): void {
  const sheet = state.data.sheetIndex;
  let min = Number.POSITIVE_INFINITY;
  let max = Number.NEGATIVE_INFINITY;
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const cell = state.data.cells.get(addrKey({ sheet, row: r, col: c }));
      if (cell?.value.kind !== 'number') continue;
      const v = cell.value.value;
      if (v < min) min = v;
      if (v > max) max = v;
    }
  }
  if (!Number.isFinite(min)) return;
  const positiveDenom = Math.max(max, 1e-9);
  const negativeDenom = Math.max(Math.abs(min), 1e-9);
  const axis =
    min < 0 && max > 0
      ? Math.max(0, Math.min(1, Math.abs(min) / (Math.abs(min) + max)))
      : max <= 0
        ? 1
        : 0;
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      if (cell?.value.kind !== 'number') continue;
      const v = cell.value.value;
      const overlay = out.get(key) ?? {};
      const negative = v < 0;
      overlay.bar = negative
        ? Math.max(0, Math.min(axis, (Math.abs(v) / negativeDenom) * axis))
        : Math.max(0, Math.min(1 - axis, (v / positiveDenom) * (1 - axis)));
      overlay.barAxis = axis;
      overlay.barDirection = negative ? 'left' : 'right';
      overlay.barColor = rule.color;
      overlay.barGradient = rule.gradient === true;
      overlay.showValue = rule.showValue !== false;
      out.set(key, overlay);
    }
  }
}

function paintIconSet(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'icon-set' }>,
  out: Map<string, ConditionalCellOverlay>,
): void {
  const sheet = state.data.sheetIndex;
  const values: number[] = [];
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const cell = state.data.cells.get(addrKey({ sheet, row: r, col: c }));
      if (cell?.value.kind !== 'number') continue;
      values.push(cell.value.value);
    }
  }
  const sorted = values
    .filter((value) => Number.isFinite(value))
    .slice()
    .sort((a, b) => a - b);
  if (sorted.length === 0) return;
  const min = sorted[0] ?? 0;
  const max = sorted[sorted.length - 1] ?? min;
  const slots = iconSetSlotCount(rule.icons);
  const thresholds = iconSetThresholdValues(rule, sorted);
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      if (cell?.value.kind !== 'number') continue;
      const v = cell.value.value;
      const t = max === min ? 0.5 : (v - min) / (max - min);
      let slot =
        thresholds === null
          ? iconSetSlotFor(rule.icons, t)
          : thresholds.reduce((count, threshold) => (v >= threshold ? count + 1 : count), 0);
      slot = Math.max(0, Math.min(slots - 1, slot));
      if (rule.reverseOrder) slot = slots - 1 - slot;
      const overlay = out.get(key) ?? {};
      overlay.iconKind = rule.icons;
      overlay.iconSlot = slot;
      overlay.showValue = rule.showValue !== false;
      out.set(key, overlay);
    }
  }
}

function iconSetThresholdValues(
  rule: Extract<ConditionalRule, { kind: 'icon-set' }>,
  sorted: readonly number[],
): number[] | null {
  const slots = iconSetSlotCount(rule.icons);
  if (!rule.thresholds || rule.thresholds.length === 0) return null;
  return rule.thresholds
    .slice(0, slots - 1)
    .map((point) => resolveScalePoint(point, sorted))
    .sort((a, b) => a - b);
}

function paintTopBottom(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'top-bottom' }>,
  out: Map<string, ConditionalCellOverlay>,
): void {
  const sheet = state.data.sheetIndex;
  const values: number[] = [];
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const cell = state.data.cells.get(addrKey({ sheet, row: r, col: c }));
      if (cell && cell.value.kind === 'number') values.push(cell.value.value);
    }
  }
  const cutoff = topBottomThreshold(values, rule.mode, rule.n, rule.percent ?? false);
  if (cutoff === null) return;
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      if (cell?.value.kind !== 'number') continue;
      const v = cell.value.value;
      const passes = rule.mode === 'top' ? v >= cutoff : v <= cutoff;
      if (!passes) continue;
      const overlay = out.get(key) ?? {};
      mergeApply(overlay, rule.apply);
      out.set(key, overlay);
    }
  }
}

function paintAverage(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'average' }>,
  out: Map<string, ConditionalCellOverlay>,
): void {
  const sheet = state.data.sheetIndex;
  const values: number[] = [];
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const cell = state.data.cells.get(addrKey({ sheet, row: r, col: c }));
      if (cell?.value.kind === 'number' && Number.isFinite(cell.value.value)) {
        values.push(cell.value.value);
      }
    }
  }
  if (values.length === 0) return;
  const avg = values.reduce((sum, v) => sum + v, 0) / values.length;
  const variance = values.reduce((sum, v) => sum + (v - avg) ** 2, 0) / values.length;
  const stdDev = Math.sqrt(variance);
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      if (cell?.value.kind !== 'number') continue;
      const v = cell.value.value;
      const passes =
        rule.mode === 'above'
          ? v > avg
          : rule.mode === 'below'
            ? v < avg
            : rule.mode === 'equal-or-above'
              ? v >= avg
              : rule.mode === 'equal-or-below'
                ? v <= avg
                : rule.mode === 'above-std-dev'
                  ? v > avg + stdDev * (rule.stdDev ?? 1)
                  : v < avg - stdDev * (rule.stdDev ?? 1);
      if (!passes) continue;
      const overlay = out.get(key) ?? {};
      mergeApply(overlay, rule.apply);
      out.set(key, overlay);
    }
  }
}

function cellText(v: CellValue): string | null {
  if (v.kind === 'text') return v.value;
  if (v.kind === 'number') return String(v.value);
  if (v.kind === 'bool') return v.value ? 'TRUE' : 'FALSE';
  if (v.kind === 'error') return v.text;
  return null;
}

function paintTextContains(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'text-contains' }>,
  out: Map<string, ConditionalCellOverlay>,
): void {
  const needle = rule.caseSensitive ? rule.text : rule.text.toLocaleLowerCase();
  if (needle.length === 0) return;
  const sheet = state.data.sheetIndex;
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      if (!cell) continue;
      const raw = cellText(cell.value);
      if (raw === null) continue;
      const haystack = rule.caseSensitive ? raw : raw.toLocaleLowerCase();
      const matches =
        rule.mode === 'not-contains'
          ? !haystack.includes(needle)
          : rule.mode === 'begins-with'
            ? haystack.startsWith(needle)
            : rule.mode === 'ends-with'
              ? haystack.endsWith(needle)
              : haystack.includes(needle);
      if (!matches) continue;
      const overlay = out.get(key) ?? {};
      mergeApply(overlay, rule.apply);
      out.set(key, overlay);
    }
  }
}

const DAY_MS = 86_400_000;

function normalizeDate(d: Date): number {
  return Math.floor(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()) / DAY_MS);
}

function excelSerialToDate(serial: number): Date {
  return new Date(Date.UTC(1899, 11, 30) + Math.floor(serial) * DAY_MS);
}

function cellDateDay(v: CellValue): number | null {
  if (v.kind === 'number' && Number.isFinite(v.value))
    return normalizeDate(excelSerialToDate(v.value));
  if (v.kind === 'text') {
    const time = Date.parse(v.value);
    if (Number.isFinite(time)) return normalizeDate(new Date(time));
  }
  return null;
}

function weekStart(day: number): number {
  const d = new Date(day * DAY_MS);
  const dow = (d.getUTCDay() + 6) % 7;
  return day - dow;
}

function monthKey(day: number): number {
  const d = new Date(day * DAY_MS);
  return d.getUTCFullYear() * 12 + d.getUTCMonth();
}

function datePeriodMatches(
  day: number,
  period: Extract<ConditionalRule, { kind: 'date-occurring' }>['period'],
): boolean {
  const today = normalizeDate(new Date());
  switch (period) {
    case 'yesterday':
      return day === today - 1;
    case 'today':
      return day === today;
    case 'tomorrow':
      return day === today + 1;
    case 'last7':
      return day >= today - 6 && day <= today;
    case 'last-week':
      return weekStart(day) === weekStart(today) - 7;
    case 'this-week':
      return weekStart(day) === weekStart(today);
    case 'next-week':
      return weekStart(day) === weekStart(today) + 7;
    case 'last-month':
      return monthKey(day) === monthKey(today) - 1;
    case 'this-month':
      return monthKey(day) === monthKey(today);
    case 'next-month':
      return monthKey(day) === monthKey(today) + 1;
    default:
      return false;
  }
}

function paintDateOccurring(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'date-occurring' }>,
  out: Map<string, ConditionalCellOverlay>,
): void {
  const sheet = state.data.sheetIndex;
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      if (!cell) continue;
      const day = cellDateDay(cell.value);
      if (day === null || !datePeriodMatches(day, rule.period)) continue;
      const overlay = out.get(key) ?? {};
      mergeApply(overlay, rule.apply);
      out.set(key, overlay);
    }
  }
}

function paintFormula(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'formula' }>,
  out: Map<string, ConditionalCellOverlay>,
): void {
  const formulaPredicate = parseFormulaCellPredicate(state, rule);
  const predicate = parseFormulaPredicate(rule.formula);
  if (!formulaPredicate && !predicate) return;
  const sheet = state.data.sheetIndex;
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      const passes = formulaPredicate?.test(r, c) ?? (cell ? predicate?.test(cell.value) : false);
      if (!passes) continue;
      const overlay = out.get(key) ?? {};
      mergeApply(overlay, rule.apply);
      out.set(key, overlay);
    }
  }
}

function paintDupsUnique(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'duplicates' | 'unique' }>,
  out: Map<string, ConditionalCellOverlay>,
): void {
  const sheet = state.data.sheetIndex;
  const counts = new Map<string, number>();
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const cell = state.data.cells.get(addrKey({ sheet, row: r, col: c }));
      if (!cell) continue;
      const k = valueKey(cell.value);
      if (k === null) continue;
      counts.set(k, (counts.get(k) ?? 0) + 1);
    }
  }
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      if (!cell) continue;
      const k = valueKey(cell.value);
      if (k === null) continue;
      const count = counts.get(k) ?? 0;
      const passes = rule.kind === 'duplicates' ? count > 1 : count === 1;
      if (!passes) continue;
      const overlay = out.get(key) ?? {};
      mergeApply(overlay, rule.apply);
      out.set(key, overlay);
    }
  }
}

function paintBlankErrorPredicate(
  state: State,
  rule: Extract<ConditionalRule, { kind: 'blanks' | 'non-blanks' | 'errors' | 'no-errors' }>,
  out: Map<string, ConditionalCellOverlay>,
): void {
  const sheet = state.data.sheetIndex;
  for (let r = rule.range.r0; r <= rule.range.r1; r += 1) {
    for (let c = rule.range.c0; c <= rule.range.c1; c += 1) {
      const key = addrKey({ sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      const value: CellValue = cell?.value ?? { kind: 'blank' };
      let passes = false;
      if (rule.kind === 'blanks') passes = isBlankValue(value);
      else if (rule.kind === 'non-blanks') passes = !isBlankValue(value);
      else if (rule.kind === 'errors') passes = isErrorValue(value);
      else if (rule.kind === 'no-errors') passes = !isErrorValue(value) && !isBlankValue(value);
      if (!passes) continue;
      const overlay = out.get(key) ?? {};
      mergeApply(overlay, rule.apply);
      out.set(key, overlay);
    }
  }
}

function testCellValue(
  value: CellValue,
  op: '>' | '<' | '>=' | '<=' | '=' | '<>' | 'between' | 'not-between',
  a: number | string,
  b: number | string | undefined,
): boolean {
  const v = value.kind === 'number' ? value.value : cellText(value);
  if (v === null) return false;
  if (typeof v === 'number' && typeof a === 'number') {
    return testComparableValue(v, op, a, typeof b === 'number' ? b : undefined);
  }
  return testComparableValue(
    String(v).toLocaleLowerCase(),
    op,
    String(a).toLocaleLowerCase(),
    b === undefined ? undefined : String(b).toLocaleLowerCase(),
  );
}

function testComparableValue<T extends number | string>(
  v: T,
  op: '>' | '<' | '>=' | '<=' | '=' | '<>' | 'between' | 'not-between',
  a: T,
  b: T | undefined,
): boolean {
  switch (op) {
    case '>':
      return v > a;
    case '<':
      return v < a;
    case '>=':
      return v >= a;
    case '<=':
      return v <= a;
    case '=':
      return v === a;
    case '<>':
      return v !== a;
    case 'between':
      return b !== undefined && v >= (a <= b ? a : b) && v <= (a <= b ? b : a);
    case 'not-between':
      return b !== undefined && (v < (a <= b ? a : b) || v > (a <= b ? b : a));
    default:
      return false;
  }
}

function mergeApply(target: ConditionalCellOverlay, patch: Partial<CellFormat>): void {
  if (patch.fill) target.fill = patch.fill;
  if (patch.color) target.color = patch.color;
  if (patch.bold) target.bold = true;
  if (patch.italic) target.italic = true;
  if (patch.underline) target.underline = true;
  if (patch.strike) target.strike = true;
}

function mergeOverlayByPriority(
  target: ConditionalCellOverlay,
  source: ConditionalCellOverlay,
): void {
  if (target.fill === undefined && source.fill !== undefined) target.fill = source.fill;
  if (target.color === undefined && source.color !== undefined) target.color = source.color;
  if (target.bold === undefined && source.bold !== undefined) target.bold = source.bold;
  if (target.italic === undefined && source.italic !== undefined) target.italic = source.italic;
  if (target.underline === undefined && source.underline !== undefined) {
    target.underline = source.underline;
  }
  if (target.strike === undefined && source.strike !== undefined) target.strike = source.strike;
  if (target.bar === undefined && source.bar !== undefined) {
    target.bar = source.bar;
    target.barAxis = source.barAxis;
    target.barDirection = source.barDirection;
    target.barColor = source.barColor;
    target.barGradient = source.barGradient;
  }
  if (target.iconKind === undefined && source.iconKind !== undefined) {
    target.iconKind = source.iconKind;
    target.iconSlot = source.iconSlot;
  }
  if (target.showValue === undefined && source.showValue !== undefined) {
    target.showValue = source.showValue;
  }
}

function pickStop(stops: readonly string[], t: number): string {
  const s0 = stops[0] ?? '#000000';
  const s1 = stops[1] ?? s0;
  const s2 = stops[2] ?? s1;
  if (stops.length === 2) return interpolate(s0, s1, t);
  // Three-stop: low, mid, high
  if (t <= 0.5) return interpolate(s0, s1, t * 2);
  return interpolate(s1, s2, (t - 0.5) * 2);
}

function interpolate(a: string, b: string, t: number): string {
  const ca = parseColor(a);
  const cb = parseColor(b);
  if (!ca || !cb) return a;
  const r = Math.round(ca[0] + (cb[0] - ca[0]) * t);
  const g = Math.round(ca[1] + (cb[1] - ca[1]) * t);
  const blu = Math.round(ca[2] + (cb[2] - ca[2]) * t);
  return `rgb(${r}, ${g}, ${blu})`;
}

function parseColor(s: string): [number, number, number] | null {
  const m = s.trim().match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/i);
  if (m) {
    const hex = m[1] ?? '';
    if (hex.length === 3) {
      const h0 = hex[0] ?? '0';
      const h1 = hex[1] ?? '0';
      const h2 = hex[2] ?? '0';
      return [
        Number.parseInt(h0 + h0, 16),
        Number.parseInt(h1 + h1, 16),
        Number.parseInt(h2 + h2, 16),
      ];
    }
    return [
      Number.parseInt(hex.slice(0, 2), 16),
      Number.parseInt(hex.slice(2, 4), 16),
      Number.parseInt(hex.slice(4, 6), 16),
    ];
  }
  const rgb = s.match(/^rgb\((\d+),\s*(\d+),\s*(\d+)\)$/);
  if (rgb) {
    return [
      Number.parseInt(rgb[1] ?? '0', 10),
      Number.parseInt(rgb[2] ?? '0', 10),
      Number.parseInt(rgb[3] ?? '0', 10),
    ];
  }
  return null;
}

/** Used by ConditionalRule consumer types — re-exported through index. */
export type { ConditionalRule };
