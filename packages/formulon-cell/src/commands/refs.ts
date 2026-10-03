import { colFromLetters, formatA1Cell, MAX_COL, MAX_ROW, parseA1Atom } from '../engine/address.js';

/**
 * F4 reference rotation: cycles between A1, $A$1, A$1, $A1, then back to A1.
 * Operates on the cell reference at `caret` in `text`. Returns the new text
 * + new caret position. If no reference is found, returns the input unchanged.
 */
const REF_RE = /(\$?)([A-Za-z]+)(\$?)([0-9]+)/;

/** A1-style cell ref or A1:B5 range surfaced in a formula text. */
export interface FormulaRef {
  /** 0-indexed inclusive bounds. */
  r0: number;
  c0: number;
  r1: number;
  c1: number;
  /** Character offsets in the source text (start inclusive, end exclusive). */
  start: number;
  end: number;
  /** Color index 0..N for distinct highlighting. */
  colorIndex: number;
}

/** Extract every cell or range reference from a formula text. The returned
 *  list is in source order with a stable per-target color index so callers
 *  can paint distinct highlights for each ref, the same way spreadsheets do
 *  while editing a formula. */
export function extractRefs(text: string): FormulaRef[] {
  if (!text.startsWith('=')) return [];
  // Match: optional sheet prefix (Sheet1!), then A1 or A1:B5.
  const re =
    /(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_]*))?!?(\$?[A-Za-z]+\$?\d+)(?::(\$?[A-Za-z]+\$?\d+))?/g;
  const out: FormulaRef[] = [];
  const colorMap = new Map<string, number>();
  re.lastIndex = 0;
  for (let m = re.exec(text); m !== null; m = re.exec(text)) {
    const headM = m[3] ?? '';
    const tailM = m[4];
    const head = parseA1Atom(headM);
    const tail = tailM ? parseA1Atom(tailM) : head;
    if (!head || !tail) continue;
    // Skip false positives — function names like SIN1 don't have a digit-letter
    //  shape in the head capture, so the regex naturally rejects them. But
    //  we still need to skip when the ref start is inside a quoted string.
    const before = text.slice(0, m.index);
    const quoteCount = (before.match(/"/g) ?? []).length;
    if (quoteCount % 2 === 1) continue;
    const r0 = Math.min(head.row, tail.row);
    const r1 = Math.max(head.row, tail.row);
    const c0 = Math.min(head.col, tail.col);
    const c1 = Math.max(head.col, tail.col);
    const key = `${r0}:${c0}:${r1}:${c1}`;
    let colorIndex = colorMap.get(key);
    if (colorIndex === undefined) {
      colorIndex = colorMap.size;
      colorMap.set(key, colorIndex);
    }
    out.push({
      r0,
      c0,
      r1,
      c1,
      start: m.index,
      end: m.index + m[0].length,
      colorIndex,
    });
  }
  return out;
}

/** Distinct accent colors used for formula-edit reference highlighting. They
 *  loop after this list is exhausted (spreadsheets use ~8). */
export const REF_HIGHLIGHT_COLORS: readonly string[] = [
  '#1f7ae0',
  '#d96f2c',
  '#3aa757',
  '#a83cb2',
  '#cf3a4c',
  '#1f998c',
  '#946a00',
  '#3953c4',
];

export interface F4Result {
  text: string;
  caret: number;
}

export function rotateRefAt(text: string, caret: number): F4Result {
  // Walk left from caret to find a candidate ref start.
  // We bound the search to ~16 chars left.
  const start = Math.max(0, caret - 16);
  const window = text.slice(start, caret + 16);
  const offset = start;
  const re = new RegExp(REF_RE, 'g');
  let chosen: { match: RegExpExecArray; absoluteStart: number } | null = null;
  re.lastIndex = 0;
  for (let m = re.exec(window); m !== null; m = re.exec(window)) {
    const matchStart = offset + m.index;
    const matchEnd = matchStart + m[0].length;
    if (caret >= matchStart && caret <= matchEnd) {
      chosen = { match: m, absoluteStart: matchStart };
      break;
    }
  }
  if (!chosen) return { text, caret };
  const [whole, dCol, letters, dRow, digits] = chosen.match;
  const next = nextStep(dCol === '$', dRow === '$');
  const replacement = `${next.col ? '$' : ''}${letters}${next.row ? '$' : ''}${digits}`;
  const before = text.slice(0, chosen.absoluteStart);
  const after = text.slice(chosen.absoluteStart + whole.length);
  return {
    text: before + replacement + after,
    caret: chosen.absoluteStart + replacement.length,
  };
}

function nextStep(absCol: boolean, absRow: boolean): { col: boolean; row: boolean } {
  // Spreadsheet order: A1 -> $A$1 -> A$1 -> $A1 -> A1
  if (!absCol && !absRow) return { col: true, row: true };
  if (absCol && absRow) return { col: false, row: true };
  if (!absCol && absRow) return { col: true, row: false };
  return { col: false, row: false };
}

const R1C1_REF_RE = /R(?:\[(-?\d+)\]|([1-9]\d*))?C(?:\[(-?\d+)\]|([1-9]\d*))?/iy;

const r1c1AtomToA1 = (raw: string, base: { row: number; col: number }): string => {
  const m = /^R(?:\[(-?\d+)\]|([1-9]\d*))?C(?:\[(-?\d+)\]|([1-9]\d*))?$/i.exec(raw);
  if (!m) return raw;
  const row =
    m[2] !== undefined
      ? Number.parseInt(m[2], 10) - 1
      : base.row + (m[1] !== undefined ? Number.parseInt(m[1], 10) : 0);
  const col =
    m[4] !== undefined
      ? Number.parseInt(m[4], 10) - 1
      : base.col + (m[3] !== undefined ? Number.parseInt(m[3], 10) : 0);
  if (row < 0 || col < 0 || row > MAX_ROW || col > MAX_COL) return '#REF!';
  return formatA1Cell(row, col);
};

/** Convert R1C1 references in a user-authored formula into A1 references for
 *  the workbook engine. This keeps the public R1C1 editing surface while
 *  preserving the existing A1 formula backend. */
export function normalizeR1C1Formula(formula: string, base: { row: number; col: number }): string {
  if (!formula.startsWith('=')) return formula;
  let out = '';
  let i = 0;
  let inString = false;
  while (i < formula.length) {
    const ch = formula[i] ?? '';
    if (ch === '"') {
      inString = !inString;
      out += ch;
      i += 1;
      continue;
    }
    if (!inString) {
      const prev = i > 0 ? (formula[i - 1] ?? '') : '';
      const prevIsIdent = /[A-Za-z0-9_]/.test(prev);
      if (!prevIsIdent) {
        R1C1_REF_RE.lastIndex = i;
        const m = R1C1_REF_RE.exec(formula);
        if (m && m.index === i) {
          const raw = m[0];
          const next = formula[i + raw.length] ?? '';
          if (!/[A-Za-z0-9_[]/.test(next)) {
            out += r1c1AtomToA1(raw, base);
            i += raw.length;
            continue;
          }
        }
      }
    }
    out += ch;
    i += 1;
  }
  return out;
}

const a1AtomToR1C1 = (raw: string, base: { row: number; col: number }): string => {
  const m = /^(\$?)([A-Za-z]+)(\$?)(\d+)$/.exec(raw);
  if (!m) return raw;
  const colAbs = m[1] === '$';
  const rowAbs = m[3] === '$';
  const col = colFromLetters(m[2] ?? '');
  const row = Number.parseInt(m[4] ?? '', 10) - 1;
  if (row < 0 || col < 0 || row > MAX_ROW || col > MAX_COL) return raw;
  const r = rowAbs ? `R${row + 1}` : row === base.row ? 'R' : `R[${row - base.row}]`;
  const c = colAbs ? `C${col + 1}` : col === base.col ? 'C' : `C[${col - base.col}]`;
  return `${r}${c}`;
};

/** Format stored A1 formulas as R1C1 for formula-bar / inline-editor display.
 *  The workbook still stores A1 formulas; this is a presentation transform. */
export function formatA1FormulaAsR1C1(formula: string, base: { row: number; col: number }): string {
  if (!formula.startsWith('=')) return formula;
  let out = '';
  let i = 0;
  let inString = false;
  const atomRe = /\$?[A-Za-z]+\$?\d+/y;
  while (i < formula.length) {
    const ch = formula[i] ?? '';
    if (ch === '"') {
      inString = !inString;
      out += ch;
      i += 1;
      continue;
    }
    if (!inString) {
      const prev = i > 0 ? (formula[i - 1] ?? '') : '';
      const prevIsIdent = /[A-Za-z0-9_]/.test(prev);
      if (!prevIsIdent) {
        atomRe.lastIndex = i;
        const m = atomRe.exec(formula);
        if (m && m.index === i) {
          out += a1AtomToR1C1(m[0], base);
          i += m[0].length;
          continue;
        }
      }
    }
    out += ch;
    i += 1;
  }
  return out;
}

// `shiftFormulaRefs` (relative-offset shifting for fill/paste/ctrl-enter) now
// lives in the shared reference-rewriting module alongside the structural-edit
// transforms, so all three share one tokenizer (function-name, sheet-qualifier,
// range, and grid-bound handling). Re-exported here for existing import sites.
export { shiftFormulaRefs } from './formula-refs.js';

/** Common spreadsheet function names — used for editor autocomplete. */
export const FUNCTION_NAMES: readonly string[] = [
  'SUM',
  'AVERAGE',
  'COUNT',
  'COUNTA',
  'COUNTIF',
  'COUNTIFS',
  'SUMIF',
  'SUMIFS',
  'AVERAGEIF',
  'AVERAGEIFS',
  'MIN',
  'MAX',
  'MEDIAN',
  'IF',
  'IFS',
  'IFERROR',
  'IFNA',
  'AND',
  'OR',
  'NOT',
  'XOR',
  'TRUE',
  'FALSE',
  'VLOOKUP',
  'HLOOKUP',
  'XLOOKUP',
  'INDEX',
  'MATCH',
  'OFFSET',
  'INDIRECT',
  'CHOOSE',
  'ROUND',
  'ROUNDUP',
  'ROUNDDOWN',
  'CEILING',
  'FLOOR',
  'INT',
  'MOD',
  'ABS',
  'POWER',
  'SQRT',
  'EXP',
  'LN',
  'LOG',
  'LOG10',
  'CONCATENATE',
  'CONCAT',
  'TEXTJOIN',
  'LEFT',
  'RIGHT',
  'MID',
  'LEN',
  'UPPER',
  'LOWER',
  'PROPER',
  'TRIM',
  'SUBSTITUTE',
  'REPLACE',
  'FIND',
  'SEARCH',
  'TEXT',
  'VALUE',
  'NUMBERVALUE',
  'TODAY',
  'NOW',
  'DATE',
  'YEAR',
  'MONTH',
  'DAY',
  'HOUR',
  'MINUTE',
  'SECOND',
  'WEEKDAY',
  'EOMONTH',
  'DATEDIF',
  'NETWORKDAYS',
  'WORKDAY',
  'PMT',
  'PV',
  'FV',
  'NPV',
  'IRR',
  'RATE',
  'NPER',
  'ROW',
  'COLUMN',
  'ROWS',
  'COLUMNS',
  'TRANSPOSE',
  'UNIQUE',
  'SORT',
  'SORTBY',
  'FILTER',
  'SEQUENCE',
  'RANDARRAY',
  'XMATCH',
  'TEXTSPLIT',
  'TEXTBEFORE',
  'TEXTAFTER',
  'VSTACK',
  'HSTACK',
  'TOROW',
  'TOCOL',
  'WRAPROWS',
  'WRAPCOLS',
  'CHOOSEROWS',
  'CHOOSECOLS',
  'TAKE',
  'DROP',
  'EXPAND',
  'LAMBDA',
  'LET',
  'MAP',
  'REDUCE',
  'SCAN',
  'BYROW',
  'BYCOL',
  'MAKEARRAY',
  'GROUPBY',
  'PIVOTBY',
  'PERCENTOF',
  'IMAGE',
];

/** Find the partial function-name token immediately before `caret` and
 *  return matching candidates. Returns `null` when the caret is not in a
 *  position that warrants a function suggestion (e.g. inside a string).
 *
 *  When `opts.names` is supplied (e.g. from `wb.functionNames()`), it
 *  is preferred over `FUNCTION_NAMES` so engine-registered functions
 *  beyond our hand-curated 98-entry list show up in autocomplete. */
export function suggestFunctions(
  text: string,
  caret: number,
  max = 8,
  opts: { names?: readonly string[] } = {},
): { token: string; tokenStart: number; matches: string[] } | null {
  // Only suggest when we're inside a formula (text starts with '=').
  if (!text.startsWith('=')) return null;
  // Dotted names such as BETA.DIST are one token; decimal literals still fail
  // the leading-letter check below.
  let i = caret - 1;
  while (i >= 0) {
    const ch = text[i] ?? '';
    if (/[A-Za-z0-9_.]/.test(ch)) i -= 1;
    else break;
  }
  const tokenStart = i + 1;
  const token = text.slice(tokenStart, caret);
  if (token.length < 1) return null;
  if (!/^[A-Za-z]/.test(token)) return null;
  const upper = token.toUpperCase();
  const source = opts.names ?? FUNCTION_NAMES;
  const matches = source.filter((n) => n.startsWith(upper)).slice(0, max);
  if (matches.length === 0) return null;
  return { token, tokenStart, matches };
}

/** Hand-authored spreadsheet-style argument lists keyed by upper-cased function
 *  name. `[name]` indicates an optional argument; `...` is repeat-marker. */
export const FUNCTION_SIGNATURES: Readonly<Record<string, readonly string[]>> = {
  SUM: ['number1', '[number2]', '...'],
  AVERAGE: ['number1', '[number2]', '...'],
  COUNT: ['value1', '[value2]', '...'],
  COUNTA: ['value1', '[value2]', '...'],
  COUNTIF: ['range', 'criteria'],
  COUNTIFS: ['criteria_range1', 'criteria1', '...'],
  SUMIF: ['range', 'criteria', '[sum_range]'],
  SUMIFS: ['sum_range', 'criteria_range1', 'criteria1', '...'],
  AVERAGEIF: ['range', 'criteria', '[average_range]'],
  AVERAGEIFS: ['average_range', 'criteria_range1', 'criteria1', '...'],
  MIN: ['number1', '[number2]', '...'],
  MAX: ['number1', '[number2]', '...'],
  MEDIAN: ['number1', '[number2]', '...'],
  IF: ['logical_test', 'value_if_true', '[value_if_false]'],
  IFS: ['logical_test1', 'value1', '...'],
  IFERROR: ['value', 'value_if_error'],
  IFNA: ['value', 'value_if_na'],
  AND: ['logical1', '[logical2]', '...'],
  OR: ['logical1', '[logical2]', '...'],
  NOT: ['logical'],
  XOR: ['logical1', '[logical2]', '...'],
  VLOOKUP: ['lookup_value', 'table_array', 'col_index_num', '[range_lookup]'],
  HLOOKUP: ['lookup_value', 'table_array', 'row_index_num', '[range_lookup]'],
  XLOOKUP: [
    'lookup_value',
    'lookup_array',
    'return_array',
    '[if_not_found]',
    '[match_mode]',
    '[search_mode]',
  ],
  INDEX: ['array', 'row_num', '[col_num]'],
  MATCH: ['lookup_value', 'lookup_array', '[match_type]'],
  OFFSET: ['reference', 'rows', 'cols', '[height]', '[width]'],
  INDIRECT: ['ref_text', '[a1]'],
  CHOOSE: ['index_num', 'value1', '[value2]', '...'],
  ROUND: ['number', 'num_digits'],
  ROUNDUP: ['number', 'num_digits'],
  ROUNDDOWN: ['number', 'num_digits'],
  CEILING: ['number', 'significance'],
  FLOOR: ['number', 'significance'],
  INT: ['number'],
  MOD: ['number', 'divisor'],
  ABS: ['number'],
  ACOS: ['number'],
  POWER: ['number', 'power'],
  SQRT: ['number'],
  EXP: ['number'],
  LN: ['number'],
  LOG: ['number', '[base]'],
  LOG10: ['number'],
  CONCATENATE: ['text1', '[text2]', '...'],
  CONCAT: ['text1', '[text2]', '...'],
  TEXTJOIN: ['delimiter', 'ignore_empty', 'text1', '...'],
  LEFT: ['text', '[num_chars]'],
  RIGHT: ['text', '[num_chars]'],
  MID: ['text', 'start_num', 'num_chars'],
  LEN: ['text'],
  UPPER: ['text'],
  LOWER: ['text'],
  PROPER: ['text'],
  TRIM: ['text'],
  SUBSTITUTE: ['text', 'old_text', 'new_text', '[instance_num]'],
  REPLACE: ['old_text', 'start_num', 'num_chars', 'new_text'],
  FIND: ['find_text', 'within_text', '[start_num]'],
  SEARCH: ['find_text', 'within_text', '[start_num]'],
  TEXT: ['value', 'format_text'],
  VALUE: ['text'],
  NUMBERVALUE: ['text', '[decimal_separator]', '[group_separator]'],
  TODAY: [],
  NOW: [],
  DATE: ['year', 'month', 'day'],
  YEAR: ['serial_number'],
  MONTH: ['serial_number'],
  DAY: ['serial_number'],
  HOUR: ['serial_number'],
  MINUTE: ['serial_number'],
  SECOND: ['serial_number'],
  WEEKDAY: ['serial_number', '[return_type]'],
  EOMONTH: ['start_date', 'months'],
  DATEDIF: ['start_date', 'end_date', 'unit'],
  NETWORKDAYS: ['start_date', 'end_date', '[holidays]'],
  WORKDAY: ['start_date', 'days', '[holidays]'],
  PMT: ['rate', 'nper', 'pv', '[fv]', '[type]'],
  PV: ['rate', 'nper', 'pmt', '[fv]', '[type]'],
  FV: ['rate', 'nper', 'pmt', '[pv]', '[type]'],
  NPV: ['rate', 'value1', '[value2]', '...'],
  IRR: ['values', '[guess]'],
  RATE: ['nper', 'pmt', 'pv', '[fv]', '[type]', '[guess]'],
  NPER: ['rate', 'pmt', 'pv', '[fv]', '[type]'],
  ROW: ['[reference]'],
  COLUMN: ['[reference]'],
  ROWS: ['array'],
  COLUMNS: ['array'],
  TRANSPOSE: ['array'],
  UNIQUE: ['array', '[by_col]', '[exactly_once]'],
  SORT: ['array', '[sort_index]', '[sort_order]', '[by_col]'],
  SORTBY: ['array', 'by_array1', '[sort_order1]', '...'],
  FILTER: ['array', 'include', '[if_empty]'],
  SEQUENCE: ['rows', '[cols]', '[start]', '[step]'],
  RANDARRAY: ['[rows]', '[cols]', '[min]', '[max]', '[whole_number]'],
  XMATCH: ['lookup_value', 'lookup_array', '[match_mode]', '[search_mode]'],
  TEXTSPLIT: [
    'text',
    'col_delimiter',
    '[row_delimiter]',
    '[ignore_empty]',
    '[match_mode]',
    '[pad_with]',
  ],
  TEXTBEFORE: [
    'text',
    'delimiter',
    '[instance_num]',
    '[match_mode]',
    '[match_end]',
    '[if_not_found]',
  ],
  TEXTAFTER: [
    'text',
    'delimiter',
    '[instance_num]',
    '[match_mode]',
    '[match_end]',
    '[if_not_found]',
  ],
  VSTACK: ['array1', '[array2]', '...'],
  HSTACK: ['array1', '[array2]', '...'],
  TOROW: ['array', '[ignore]', '[scan_by_column]'],
  TOCOL: ['array', '[ignore]', '[scan_by_column]'],
  WRAPROWS: ['vector', 'wrap_count', '[pad_with]'],
  WRAPCOLS: ['vector', 'wrap_count', '[pad_with]'],
  CHOOSEROWS: ['array', 'row_num1', '...'],
  CHOOSECOLS: ['array', 'col_num1', '...'],
  TAKE: ['array', 'rows', '[cols]'],
  DROP: ['array', 'rows', '[cols]'],
  EXPAND: ['array', 'rows', '[cols]', '[pad_with]'],
  LAMBDA: ['parameter', '...', 'calculation'],
  LET: ['name1', 'value1', '...', 'calculation'],
  MAP: ['array1', '...', 'lambda'],
  REDUCE: ['initial_value', 'array', 'lambda'],
  SCAN: ['initial_value', 'array', 'lambda'],
  BYROW: ['array', 'lambda'],
  BYCOL: ['array', 'lambda'],
  MAKEARRAY: ['rows', 'cols', 'lambda'],
  GROUPBY: [
    'row_fields',
    'values',
    'function',
    '[field_headers]',
    '[total_depth]',
    '[sort_order]',
    '[filter_array]',
  ],
  PIVOTBY: [
    'row_fields',
    'col_fields',
    'values',
    'function',
    '[field_headers]',
    '[row_total_depth]',
    '[row_sort_order]',
    '[col_total_depth]',
    '[col_sort_order]',
    '[filter_array]',
    '[relative_to]',
  ],
  PERCENTOF: ['data_subset', 'data_all'],
  IMAGE: ['source', '[alt_text]', '[sizing]', '[height]', '[width]'],
};

/** A single `@` implicit-intersection operator surfaced in a formula text.
 *  Desktop spreadsheets prefix existing pre-dynamic-array formulas with `@` to opt
 *  individual references out of array spilling. The UI surfaces these as a
 *  hint while editing so the operator's effect is visible. */
export interface ImplicitIntersection {
  /** Character offset of the `@` itself. */
  at: number;
  /** Source-order index, used for stable tooltip placement. */
  index: number;
}

/** Find every top-level `@` implicit-intersection operator in `formula`. The
 *  scanner skips `@` characters inside string literals and inside structured
 *  refs (`Table[@col]` — there `@` is a ref operator, not the intersection
 *  operator). Returns positions in source order. */
export function findImplicitIntersections(formula: string): ImplicitIntersection[] {
  if (!formula.startsWith('=')) return [];
  const out: ImplicitIntersection[] = [];
  let inString = false;
  let bracketDepth = 0;
  for (let i = 0; i < formula.length; i += 1) {
    const ch = formula[i];
    if (ch === '"') {
      inString = !inString;
      continue;
    }
    if (inString) continue;
    if (ch === '[') {
      bracketDepth += 1;
      continue;
    }
    if (ch === ']') {
      if (bracketDepth > 0) bracketDepth -= 1;
      continue;
    }
    if (ch === '@' && bracketDepth === 0) {
      // the spreadsheet's implicit-intersection `@` only attaches to ref-like targets:
      // a name, a sheet-qualified ref, or an open paren (function call).
      const next = formula[i + 1] ?? '';
      if (!/[A-Za-z_(']/.test(next)) continue;
      out.push({ at: i, index: out.length });
    }
  }
  return out;
}

/** True when `caret` sits immediately after a recognised implicit-intersection
 *  `@` operator (one character or more — handy for showing a hint while the
 *  user is mid-typing the operand). */
export function caretInsideImplicitIntersection(formula: string, caret: number): boolean {
  const ops = findImplicitIntersections(formula);
  for (const op of ops) {
    if (caret > op.at && caret <= op.at + 16) return true;
  }
  return false;
}

/** Resolved signature for the function call enclosing `caret`, or null when
 *  the caret isn't inside a known function. `activeArgIndex` is 0-based and
 *  bumps once per top-level comma between the opening `(` and the caret. */
export interface ActiveSignature {
  name: string;
  args: readonly string[];
  activeArgIndex: number;
}

export function findActiveSignature(text: string, caret: number): ActiveSignature | null {
  // Resolve every syntactic name first so an unknown innermost call masks an
  // enclosing known call, matching the historical helper. The static
  // signature lookup below decides whether that selected call is supported.
  const span = findFunctionCallAtCaret(text, caret, (rawName) => rawName.toUpperCase());
  if (!span) return null;
  // The legacy helper intentionally reports no signature once the caret has
  // moved past a complete top-level call. The span API keeps that boundary
  // eligible for editors that need to select the just-closed call.
  if (span.complete && span.closeParen !== null && caret === span.closeParen + 1) {
    return null;
  }
  const args = FUNCTION_SIGNATURES[span.canonicalName];
  if (!args) return null;
  return {
    name: span.canonicalName,
    args,
    activeArgIndex: span.activeArgumentIndex,
  };
}

const isNameChar = (c: string | undefined): boolean => c != null && /[A-Za-z0-9_.]/.test(c);

/** When `pos` sits on (or directly after) an identifier that is followed by
 *  `(` — i.e. a function name — return that identifier's `[start, end)`
 *  range. Returns null otherwise. */
function functionNameRangeAt(formula: string, pos: number): FormulaSelectionRange | null {
  let start = pos;
  let end = pos;
  while (start > 0 && isNameChar(formula[start - 1])) start -= 1;
  while (end < formula.length && isNameChar(formula[end])) end += 1;
  if (end <= start) return null;
  // A function name starts with a letter or underscore (not a digit/dot).
  if (!/[A-Za-z_]/.test(formula[start] ?? '')) return null;
  let k = end;
  while (k < formula.length && formula[k] === ' ') k += 1;
  if (formula[k] !== '(') return null;
  return { start, end };
}

const trimmedRange = (
  formula: string,
  rawStart: number,
  rawEnd: number,
): FormulaSelectionRange | null => {
  let start = rawStart;
  let end = rawEnd;
  while (start < end && formula[start] === ' ') start += 1;
  while (end > start && formula[end - 1] === ' ') end -= 1;
  return start < end ? { start, end } : null;
};

/** A `[start, end)` slice of a formula string. */
export interface FormulaSelectionRange {
  start: number;
  end: number;
}

/** The source spans belonging to the function call enclosing a formula caret. */
export interface FunctionCallSpan {
  rawName: string;
  canonicalName: string;
  name: FormulaSelectionRange;
  call: FormulaSelectionRange;
  openParen: number;
  closeParen: number | null;
  argumentSpans: readonly FormulaSelectionRange[];
  activeArgumentIndex: number;
  complete: boolean;
}

type FormulaDelimiterKind = 'paren' | 'brace' | 'bracket';

interface FormulaDelimiter {
  kind: FormulaDelimiterKind;
  call: FunctionCallRecord | null;
}

interface FunctionCallRecord {
  rawName: string;
  name: FormulaSelectionRange;
  openParen: number;
  closeParen: number | null;
  separators: number[];
}

const isFormulaIdentifierChar = (value: string | undefined): boolean =>
  value !== undefined && /[A-Za-z0-9_.]/.test(value);

const functionNameBeforeParen = (
  formula: string,
  openParen: number,
): FormulaSelectionRange | null => {
  let nameEnd = openParen;
  while (nameEnd > 0 && /\s/.test(formula[nameEnd - 1] ?? '')) nameEnd -= 1;
  let nameStart = nameEnd;
  while (nameStart > 0 && isFormulaIdentifierChar(formula[nameStart - 1])) nameStart -= 1;
  if (nameStart === nameEnd || !/[A-Za-z_]/.test(formula[nameStart] ?? '')) return null;
  return { start: nameStart, end: nameEnd };
};

const delimiterKindFor = (value: string): FormulaDelimiterKind | null => {
  if (value === '(' || value === ')') return 'paren';
  if (value === '{' || value === '}') return 'brace';
  if (value === '[' || value === ']') return 'bracket';
  return null;
};

const isOpeningDelimiter = (value: string): boolean =>
  value === '(' || value === '{' || value === '[';

const isClosingDelimiter = (value: string): boolean =>
  value === ')' || value === '}' || value === ']';

const delimiterMatches = (kind: FormulaDelimiterKind, value: string): boolean =>
  (kind === 'paren' && value === ')') ||
  (kind === 'brace' && value === '}') ||
  (kind === 'bracket' && value === ']');

const argumentSpansFor = (call: FunctionCallRecord, end: number): FormulaSelectionRange[] => {
  const spans: FormulaSelectionRange[] = [];
  let start = call.openParen + 1;
  for (const separator of call.separators) {
    spans.push({ start, end: separator });
    start = separator + 1;
  }
  spans.push({ start, end });
  return spans;
};

const activeArgumentFor = (separators: readonly number[], caret: number): number => {
  let index = 0;
  for (const separator of separators) {
    if (caret <= separator) break;
    index += 1;
  }
  return index;
};

/**
 * Find the live-known function call containing `caret` without rewriting the
 * formula text. Every returned offset is a JavaScript UTF-16 offset and every
 * range is half-open, so consumers can splice the original source verbatim.
 */
export function findFunctionCallAtCaret(
  text: string,
  caret: number,
  resolveKnownName: (rawName: string) => string | null,
): FunctionCallSpan | null {
  if (!text.startsWith('=') || !Number.isInteger(caret) || caret < 0 || caret > text.length) {
    return null;
  }

  const calls: FunctionCallRecord[] = [];
  const stack: FormulaDelimiter[] = [];
  let inDoubleQuote = false;
  let inSingleQuote = false;

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i] ?? '';
    const next = text[i + 1] ?? '';

    if (inDoubleQuote) {
      if (ch === '"') {
        if (next === '"') i += 1;
        else inDoubleQuote = false;
      }
      continue;
    }

    if (inSingleQuote) {
      if (ch === "'") {
        if (next === "'") i += 1;
        else inSingleQuote = false;
      }
      continue;
    }

    // Structured-reference headers are their own lexical mode. In that mode
    // apostrophe escapes the following bracket/operator character rather than
    // opening a sheet-name quote; this mirrors the formula argument scanner.
    const top = stack[stack.length - 1];
    if (top?.kind === 'bracket') {
      if (ch === '[') {
        stack.push({ kind: 'bracket', call: null });
      } else if (ch === ']') {
        stack.pop();
      } else if (ch === "'" && "[]#'@".includes(next)) {
        i += 1;
      }
      continue;
    }

    if (ch === '"') {
      inDoubleQuote = true;
      continue;
    }
    if (ch === "'") {
      inSingleQuote = true;
      continue;
    }

    if (isOpeningDelimiter(ch)) {
      const kind = delimiterKindFor(ch);
      if (!kind) continue;
      let call: FunctionCallRecord | null = null;
      if (kind === 'paren') {
        const name = functionNameBeforeParen(text, i);
        const inStructuredReference = stack.some((entry) => entry.kind === 'bracket');
        if (name && !inStructuredReference) {
          call = {
            rawName: text.slice(name.start, name.end),
            name,
            openParen: i,
            closeParen: null,
            separators: [],
          };
          calls.push(call);
        }
      }
      stack.push({ kind, call });
      continue;
    }

    if (isClosingDelimiter(ch)) {
      const kind = delimiterKindFor(ch);
      const open = stack[stack.length - 1];
      if (!kind || !open || !delimiterMatches(open.kind, ch)) return null;
      stack.pop();
      if (open.call) open.call.closeParen = i;
      continue;
    }

    if (ch === ',' && top?.kind === 'paren' && top.call) {
      top.call.separators.push(i);
    }
  }

  const knownCalls: FunctionCallSpan[] = [];
  for (const call of calls) {
    let canonicalName: string | null = null;
    try {
      canonicalName = resolveKnownName(call.rawName);
    } catch {
      canonicalName = null;
    }
    if (!canonicalName) continue;
    const closeParen = call.closeParen;
    const callEnd = closeParen === null ? text.length : closeParen + 1;
    knownCalls.push({
      rawName: call.rawName,
      canonicalName,
      name: { ...call.name },
      call: { start: call.name.start, end: callEnd },
      openParen: call.openParen,
      closeParen,
      argumentSpans: argumentSpansFor(call, closeParen ?? text.length),
      activeArgumentIndex: activeArgumentFor(call.separators, caret),
      complete: closeParen !== null,
    });
  }

  const containing = knownCalls.filter((call) => {
    const end = call.closeParen ?? text.length;
    return caret >= call.name.start && caret <= end;
  });
  if (containing.length > 0) {
    containing.sort((a, b) => b.name.start - a.name.start || a.call.end - b.call.end);
    return containing[0] ?? null;
  }

  // A just-closed top-level call remains selectable at its one-past-close
  // boundary. If it is nested, the enclosing call was already selected above.
  const justClosed = knownCalls.filter(
    (call) => call.closeParen !== null && caret === call.closeParen + 1,
  );
  justClosed.sort((a, b) => b.name.start - a.name.start || a.call.end - b.call.end);
  return justClosed[0] ?? null;
}

/**
 * Resolve the selection range for a double-click inside a formula editor at
 * probe offset `probe`:
 *  - the function-name identifier when the probe sits on a name followed by
 *    `(`;
 *  - otherwise the enclosing call argument (comma/paren-delimited at its own
 *    nesting depth, trimmed of surrounding whitespace) — so a range like
 *    `F4:F8` is selected as a whole rather than split at the colon;
 *  - null when the probe is at the top level, so the caller can fall back to
 *    the browser's native word selection.
 */
export function dblClickRange(formula: string, probe: number): FormulaSelectionRange | null {
  const n = formula.length;
  if (n === 0) return null;
  const p = Math.max(0, Math.min(probe, n));

  const fnName = functionNameRangeAt(formula, p);
  if (fnName) return fnName;

  // Single forward pass — tracks string state so quoted commas/parens inside
  // string arguments don't confuse the segment boundaries.
  let inString = false;
  const stack: { sep: number }[] = [];
  let probeDepth = -1;
  let probeSep = -1;
  for (let i = 0; i <= n; i += 1) {
    if (i === p) {
      if (stack.length === 0) return null; // top level
      probeDepth = stack.length;
      probeSep = stack[stack.length - 1]?.sep ?? -1;
    }
    if (i === n) break;
    const ch = formula[i];
    if (ch === '"') {
      inString = !inString;
      continue;
    }
    if (inString) continue;
    if (ch === '(') {
      stack.push({ sep: i });
    } else if (ch === ')') {
      if (probeDepth >= 0 && stack.length === probeDepth) {
        return trimmedRange(formula, probeSep + 1, i);
      }
      stack.pop();
    } else if (ch === ',') {
      if (probeDepth >= 0 && stack.length === probeDepth) {
        return trimmedRange(formula, probeSep + 1, i);
      }
      const top = stack[stack.length - 1];
      if (top) top.sep = i;
    }
  }
  // Unclosed enclosing paren — extend the argument to the end of the formula.
  if (probeDepth >= 0) return trimmedRange(formula, probeSep + 1, n);
  return null;
}
