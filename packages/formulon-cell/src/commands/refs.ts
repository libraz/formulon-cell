import { colFromLetters, formatA1Cell, MAX_COL, MAX_ROW, parseA1Atom } from '../engine/address.js';
import { formulaCaretContext, formulaIdentifierPrefix } from './formula-edit-context.js';
import { FUNCTION_NAMES } from './formula-function-names.js';
import { rewriteRefs } from './formula-ref-lexer.js';
import { type R1C1Axis, type R1C1Endpoint, scanR1C1References } from './r1c1-reference.js';

/**
 * F4 reference rotation: cycles between A1, $A$1, A$1, $A1, then back to A1.
 * Operates on the cell reference at `caret` in `text`. Returns the new text
 * + new caret position. If no reference is found, returns the input unchanged.
 */
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

export interface FormulaReferenceFormatOptions {
  r1c1?: boolean;
  base?: { row: number; col: number };
}

type FormulaRefCandidate = Omit<FormulaRef, 'colorIndex'>;

function validR1C1Base(
  base: { row: number; col: number } | undefined,
): base is { row: number; col: number } {
  return (
    base !== undefined &&
    Number.isInteger(base.row) &&
    Number.isInteger(base.col) &&
    base.row >= 0 &&
    base.row <= MAX_ROW &&
    base.col >= 0 &&
    base.col <= MAX_COL
  );
}

function resolveR1C1Axis(axis: R1C1Axis, base: number | undefined): number | null {
  if (axis.absolute) return axis.value - 1;
  return base === undefined ? null : base + axis.value;
}

function resolveR1C1Endpoint(
  endpoint: R1C1Endpoint,
  base: { row: number; col: number } | undefined,
): { row: number; col: number } | null {
  if ((!endpoint.row.absolute || !endpoint.column.absolute) && !validR1C1Base(base)) return null;
  const row = resolveR1C1Axis(endpoint.row, base?.row);
  const col = resolveR1C1Axis(endpoint.column, base?.col);
  if (row === null || col === null || row < 0 || col < 0 || row > MAX_ROW || col > MAX_COL)
    return null;
  return { row, col };
}

function mergeColonConnectedCandidates(
  text: string,
  candidates: FormulaRefCandidate[],
): FormulaRefCandidate[] {
  const merged: FormulaRefCandidate[] = [];
  for (const candidate of candidates) {
    const previous = merged[merged.length - 1];
    if (previous && /^\s*:\s*$/u.test(text.slice(previous.end, candidate.start))) {
      previous.r0 = Math.min(previous.r0, candidate.r0);
      previous.c0 = Math.min(previous.c0, candidate.c0);
      previous.r1 = Math.max(previous.r1, candidate.r1);
      previous.c1 = Math.max(previous.c1, candidate.c1);
      previous.end = candidate.end;
    } else {
      merged.push({ ...candidate });
    }
  }
  return merged;
}

function assignReferenceColors(candidates: FormulaRefCandidate[]): FormulaRef[] {
  const colorMap = new Map<string, number>();
  return candidates.map((ref) => {
    const key = `${ref.r0}:${ref.c0}:${ref.r1}:${ref.c1}`;
    let colorIndex = colorMap.get(key);
    if (colorIndex === undefined) {
      colorIndex = colorMap.size;
      colorMap.set(key, colorIndex);
    }
    return { ...ref, colorIndex };
  });
}

function a1ReferenceCandidates(text: string): FormulaRefCandidate[] {
  const candidates: FormulaRefCandidate[] = [];
  rewriteRefs(text, (token, start) => {
    if (token.kind !== 'cell') return null;
    const headStart = start + token.sheetQual.length;
    const headEnd =
      headStart +
      token.a.label.length +
      token.a.rowStr.length +
      (token.a.absCol ? 1 : 0) +
      (token.a.absRow ? 1 : 0);
    const headM = text.slice(headStart, headEnd);
    const head = parseA1Atom(headM);
    const tailM = token.b
      ? text.slice(start + token.sheetQual.length + headM.length + 1, token.end)
      : undefined;
    const tail = tailM ? parseA1Atom(tailM) : head;
    if (!head || !tail) return null;
    candidates.push({
      r0: Math.min(head.row, tail.row),
      c0: Math.min(head.col, tail.col),
      r1: Math.max(head.row, tail.row),
      c1: Math.max(head.col, tail.col),
      start,
      end: token.end,
    });
    return null;
  });
  return mergeColonConnectedCandidates(text, candidates);
}

function r1c1ReferenceCandidates(
  text: string,
  base: { row: number; col: number } | undefined,
): FormulaRefCandidate[] {
  const candidates: FormulaRefCandidate[] = [];
  for (const reference of scanR1C1References(text)) {
    const endpoints: { row: number; col: number }[] = [];
    let isValid = true;
    for (const endpoint of reference.endpoints) {
      const coordinate = resolveR1C1Endpoint(endpoint, base);
      if (!coordinate) {
        isValid = false;
        break;
      }
      endpoints.push(coordinate);
    }
    if (!isValid || endpoints.length === 0) continue;
    const rows = endpoints.map(({ row }) => row);
    const cols = endpoints.map(({ col }) => col);
    candidates.push({
      r0: Math.min(...rows),
      c0: Math.min(...cols),
      r1: Math.max(...rows),
      c1: Math.max(...cols),
      start: reference.start,
      end: reference.end,
    });
  }
  return candidates;
}

/** Extract every cell or range reference from a formula text. The returned
 *  list is in source order with a stable per-target color index so callers
 *  can paint distinct highlights for each ref, the same way spreadsheets do
 *  while editing a formula. */
export function extractRefs(
  text: string,
  options: FormulaReferenceFormatOptions = {},
): FormulaRef[] {
  if (!text.startsWith('=')) return [];
  const candidates = options.r1c1
    ? r1c1ReferenceCandidates(text, options.base)
    : a1ReferenceCandidates(text);
  return assignReferenceColors(candidates);
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

export function rotateRefAt(
  text: string,
  caret: number,
  options: FormulaReferenceFormatOptions = {},
): F4Result {
  if (options.r1c1) return rotateR1C1ReferenceAt(text, caret, options.base);
  let rotated: F4Result | null = null;
  rewriteRefs(text, (token, tokenStart) => {
    if (rotated || token.kind !== 'cell') return null;
    const firstStart = tokenStart + token.sheetQual.length;
    const firstEnd =
      firstStart +
      token.a.label.length +
      token.a.rowStr.length +
      (token.a.absCol ? 1 : 0) +
      (token.a.absRow ? 1 : 0);
    const secondStart = firstEnd + 1;
    const secondEnd = token.end;
    const endpoint =
      caret >= secondStart && caret <= secondEnd && token.b
        ? { start: secondStart, end: secondEnd }
        : caret >= firstStart && caret <= firstEnd
          ? { start: firstStart, end: firstEnd }
          : null;
    if (!endpoint) return null;
    const raw = text.slice(endpoint.start, endpoint.end);
    const match = /^(\$?)([A-Za-z]+)(\$?)(\d+)$/.exec(raw);
    if (!match) return null;
    const [, dCol, letters, dRow, digits] = match;
    const next = nextStep(dCol === '$', dRow === '$');
    const replacement = `${next.col ? '$' : ''}${letters}${next.row ? '$' : ''}${digits}`;
    rotated = {
      text: `${text.slice(0, endpoint.start)}${replacement}${text.slice(endpoint.end)}`,
      caret: endpoint.start + replacement.length,
    };
    return null;
  });
  return rotated ?? { text, caret };
}

function formatR1C1Axis(
  prefix: 'R' | 'C',
  absolute: boolean,
  coordinate: number,
  base: number,
): string {
  if (absolute) return `${prefix}${coordinate + 1}`;
  const delta = coordinate - base;
  return delta === 0 ? prefix : `${prefix}[${delta}]`;
}

function rotateR1C1ReferenceAt(
  text: string,
  caret: number,
  base: { row: number; col: number } | undefined,
): F4Result {
  if (!validR1C1Base(base)) return { text, caret };
  for (const reference of scanR1C1References(text)) {
    const endpoint = reference.endpoints.find(
      (candidate) => caret >= candidate.start && caret <= candidate.end,
    );
    if (!endpoint) continue;
    const coordinate = resolveR1C1Endpoint(endpoint, base);
    if (!coordinate) return { text, caret };
    const next = nextStep(endpoint.column.absolute, endpoint.row.absolute);
    const replacement =
      `${formatR1C1Axis('R', next.row, coordinate.row, base.row)}` +
      `${formatR1C1Axis('C', next.col, coordinate.col, base.col)}`;
    return {
      text: `${text.slice(0, endpoint.start)}${replacement}${text.slice(endpoint.end)}`,
      caret: endpoint.start + replacement.length,
    };
  }
  return { text, caret };
}

function nextStep(absCol: boolean, absRow: boolean): { col: boolean; row: boolean } {
  // Spreadsheet order: A1 -> $A$1 -> A$1 -> $A1 -> A1
  if (!absCol && !absRow) return { col: true, row: true };
  if (absCol && absRow) return { col: false, row: true };
  if (!absCol && absRow) return { col: true, row: false };
  return { col: false, row: false };
}

/** Resolve one R1C1 atom against `base` into A1 text (`$B$5`), or `#REF!`. */
export const r1c1AtomToA1 = (raw: string, base: { row: number; col: number }): string => {
  const reference = scanR1C1References(raw).find(
    (candidate) => candidate.start === 0 && candidate.end === raw.length && !candidate.qualifier,
  );
  const endpoint = reference?.endpoints[0];
  if (reference?.endpoints.length !== 1 || !endpoint) return raw;
  const coordinate = resolveR1C1Endpoint(endpoint, base);
  if (!coordinate) return '#REF!';
  return formatResolvedA1Endpoint(endpoint, coordinate);
};

function formatResolvedA1Endpoint(
  endpoint: R1C1Endpoint,
  coordinate: { row: number; col: number },
): string {
  const a1 = formatA1Cell(coordinate.row, coordinate.col);
  const parts = /^([A-Z]+)(\d+)$/.exec(a1);
  if (!parts) return a1;
  return `${endpoint.column.absolute ? '$' : ''}${parts[1]}${endpoint.row.absolute ? '$' : ''}${parts[2]}`;
}

/** Convert R1C1 references in a user-authored formula into A1 references for
 *  the workbook engine. This keeps the public R1C1 editing surface while
 *  preserving the existing A1 formula backend. */
export function normalizeR1C1Formula(formula: string, base: { row: number; col: number }): string {
  if (!formula.startsWith('=')) return formula;
  let output = '';
  let cursor = 0;
  for (const reference of scanR1C1References(formula, { includeExternal: true })) {
    const converted = reference.endpoints.map((endpoint) => {
      const coordinate = resolveR1C1Endpoint(endpoint, base);
      return coordinate ? formatResolvedA1Endpoint(endpoint, coordinate) : '#REF!';
    });
    output += formula.slice(cursor, reference.start);
    let endpointCursor = reference.start;
    for (let index = 0; index < reference.endpoints.length; index += 1) {
      const endpoint = reference.endpoints[index];
      const formatted = converted[index];
      if (!endpoint || formatted === undefined) continue;
      output += formula.slice(endpointCursor, endpoint.start);
      output += formatted;
      endpointCursor = endpoint.end;
    }
    output += formula.slice(endpointCursor, reference.end);
    cursor = reference.end;
  }
  return `${output}${formula.slice(cursor)}`;
}

const a1AtomToR1C1 = (raw: string, base: { row: number; col: number }): string => {
  if (!validR1C1Base(base)) return raw;
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
  return rewriteRefs(formula, (token, start) => {
    if (token.kind !== 'cell') return formula.slice(start, token.end);
    const headStart = start + token.sheetQual.length;
    const headEnd =
      headStart +
      token.a.label.length +
      token.a.rowStr.length +
      (token.a.absCol ? 1 : 0) +
      (token.a.absRow ? 1 : 0);
    const head = a1AtomToR1C1(formula.slice(headStart, headEnd), base);
    if (!token.b) return `${formula.slice(start, headStart)}${head}`;
    const tailStart = headEnd + 1;
    const tail = a1AtomToR1C1(formula.slice(tailStart, token.end), base);
    return `${formula.slice(start, headStart)}${head}${formula.slice(headEnd, tailStart)}${tail}`;
  });
}

// Keep this export path stable while sharing the catalog with the formula lexer.
export { FUNCTION_NAMES } from './formula-function-names.js';
// `shiftFormulaRefs` (relative-offset shifting for fill/paste/ctrl-enter) now
// lives in the shared reference-rewriting module alongside the structural-edit
// transforms, so all three share one tokenizer (function-name, sheet-qualifier,
// range, and grid-bound handling). Re-exported here for existing import sites.
export { shiftFormulaRefs } from './formula-refs.js';

/** Formula-text options shared by the argument scanner and editor helpers. */
export interface FormulaScanOptions {
  /** Active workbook list separator. Excel uses `;` for the German/French
   * profiles and `,` for the canonical and other supported profiles. */
  listSeparator?: string;
  /** Resolve a user-facing function alias to its canonical catalog name. */
  resolveKnownName?: (rawName: string) => string | null;
}

const formulaListSeparator = (value: string | undefined): ',' | ';' => (value === ';' ? ';' : ',');

const isUnicodeIdentifierChar = (value: string | undefined): boolean =>
  value !== undefined && /[\p{L}\p{N}_.]/u.test(value);

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
  // Only suggest in formula expressions and at a function name immediately
  // before its opening parenthesis.
  const caretContext = formulaCaretContext(text, caret);
  if (!text.startsWith('=') || (caretContext !== 'expression' && caretContext !== 'function'))
    return null;
  // Same token boundary as custom functions and defined names, so a candidate
  // cannot match only the suffix after a dot or backslash. Dotted names such as
  // BETA.DIST are one token; decimal literals still fail the leading-letter check below.
  const { token, tokenStart } = formulaIdentifierPrefix(text, caret);
  if (token.length < 1) return null;
  if (!/^[\p{L}_]/u.test(token)) return null;
  const upper = token.toLocaleUpperCase();
  const source = opts.names ?? FUNCTION_NAMES;
  const matches = source.filter((n) => n.toLocaleUpperCase().startsWith(upper)).slice(0, max);
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

export function findActiveSignature(
  text: string,
  caret: number,
  options: FormulaScanOptions & {
    /** Argument labels for a host-registered function without a built-in signature. */
    getCustomFunctionArgs?: (name: string) => readonly string[] | undefined;
  } = {},
): ActiveSignature | null {
  // Resolve every syntactic name first so an unknown innermost call masks an
  // enclosing known call, matching the historical helper. The static
  // signature lookup below decides whether that selected call is supported.
  const span = findFunctionCallAtCaret(
    text,
    caret,
    options.resolveKnownName ?? ((rawName) => rawName.toUpperCase()),
    options.listSeparator,
  );
  if (!span) return null;
  // The legacy helper intentionally reports no signature once the caret has
  // moved past a complete top-level call. The span API keeps that boundary
  // eligible for editors that need to select the just-closed call.
  if (span.complete && span.closeParen !== null && caret === span.closeParen + 1) {
    return null;
  }
  const args =
    FUNCTION_SIGNATURES[span.canonicalName] ?? options.getCustomFunctionArgs?.(span.canonicalName);
  if (!args) return null;
  return {
    name: span.canonicalName,
    args,
    activeArgIndex: span.activeArgumentIndex,
  };
}

const isNameChar = (c: string | undefined): boolean => isUnicodeIdentifierChar(c);

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
  if (!/[\p{L}_]/u.test(formula[start] ?? '')) return null;
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
  /** Whether an EOF-incomplete call can be safely closed by an editor. */
  safeToComplete: boolean;
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
  isUnicodeIdentifierChar(value);

const functionNameBeforeParen = (
  formula: string,
  openParen: number,
): FormulaSelectionRange | null => {
  let nameEnd = openParen;
  while (nameEnd > 0 && /\s/.test(formula[nameEnd - 1] ?? '')) nameEnd -= 1;
  let nameStart = nameEnd;
  while (nameStart > 0 && isFormulaIdentifierChar(formula[nameStart - 1])) nameStart -= 1;
  if (nameStart === nameEnd || !/[\p{L}_]/u.test(formula[nameStart] ?? '')) return null;
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
  listSeparator?: string,
): FunctionCallSpan | null {
  if (!text.startsWith('=') || !Number.isInteger(caret) || caret < 0 || caret > text.length) {
    return null;
  }

  const calls: FunctionCallRecord[] = [];
  const stack: FormulaDelimiter[] = [];
  const separator = formulaListSeparator(listSeparator);
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

    if (ch === separator && top?.kind === 'paren' && top.call) {
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
      safeToComplete:
        closeParen !== null ||
        (!inDoubleQuote && !inSingleQuote && stack[stack.length - 1]?.call === call),
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
export function dblClickRange(
  formula: string,
  probe: number,
  listSeparator?: string,
): FormulaSelectionRange | null {
  const n = formula.length;
  if (n === 0) return null;
  const p = Math.max(0, Math.min(probe, n));

  const fnName = functionNameRangeAt(formula, p);
  if (fnName) return fnName;

  // Single forward pass — tracks strings and every nested delimiter so
  // separators inside quoted, structured, or array arguments don't become
  // argument boundaries.
  let quote: '"' | "'" | null = null;
  const separator = formulaListSeparator(listSeparator);
  const stack: { kind: 'paren' | 'brace' | 'bracket'; sep: number }[] = [];
  let probeCallIndex = -1;
  let probeSep = -1;
  for (let i = 0; i <= n; i += 1) {
    if (i === p) {
      probeCallIndex = -1;
      for (let j = stack.length - 1; j >= 0; j -= 1) {
        if (stack[j]?.kind === 'paren') {
          probeCallIndex = j;
          break;
        }
      }
      if (probeCallIndex < 0) return null; // top level
      probeSep = stack[probeCallIndex]?.sep ?? -1;
    }
    if (i === n) break;
    const ch = formula[i] ?? '';
    const top = stack[stack.length - 1];
    if (top?.kind === 'bracket') {
      if (ch === '[') {
        stack.push({ kind: 'bracket', sep: i });
        continue;
      }
      if (ch === ']') {
        stack.pop();
        continue;
      }
      // Structured-reference headers use apostrophes to escape [, ], #, @,
      // and apostrophe itself. This is distinct from a quoted sheet name.
      if (ch === "'" && "[]#'@".includes(formula[i + 1] ?? '')) i += 1;
      continue;
    }
    if (quote) {
      if (ch === quote && formula[i + 1] === quote) {
        i += 1;
      } else if (ch === quote) {
        quote = null;
      }
      continue;
    }
    if (ch === '"' || ch === "'") {
      quote = ch;
      continue;
    }
    if (ch === '(' || ch === '{' || ch === '[') {
      stack.push({ kind: ch === '(' ? 'paren' : ch === '{' ? 'brace' : 'bracket', sep: i });
    } else if (ch === ')' || ch === '}' || ch === ']') {
      const kind = ch === ')' ? 'paren' : ch === '}' ? 'brace' : 'bracket';
      if (probeCallIndex >= 0 && stack.length - 1 === probeCallIndex && kind === 'paren') {
        return trimmedRange(formula, probeSep + 1, i);
      }
      stack.pop();
    } else if (ch === separator) {
      if (probeCallIndex >= 0 && stack.length - 1 === probeCallIndex) {
        return trimmedRange(formula, probeSep + 1, i);
      }
      const top = stack[stack.length - 1];
      if (top?.kind === 'paren') top.sep = i;
    }
  }
  // Unclosed enclosing paren — extend the argument to the end of the formula.
  if (probeCallIndex >= 0) return trimmedRange(formula, probeSep + 1, n);
  return null;
}
