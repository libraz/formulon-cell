import { addrKey } from '../engine/address.js';
import type { CellValue, EvalArrayResult, EvalResult } from '../engine/types.js';
import { fromEngineValue } from '../engine/value.js';

/**
 * the spreadsheet's F9 key, while editing a formula, replaces the highlighted
 * sub-expression with its evaluated value. The baseline path handles the two
 * common cases that don't need an engine call:
 *
 *   - A selection that is a numeric or text literal — return the literal.
 *   - A selection that is a single A1-style reference (optionally
 *     sheet-prefixed) — look the cell up in the supplied cell map and
 *     return its current value.
 *
 * When the host supplies an evaluator, we first ask it to evaluate the
 * selected expression as-is. Modern engines use the active cell context for
 * names, sheet-qualified refs, and range anchors. If that fails, we fall back
 * to substituting cell references with literals and evaluating the resulting
 * self-contained expression.
 */
export interface F9Preview {
  /** Resolved display string (`"3.14"`, `"Hello"`, `"true"`, `"#REF!"`). */
  display: string;
  /** True when the caller can safely substitute `display` into the formula
   *  in place of the original selection. Falls back to false for partial
   *  evaluations (refs that the cell map doesn't carry, complex
   *  sub-expressions, etc.). */
  substitutable: boolean;
}

export interface F9Replacement {
  text: string;
  start: number;
  end: number;
  preview: F9Preview;
}

const REF_RE = /^(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_]*))?!?(\$?)([A-Za-z]+)(\$?)(\d+)$/;
const NUMBER_RE = /^-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?$/;
const STRING_RE = /^"([^"]*)"$/;
type F9FormulaEvaluator = (formula: string) => EvalResult;
type F9ArrayEvaluator = (formula: string) => EvalArrayResult;
const MAX_RANGE_EXPANSION_CELLS = 10_000;
/** Upper bound on cells rendered into an F9 array constant. Beyond this the
 *  preview collapses to the array's top-left element rather than pasting an
 *  unwieldy literal into the formula bar. */
const MAX_ARRAY_PREVIEW_CELLS = 4_096;

/** Render an `evaluateFormulaArray` result as a spreadsheet-style array
 *  constant (`{a,b;c,d}` — columns by `,`, rows by `;`). A 1x1 result renders as the
 *  bare scalar. Returns `null` when the result is too large to inline. */
function renderArrayCell(v: EvalArrayResult['cells'][number][number] | undefined): string {
  return renderCellValueForF9(v === undefined ? { kind: 'blank' } : fromEngineValue(v));
}

function renderArrayForF9(result: EvalArrayResult): string | null {
  const { rows, cols, cells } = result;
  if (rows <= 0 || cols <= 0) return null;
  if (rows === 1 && cols === 1) {
    return renderArrayCell(cells[0]?.[0]);
  }
  if (rows * cols > MAX_ARRAY_PREVIEW_CELLS) return null;
  const rowLiterals: string[] = [];
  for (let r = 0; r < rows; r += 1) {
    const row = cells[r] ?? [];
    const colLiterals: string[] = [];
    for (let c = 0; c < cols; c += 1) {
      colLiterals.push(renderArrayCell(row[c]));
    }
    rowLiterals.push(colLiterals.join(','));
  }
  return `{${rowLiterals.join(';')}}`;
}

const lettersToCol = (letters: string): number => {
  let col = 0;
  for (let i = 0; i < letters.length; i += 1) {
    col = col * 26 + (letters.toUpperCase().charCodeAt(i) - 64);
  }
  return col - 1;
};

/** Render a CellValue the way the formula bar would substitute it after F9. */
export function renderCellValueForF9(v: CellValue | undefined): string {
  if (!v || v.kind === 'blank') return '0';
  if (v.kind === 'number') return String(v.value);
  if (v.kind === 'text') return `"${v.value}"`;
  if (v.kind === 'bool') return v.value ? 'TRUE' : 'FALSE';
  if (v.kind === 'error') return v.text || '#ERROR!';
  return '';
}

function cellValueLiteral(v: CellValue | undefined): string {
  if (!v || v.kind === 'blank') return '0';
  if (v.kind === 'number') return String(v.value);
  if (v.kind === 'bool') return v.value ? 'TRUE' : 'FALSE';
  if (v.kind === 'text') return JSON.stringify(v.value);
  if (v.kind === 'error') return v.text || '#ERROR!';
  return '0';
}

/** Compute the F9 substitution for `selection` taken from `formula`. The
 *  selection is the substring the user has highlighted while editing. The
 *  `cells` map mirrors `DataSlice.cells` and `sheetByName` translates a
 *  sheet name (spreadsheet-side) to its 0-based index — when omitted, sheet-
 *  qualified refs are unresolved. */
export function computeF9Preview(
  _formula: string,
  selection: string,
  activeSheet: number,
  cells: ReadonlyMap<string, { value: CellValue; formula: string | null }>,
  sheetByName?: (name: string) => number,
  evalFormula?: F9FormulaEvaluator,
  preferContextualEvaluation = false,
  evalFormulaArray?: F9ArrayEvaluator,
): F9Preview {
  const trimmed = selection.trim();
  if (!trimmed) {
    return { display: '', substitutable: false };
  }
  if (NUMBER_RE.test(trimmed)) {
    return { display: trimmed, substitutable: true };
  }
  if (STRING_RE.test(trimmed)) {
    return { display: trimmed, substitutable: true };
  }
  if (/^(true|false)$/i.test(trimmed)) {
    return { display: trimmed.toUpperCase(), substitutable: true };
  }
  const ref = trimmed.match(REF_RE);
  if (ref) {
    const sheetName = ref[1] ?? ref[2];
    const letters = ref[4] ?? '';
    const digits = ref[6] ?? '';
    const col = lettersToCol(letters);
    const row = Number.parseInt(digits, 10) - 1;
    if (row < 0 || col < 0) {
      return { display: '#REF!', substitutable: false };
    }
    let sheet = activeSheet;
    if (sheetName) {
      const resolved = sheetByName?.(sheetName);
      if (resolved === undefined || resolved < 0) {
        return { display: '#REF!', substitutable: false };
      }
      sheet = resolved;
    }
    const cell = cells.get(addrKey({ sheet, row, col }));
    return { display: renderCellValueForF9(cell?.value), substitutable: true };
  }
  if (evalFormulaArray && preferContextualEvaluation) {
    const arr = evalFormulaArray(trimmed.startsWith('=') ? trimmed : `=${trimmed}`);
    if (arr.status.status === 0) {
      const display = renderArrayForF9(arr);
      if (display !== null) return { display, substitutable: true };
    }
  }
  if (evalFormula && preferContextualEvaluation) {
    const contextual = evalFormula(trimmed.startsWith('=') ? trimmed : `=${trimmed}`);
    if (contextual.status.status === 0) {
      return {
        display: renderCellValueForF9(fromEngineValue(contextual.value)),
        substitutable: true,
      };
    }
  }
  const literalFormula = evalFormula
    ? substituteSingleCellRefs(trimmed, activeSheet, cells, sheetByName)
    : null;
  if (evalFormula && literalFormula) {
    const res = evalFormula(literalFormula);
    if (res.status.status === 0) {
      return { display: renderCellValueForF9(fromEngineValue(res.value)), substitutable: true };
    }
  }
  return { display: '', substitutable: false };
}

function substituteSingleCellRefs(
  expression: string,
  activeSheet: number,
  cells: ReadonlyMap<string, { value: CellValue; formula: string | null }>,
  sheetByName?: (name: string) => number,
): string | null {
  const formula = expression.startsWith('=') ? expression : `=${expression}`;
  const refs = extractF9Refs(formula);
  if (refs === null) return null;
  if (refs.length === 0) return expression;
  let out = formula.slice(1);
  for (const ref of refs.sort((a, b) => b.start - a.start)) {
    let sheet = activeSheet;
    if (ref.sheetName) {
      const resolved = sheetByName?.(ref.sheetName);
      if (resolved === undefined || resolved < 0) return null;
      sheet = resolved;
    }
    const replacement =
      ref.kind === 'range'
        ? rangeLiteralList(sheet, ref, cells)
        : cellValueLiteral(cells.get(addrKey({ sheet, row: ref.row, col: ref.col }))?.value);
    if (replacement === null) return null;
    out = `${out.slice(0, ref.start - 1)}${replacement}${out.slice(ref.end - 1)}`;
  }
  return out;
}

type ExtractedF9Ref =
  | {
      kind: 'cell';
      start: number;
      end: number;
      sheetName: string | null;
      row: number;
      col: number;
    }
  | {
      kind: 'range';
      start: number;
      end: number;
      sheetName: string | null;
      row: number;
      col: number;
      row2: number;
      col2: number;
    };

function rangeLiteralList(
  sheet: number,
  ref: ExtractedF9Ref & { kind: 'range' },
  cells: ReadonlyMap<string, { value: CellValue; formula: string | null }>,
): string | null {
  const r0 = Math.min(ref.row, ref.row2);
  const r1 = Math.max(ref.row, ref.row2);
  const c0 = Math.min(ref.col, ref.col2);
  const c1 = Math.max(ref.col, ref.col2);
  const count = (r1 - r0 + 1) * (c1 - c0 + 1);
  if (count > MAX_RANGE_EXPANSION_CELLS) return null;
  const values: string[] = [];
  for (let row = r0; row <= r1; row += 1) {
    for (let col = c0; col <= c1; col += 1) {
      values.push(cellValueLiteral(cells.get(addrKey({ sheet, row, col }))?.value));
    }
  }
  return values.join(',');
}

function extractF9Refs(formula: string): ExtractedF9Ref[] | null {
  const re =
    /(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_]*))?!?(\$?[A-Za-z]+\$?\d+)(?::(\$?[A-Za-z]+\$?\d+))?/g;
  const refs: ExtractedF9Ref[] = [];
  for (let m = re.exec(formula); m !== null; m = re.exec(formula)) {
    const before = formula.slice(0, m.index);
    const quoteCount = (before.match(/"/g) ?? []).length;
    if (quoteCount % 2 === 1) continue;
    if (formula[m.index + m[0].length] === '(') continue;
    const parsed = parseCellRef(m[3] ?? '');
    if (!parsed) return null;
    const end = m.index + m[0].length;
    const base = {
      start: m.index,
      end,
      sheetName: m[1] ?? m[2] ?? null,
      row: parsed.row,
      col: parsed.col,
    };
    if (!m[4]) {
      refs.push({ kind: 'cell', ...base });
      continue;
    }
    const parsedEnd = parseCellRef(m[4]);
    if (!parsedEnd) return null;
    refs.push({
      kind: 'range',
      ...base,
      row2: parsedEnd.row,
      col2: parsedEnd.col,
    });
  }
  return refs;
}

function parseCellRef(raw: string): { row: number; col: number } | null {
  const m = raw.match(/^\$?([A-Za-z]+)\$?(\d+)$/);
  if (!m) return null;
  const col = lettersToCol(m[1] ?? '');
  const row = Number.parseInt(m[2] ?? '', 10) - 1;
  if (row < 0 || col < 0 || row > 1048575 || col > 16383) return null;
  return { row, col };
}

export function replaceFormulaSelectionWithF9Preview(
  formula: string,
  start: number,
  end: number,
  activeSheet: number,
  cells: ReadonlyMap<string, { value: CellValue; formula: string | null }>,
  sheetByName?: (name: string) => number,
  evalFormula?: F9FormulaEvaluator,
  preferContextualEvaluation = false,
  evalFormulaArray?: F9ArrayEvaluator,
): F9Replacement | null {
  if (!formula.startsWith('=') || start === end) return null;
  const left = Math.max(0, Math.min(start, end));
  const right = Math.min(formula.length, Math.max(start, end));
  const selection = formula.slice(left, right);
  const preview = computeF9Preview(
    formula,
    selection,
    activeSheet,
    cells,
    sheetByName,
    evalFormula,
    preferContextualEvaluation,
    evalFormulaArray,
  );
  if (!preview.substitutable) return { text: formula, start: left, end: right, preview };
  const text = `${formula.slice(0, left)}${preview.display}${formula.slice(right)}`;
  const caret = left + preview.display.length;
  return { text, start: caret, end: caret, preview };
}
