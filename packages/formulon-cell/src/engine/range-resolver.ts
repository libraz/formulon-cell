import { colFromLetters, MAX_COL, MAX_ROW, parseA1Atom } from './address.js';
import type { Range } from './types.js';
import { formatCell } from './value.js';
import type { WorkbookHandle } from './workbook-handle.js';

/** Resolve a literal string-array list source, or evaluate a range reference
 *  against the workbook to produce one string per non-blank cell. The returned
 *  values are de-duplicated in source order — the spreadsheet's list dropdown does the
 *  same for range-backed lists. Returns `[]` when the ref can't be parsed or
 *  the engine reports no values. */
export type RangeResolver = (ref: string) => string[];

/** Build a resolver bound to a workbook handle and a fallback sheet index
 *  (used when the ref omits a sheet prefix). */
export function makeRangeResolver(wb: WorkbookHandle, fallbackSheet: number): RangeResolver {
  return (ref: string): string[] => resolveRangeRef(wb, ref, fallbackSheet);
}

/** The range `ref` names, on `fallbackSheet` when it has no sheet prefix; null
 *  when it does not parse or names no sheet of the workbook. */
export function rangeForRef(wb: WorkbookHandle, ref: string, fallbackSheet: number): Range | null {
  const parsed = parseRangeRef(ref);
  if (!parsed) return null;
  const sheet = parsed.sheetName === null ? fallbackSheet : sheetIndexByName(wb, parsed.sheetName);
  if (sheet < 0) return null;
  return { sheet, r0: parsed.r0, c0: parsed.c0, r1: parsed.r1, c1: parsed.c1 };
}

export function resolveRangeRef(wb: WorkbookHandle, ref: string, fallbackSheet: number): string[] {
  const range = rangeForRef(wb, ref, fallbackSheet);
  if (!range) return [];
  const { sheet } = range;
  const seen = new Set<string>();
  const out: string[] = [];
  for (let row = range.r0; row <= range.r1; row += 1) {
    for (let col = range.c0; col <= range.c1; col += 1) {
      const value = wb.getValue({ sheet, row, col });
      if (value.kind === 'blank') continue;
      const text = formatCell(value);
      if (!text || seen.has(text)) continue;
      seen.add(text);
      out.push(text);
    }
  }
  return out;
}

export interface ParsedRangeRef {
  sheetName: string | null;
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

function splitSheetPrefix(raw: string): { sheetName: string | null; body: string } | null {
  const trimmed = raw.trim().replace(/^=/, '');
  if (!trimmed) return null;
  if (trimmed.startsWith("'")) {
    // A quoted name may hold "!" and doubles any "'" it contains.
    let i = 1;
    while (i < trimmed.length) {
      if (trimmed[i] !== "'") i += 1;
      else if (trimmed[i + 1] === "'") i += 2;
      else break;
    }
    if (i >= trimmed.length || trimmed[i + 1] !== '!') return null;
    const name = trimmed.slice(1, i).replace(/''/g, "'");
    return name ? { sheetName: name, body: trimmed.slice(i + 2) } : null;
  }
  const bang = trimmed.indexOf('!');
  if (bang === -1) return { sheetName: null, body: trimmed };
  const prefix = trimmed.slice(0, bang);
  if (!prefix) return null;
  return { sheetName: prefix, body: trimmed.slice(bang + 1) };
}

/** Parse `Sheet1!$A$1:$B$5`, `'Sheet 1'!A1:B5`, `A1:B5`, `A1`, `$A$1`. */
export function parseRangeRef(raw: string): ParsedRangeRef | null {
  const split = splitSheetPrefix(raw);
  if (!split) return null;
  const { sheetName, body } = split;
  const parts = body.split(':');
  if (parts.length === 0 || parts.length > 2) return null;
  const head = parseA1Atom(parts[0] ?? '');
  if (!head) return null;
  const tail = parts.length === 2 ? parseA1Atom(parts[1] ?? '') : head;
  if (!tail) return null;
  return {
    sheetName,
    r0: Math.min(head.row, tail.row),
    c0: Math.min(head.col, tail.col),
    r1: Math.max(head.row, tail.row),
    c1: Math.max(head.col, tail.col),
  };
}

const WHOLE_COLUMNS = /^\$?([A-Za-z]+):\$?([A-Za-z]+)$/;
const WHOLE_ROWS = /^\$?(\d+):\$?(\d+)$/;

/** `parseRangeRef` plus whole columns (`B:D`) and whole rows (`2:11`), the
 *  grammar accepted for host-supplied range text. */
export function parseA1RangeText(raw: string): ParsedRangeRef | null {
  const cells = parseRangeRef(raw);
  if (cells) return cells;
  const split = splitSheetPrefix(raw);
  if (!split) return null;
  const { sheetName, body } = split;
  const cols = body.match(WHOLE_COLUMNS);
  if (cols) {
    const a = colFromLetters(cols[1] ?? '');
    const b = colFromLetters(cols[2] ?? '');
    if (a < 0 || b < 0 || a > MAX_COL || b > MAX_COL) return null;
    return { sheetName, r0: 0, c0: Math.min(a, b), r1: MAX_ROW, c1: Math.max(a, b) };
  }
  const rows = body.match(WHOLE_ROWS);
  if (rows) {
    const a = Number.parseInt(rows[1] ?? '', 10) - 1;
    const b = Number.parseInt(rows[2] ?? '', 10) - 1;
    if (a < 0 || b < 0 || a > MAX_ROW || b > MAX_ROW) return null;
    return { sheetName, r0: Math.min(a, b), c0: 0, r1: Math.max(a, b), c1: MAX_COL };
  }
  return null;
}

export function sheetIndexByName(wb: WorkbookHandle, name: string): number {
  const n = wb.sheetCount;
  const target = name.toLowerCase();
  for (let i = 0; i < n; i += 1) {
    if (wb.sheetName(i).toLowerCase() === target) return i;
  }
  return -1;
}

/** True when a list source carries a range reference rather than literal values. */
export function isRangeSource(source: string[] | { ref: string }): source is { ref: string } {
  return !Array.isArray(source);
}

/** Resolve a range ref to its numeric values in source order. Non-numeric
 *  cells (blank, text, bool, error) are skipped — sparklines plot a numeric
 *  series, mirroring the spreadsheet behaviour. */
export function resolveNumericRange(
  wb: WorkbookHandle,
  ref: string,
  fallbackSheet: number,
): number[] {
  const parsed = parseRangeRef(ref);
  if (!parsed) return [];
  let sheet = fallbackSheet;
  if (parsed.sheetName !== null) {
    const idx = sheetIndexByName(wb, parsed.sheetName);
    if (idx < 0) return [];
    sheet = idx;
  }
  const out: number[] = [];
  for (let row = parsed.r0; row <= parsed.r1; row += 1) {
    for (let col = parsed.c0; col <= parsed.c1; col += 1) {
      const value = wb.getValue({ sheet, row, col });
      if (value.kind === 'number' && Number.isFinite(value.value)) out.push(value.value);
    }
  }
  return out;
}

/** Engine-free numeric resolver — pulls from a `data.cells` map keyed by
 *  `${sheet}:${row}:${col}`. Used by the renderer when no workbook handle is
 *  attached and by tests that exercise the painter in isolation. Refs that
 *  carry a sheet name are skipped unless `sheetByName` resolves them. */
export function resolveNumericRangeFromCells(
  cells: ReadonlyMap<string, { value: { kind: string; value?: unknown } }>,
  ref: string,
  fallbackSheet: number,
  sheetByName?: (name: string) => number,
): number[] {
  const parsed = parseRangeRef(ref);
  if (!parsed) return [];
  let sheet = fallbackSheet;
  if (parsed.sheetName !== null) {
    if (!sheetByName) return [];
    const idx = sheetByName(parsed.sheetName);
    if (idx < 0) return [];
    sheet = idx;
  }
  const out: number[] = [];
  for (let row = parsed.r0; row <= parsed.r1; row += 1) {
    for (let col = parsed.c0; col <= parsed.c1; col += 1) {
      const cell = cells.get(`${sheet}:${row}:${col}`);
      if (!cell) continue;
      const v = cell.value;
      if (v.kind === 'number' && typeof v.value === 'number' && Number.isFinite(v.value)) {
        out.push(v.value);
      }
    }
  }
  return out;
}
