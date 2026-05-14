import type { WorkbookHandle } from '../engine/workbook-handle.js';

export function colName(col: number): string {
  let n = col;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

function cellRef(row: number, col: number, r1c1: boolean): string {
  return r1c1 ? `R${row + 1}C${col + 1}` : `${colName(col)}${row + 1}`;
}

export function formatSelectionRef(
  range: { r0: number; c0: number; r1: number; c1: number },
  active: { row: number; col: number },
  r1c1: boolean,
): string {
  if (range.r0 === range.r1 && range.c0 === range.c1) {
    return cellRef(active.row, active.col, r1c1);
  }
  if (!r1c1) {
    if (range.r0 === 0 && range.r1 === 1048575) {
      return range.c0 === range.c1
        ? colName(range.c0)
        : `${colName(range.c0)}:${colName(range.c1)}`;
    }
    if (range.c0 === 0 && range.c1 === 16383) {
      return range.r0 === range.r1 ? `${range.r0 + 1}` : `${range.r0 + 1}:${range.r1 + 1}`;
    }
  }
  return `${cellRef(range.r0, range.c0, r1c1)}:${cellRef(range.r1, range.c1, r1c1)}`;
}

/** Case-insensitive defined-name lookup. Returns the formula text stripped
 *  of any leading `=`, sheet qualifier, and `$` anchors so it can be parsed
 *  by `parseRangeRef` / `parseCellRef`. */
export function lookupDefinedName(wb: WorkbookHandle, query: string): string | null {
  if (!query) return null;
  const q = query.toLowerCase();
  for (const dn of wb.definedNames()) {
    if (dn.name.toLowerCase() !== q) continue;
    const eq = dn.formula.replace(/^=/, '');
    const bang = eq.lastIndexOf('!');
    return (bang >= 0 ? eq.slice(bang + 1) : eq).replace(/\$/g, '');
  }
  return null;
}

export function parseCellRef(raw: string): { row: number; col: number } | null {
  const trimmed = raw.trim().toUpperCase();
  // R1C1 form: e.g. "R5C2"
  const r1c1 = trimmed.match(/^R([1-9][0-9]*)C([1-9][0-9]*)$/);
  if (r1c1) {
    const row = Number.parseInt(r1c1[1] ?? '', 10) - 1;
    const col = Number.parseInt(r1c1[2] ?? '', 10) - 1;
    if (row < 0 || col < 0) return null;
    if (col > 16383 || row > 1048575) return null;
    return { row, col };
  }
  const m = trimmed.match(/^\$?([A-Z]+)\$?([1-9][0-9]*)$/);
  if (!m) return null;
  const letters = m[1] ?? '';
  const rowStr = m[2] ?? '';
  let col = 0;
  for (let i = 0; i < letters.length; i += 1) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  col -= 1;
  const row = Number.parseInt(rowStr, 10) - 1;
  if (col < 0 || row < 0) return null;
  if (col > 16383 || row > 1048575) return null;
  return { row, col };
}

/** Parse A1:B5 style range. Returns null when the input doesn't match. */
export function parseRangeRef(
  raw: string,
): { r0: number; c0: number; r1: number; c1: number } | null {
  const trimmed = raw.trim().toUpperCase();
  const wholeCol = trimmed.match(/^\$?([A-Z]+)(?::\$?([A-Z]+))?$/);
  if (wholeCol) {
    const c0 = parseColRef(wholeCol[1] ?? '');
    const c1 = parseColRef(wholeCol[2] ?? wholeCol[1] ?? '');
    if (c0 == null || c1 == null) return null;
    return { r0: 0, c0: Math.min(c0, c1), r1: 1048575, c1: Math.max(c0, c1) };
  }
  const wholeRow = trimmed.match(/^\$?([1-9][0-9]*)(?::\$?([1-9][0-9]*))?$/);
  if (wholeRow) {
    const r0 = Number.parseInt(wholeRow[1] ?? '', 10) - 1;
    const r1 = Number.parseInt(wholeRow[2] ?? wholeRow[1] ?? '', 10) - 1;
    if (r0 < 0 || r1 < 0 || r0 > 1048575 || r1 > 1048575) return null;
    return { r0: Math.min(r0, r1), c0: 0, r1: Math.max(r0, r1), c1: 16383 };
  }
  const parts = trimmed.split(':');
  if (parts.length !== 2) return null;
  const a = parseCellRef(parts[0] ?? '');
  const b = parseCellRef(parts[1] ?? '');
  if (!a || !b) return null;
  return {
    r0: Math.min(a.row, b.row),
    c0: Math.min(a.col, b.col),
    r1: Math.max(a.row, b.row),
    c1: Math.max(a.col, b.col),
  };
}

function parseColRef(raw: string): number | null {
  const letters = raw.trim().toUpperCase();
  if (!/^[A-Z]+$/.test(letters)) return null;
  let col = 0;
  for (let i = 0; i < letters.length; i += 1) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  col -= 1;
  return col >= 0 && col <= 16383 ? col : null;
}
