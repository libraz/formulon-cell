import type { Addr } from './types.js';

export const addrKey = (a: Addr): string => `${a.sheet}:${a.row}:${a.col}`;

/** Last zero-based row index of a worksheet (row 1048576). */
export const MAX_ROW = 1_048_575;
/** Last zero-based column index of a worksheet (column XFD). */
export const MAX_COL = 16_383;

/** Convert a 0-indexed column number to A1 letter form ("A", "Z", "AA", …). */
export function colLetter(col: number): string {
  let n = col;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

/** Reverse of `colLetter`, case-insensitive. Returns -1 when `letters` is
 *  empty or holds a non-letter; the result is not bounds-checked. */
export function colFromLetters(letters: string): number {
  if (!letters) return -1;
  const upper = letters.toUpperCase();
  let col = 0;
  for (let i = 0; i < upper.length; i += 1) {
    const code = upper.charCodeAt(i);
    if (code < 65 || code > 90) return -1;
    col = col * 26 + (code - 64);
  }
  return col - 1;
}

/** Parse a single A1 atom like `B2` or `$A$1` (case-insensitive, surrounding
 *  whitespace ignored). Returns null when malformed or outside the sheet. */
export function parseA1Atom(raw: string): { row: number; col: number } | null {
  const match = raw.trim().match(/^\$?([A-Za-z]+)\$?(\d+)$/);
  if (!match) return null;
  const col = colFromLetters(match[1] ?? '');
  const row = Number.parseInt(match[2] ?? '', 10) - 1;
  if (col < 0 || row < 0 || col > MAX_COL || row > MAX_ROW) return null;
  return { row, col };
}

/** A1 text of a cell (`B3`), or `$B$3` when `absolute`. */
export function formatA1Cell(row: number, col: number, absolute = false): string {
  const pin = absolute ? '$' : '';
  return `${pin}${colLetter(col)}${pin}${row + 1}`;
}

/** A1 text of a range (`A1:B3`). A single cell collapses to `A1` unless
 *  `collapse` is false; `absolute` pins every coordinate (`$A$1:$B$3`). */
export function formatA1Range(
  range: { r0: number; c0: number; r1: number; c1: number },
  opts: { absolute?: boolean; collapse?: boolean } = {},
): string {
  const start = formatA1Cell(range.r0, range.c0, opts.absolute);
  const end = formatA1Cell(range.r1, range.c1, opts.absolute);
  return start === end && opts.collapse !== false ? start : `${start}:${end}`;
}
