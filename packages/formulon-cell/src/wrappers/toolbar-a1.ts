// A1-notation helpers shared by the React and Vue toolbar wrappers (range
// formatting and parsing for inline reference inputs in dialogs and the
// cell-label rendering used by report summaries).

import { colLetter } from '../commands/print.js';
import type { SheetCell, SheetRange } from './toolbar-types.js';

/** Render a `SheetRange` as A1 ("A1:B3" or "A1" when start === end). */
export const formatA1Range = (range: SheetRange): string => {
  const start = `${colLetter(range.c0)}${range.r0 + 1}`;
  const end = `${colLetter(range.c1)}${range.r1 + 1}`;
  return start === end ? start : `${start}:${end}`;
};

/** Whether a sheet name can precede `!` without quotes. Being a plain
 *  identifier is not enough: a name shaped like a reference (`A1`, `R1C1`) or
 *  made only of digits (`2024`) reads as part of the address once the quotes
 *  are gone, so those keep them. */
const isBareSheetName = (name: string): boolean =>
  /^[A-Za-z_][A-Za-z0-9_]*$/.test(name) &&
  !/^[A-Za-z]{1,3}\d+$/.test(name) &&
  !/^[Rr]\d+[Cc]\d+$/.test(name);

/** Render a `SheetRange` the way a spreadsheet's Create Table / Create
 *  PivotTable dialogs show a source range: sheet-qualified and absolute
 *  (`Sheet1!$A$1:$B$3`). Sheet names that cannot stand unquoted are
 *  single-quoted, matching what `parseA1Range` accepts back. */
export const formatSheetAbsoluteRange = (sheetName: string, range: SheetRange): string => {
  const start = `$${colLetter(range.c0)}$${range.r0 + 1}`;
  const end = `$${colLetter(range.c1)}$${range.r1 + 1}`;
  const body = start === end ? start : `${start}:${end}`;
  if (!sheetName) return body;
  const prefix = isBareSheetName(sheetName) ? sheetName : `'${sheetName.replace(/'/g, "''")}'`;
  return `${prefix}!${body}`;
};

/** Parse a single A1 atom like `$A$1` or `B2`. Returns null when malformed. */
export const parseA1Atom = (raw: string): { row: number; col: number } | null => {
  const match = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(raw.trim());
  if (!match) return null;
  const letters = match[1] ?? '';
  let col = 0;
  for (let i = 0; i < letters.length; i += 1) {
    col = col * 26 + (letters.toUpperCase().charCodeAt(i) - 64);
  }
  const row = Number.parseInt(match[2] ?? '', 10) - 1;
  col -= 1;
  return col >= 0 && row >= 0 ? { row, col } : null;
};

/** Parse `A1:B3` / `Sheet1!A1:B3` / `'My Sheet'!A1` into a `SheetRange` on the
 *  supplied sheet index. Cross-sheet refs are rejected unless they target
 *  `currentSheetName`. Returns null on bad input. */
export const parseA1Range = (
  raw: string,
  sheet: number,
  currentSheetName: string,
): SheetRange | null => {
  const trimmed = raw.trim().replace(/^=/, '');
  if (!trimmed) return null;
  let body = trimmed;
  const bang = trimmed.indexOf('!');
  if (bang !== -1) {
    let sheetName = trimmed.slice(0, bang);
    body = trimmed.slice(bang + 1);
    if (sheetName.startsWith("'") && sheetName.endsWith("'")) {
      sheetName = sheetName.slice(1, -1).replace(/''/g, "'");
    }
    if (sheetName.toLowerCase() !== currentSheetName.toLowerCase()) return null;
  }
  const parts = body.split(':');
  if (parts.length < 1 || parts.length > 2) return null;
  const head = parseA1Atom(parts[0] ?? '');
  const tail = parts.length === 2 ? parseA1Atom(parts[1] ?? '') : head;
  if (!head || !tail) return null;
  return {
    sheet,
    r0: Math.min(head.row, tail.row),
    c0: Math.min(head.col, tail.col),
    r1: Math.max(head.row, tail.row),
    c1: Math.max(head.col, tail.col),
  };
};

/** Render a cell as a human-readable string for dialog summaries. */
export const cellLabel = (cell: SheetCell | undefined): string => {
  if (!cell) return '';
  const value = cell.value;
  if (value.kind === 'number') return String(value.value);
  if (value.kind === 'text') return value.value;
  if (value.kind === 'bool') return value.value ? 'TRUE' : 'FALSE';
  return '';
};
