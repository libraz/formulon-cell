// A1-notation helpers shared by the React and Vue toolbar wrappers (range
// formatting and parsing for inline reference inputs in dialogs and the
// cell-label rendering used by report summaries).

import { colLetter } from '../engine/address.js';
import { parseRangeRef } from '../engine/range-resolver.js';
import type { SheetCell, SheetRange } from './toolbar-types.js';

export { parseA1Atom } from '../engine/address.js';

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

/** Parse `A1:B3` / `Sheet1!A1:B3` / `'My Sheet'!A1` into a `SheetRange` on the
 *  supplied sheet index. Cross-sheet refs are rejected unless they target
 *  `currentSheetName`. Returns null on bad input. */
export const parseA1Range = (
  raw: string,
  sheet: number,
  currentSheetName: string,
): SheetRange | null => {
  const parsed = parseRangeRef(raw);
  if (!parsed) return null;
  if (
    parsed.sheetName !== null &&
    parsed.sheetName.toLowerCase() !== currentSheetName.toLowerCase()
  ) {
    return null;
  }
  return { sheet, r0: parsed.r0, c0: parsed.c0, r1: parsed.r1, c1: parsed.c1 };
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
