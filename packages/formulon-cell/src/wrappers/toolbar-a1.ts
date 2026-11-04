// A1-notation helpers shared by the React and Vue toolbar wrappers (range
// formatting and parsing for inline reference inputs in dialogs).

import { formatA1Range, quoteSheetName } from '../engine/address.js';
import { parseRangeRef } from '../engine/range-resolver.js';
import type { SheetRange } from './toolbar-types.js';

/** Render a `SheetRange` the way a spreadsheet's Create Table / Create
 *  PivotTable dialogs show a source range: sheet-qualified and absolute
 *  (`Sheet1!$A$1:$B$3`). Sheet names that cannot stand unquoted are
 *  single-quoted, matching what `parseA1Range` accepts back. */
export const formatSheetAbsoluteRange = (sheetName: string, range: SheetRange): string => {
  const body = formatA1Range(range, { absolute: true });
  if (!sheetName) return body;
  return `${quoteSheetName(sheetName)}!${body}`;
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
