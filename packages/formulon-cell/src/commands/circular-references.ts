import { formatA1Cell, quoteSheetName } from '../engine/address.js';
import { sheetHasCells } from '../engine/sheet-kind.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';

/** One Formulas › Error Checking › Circular References entry. */
export interface CircularReferenceEntry {
  readonly sheet: number;
  readonly row: number;
  readonly col: number;
  /** The absolute address: bare on the active sheet, sheet-qualified elsewhere. */
  readonly label: string;
}

/**
 * The cells the Circular References list shows: the active sheet's cycle members, else those of
 * the first sheet in tab order that has any. Empty under iterative calculation, and when the
 * engine cannot list cycles. Each sheet read may recalculate, so call it when the list opens.
 */
export function circularReferenceEntries(
  wb: WorkbookHandle,
  activeSheet: number,
): CircularReferenceEntry[] {
  if (wb.getIterative()?.enabled === true) return [];
  const others = Array.from({ length: wb.sheetCount }, (_, sheet) => sheet).filter(
    (sheet) => sheet !== activeSheet,
  );
  for (const sheet of [activeSheet, ...others]) {
    if (!sheetHasCells(wb, sheet)) continue;
    const cells = wb.circularReferences(sheet);
    if (cells === null) return [];
    if (cells.length === 0) continue;
    const prefix = sheet === activeSheet ? '' : `${quoteSheetName(wb.sheetName(sheet))}!`;
    return cells.map(({ row, col }) => ({
      sheet,
      row,
      col,
      label: `${prefix}${formatA1Cell(row, col, true)}`,
    }));
  }
  return [];
}
