import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore } from '../store/store.js';

export function visibleSheetIndexes(wb: WorkbookHandle, store: SpreadsheetStore): number[] {
  const hidden = store.getState().layout.hiddenSheets;
  const out: number[] = [];
  for (let i = 0; i < wb.sheetCount; i += 1) {
    if (!hidden.has(i)) out.push(i);
  }
  return out;
}

/**
 * Sheets the Unhide affordance may offer. A very-hidden sheet is left out:
 * that state exists so a workbook can keep a settings or lookup sheet out of a
 * user's reach, and a spreadsheet's own Unhide list does not show it either.
 * Reaching such a sheet takes the host's own tooling, not the tab menu.
 */
export function hiddenSheetIndexes(wb: WorkbookHandle, store: SpreadsheetStore): number[] {
  const { hiddenSheets, veryHiddenSheets } = store.getState().layout;
  const out: number[] = [];
  for (let i = 0; i < wb.sheetCount; i += 1) {
    if (hiddenSheets.has(i) && !veryHiddenSheets.has(i)) out.push(i);
  }
  return out;
}
