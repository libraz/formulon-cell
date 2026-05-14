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

export function hiddenSheetIndexes(wb: WorkbookHandle, store: SpreadsheetStore): number[] {
  const hidden = store.getState().layout.hiddenSheets;
  const out: number[] = [];
  for (let i = 0; i < wb.sheetCount; i += 1) {
    if (hidden.has(i)) out.push(i);
  }
  return out;
}
