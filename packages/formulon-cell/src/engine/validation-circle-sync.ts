import { cellValueViolatesValidation } from '../commands/validate.js';
import { validationAt } from '../store/range-validations.js';
import type { SpreadsheetStore } from '../store/store.js';
import { mutators } from '../store/store.js';
import { parseAddrKey } from './address.js';
import { makeRangeResolver } from './range-resolver.js';
import type { WorkbookHandle } from './workbook-handle.js';

/** Drop the Circle Invalid Data mark of every circled cell whose refreshed
 *  value satisfies its rule, whichever write path produced the value. A sheet
 *  switch is skipped: the cells are still being hydrated. */
export function bindValidationCircleStore(wb: WorkbookHandle, store: SpreadsheetStore): () => void {
  let lastCells = store.getState().data.cells;
  let lastSheet = store.getState().data.sheetIndex;
  return store.subscribe((state) => {
    const { cells, sheetIndex } = state.data;
    const sheetSwitched = sheetIndex !== lastSheet;
    const cellsChanged = cells !== lastCells;
    lastCells = cells;
    lastSheet = sheetIndex;
    const circles = state.errorIndicators.validationCircles;
    if (!cellsChanged || sheetSwitched || circles.size === 0) return;
    const resolveRange = makeRangeResolver(wb, sheetIndex);
    const kept = new Set<string>();
    for (const key of circles) {
      const addr = parseAddrKey(key);
      if (!addr || addr.sheet !== sheetIndex) {
        kept.add(key);
        continue;
      }
      const validation =
        state.format.formats.get(key)?.validation ?? validationAt(state.format, addr);
      const value = cells.get(key)?.value ?? { kind: 'blank' as const };
      if (cellValueViolatesValidation(value, validation, resolveRange)) kept.add(key);
    }
    if (kept.size !== circles.size) mutators.setValidationCircles(store, kept);
  });
}
