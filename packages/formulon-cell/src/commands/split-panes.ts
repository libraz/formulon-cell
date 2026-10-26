import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { frozenColStart, frozenRowStart } from '../store/freeze-band.js';
import type { SpreadsheetStore } from '../store/store.js';
import { panesAtActiveCell, setFreezePanes } from './row-col-layout.js';

/**
 * View > Split. Splits above and left of the active cell (through the middle
 * of the window when it is the top-left cell), or removes a split already
 * there. Frozen panes turn into split panes at the same place, as the desktop
 * spreadsheet does. Not an undo step; with `wb` the split round-trips.
 */
export function toggleSplitPanes(store: SpreadsheetStore, wb?: WorkbookHandle): void {
  const { layout } = store.getState();
  if (layout.splitPanes === true) {
    setFreezePanes(store, 0, 0, wb);
    return;
  }
  if (layout.freezeRows > 0 || layout.freezeCols > 0) {
    const origin = { row: frozenRowStart(layout), col: frozenColStart(layout) };
    setFreezePanes(store, layout.freezeRows, layout.freezeCols, wb, origin, true);
    return;
  }
  const { rows, cols, origin } = panesAtActiveCell(store);
  setFreezePanes(store, rows, cols, wb, origin, true);
}
