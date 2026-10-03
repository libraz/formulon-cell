import { findPivotTableAtCell } from '../engine/passthrough-sync.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { WorkbookObjectsPanelHandle } from '../interact/workbook-objects.js';
import type { SpreadsheetStore } from '../store/store.js';

export interface PivotFieldListFollowDeps {
  store: SpreadsheetStore;
  getWb: () => WorkbookHandle;
  getWorkbookObjects: () => WorkbookObjectsPanelHandle | null;
}

// Keeps the pivot field list in step with the active cell: opens it on a pivot,
// closes it when the selection leaves every pivot.
export function attachPivotFieldListFollow(deps: PivotFieldListFollowDeps): { detach: () => void } {
  const { store, getWb, getWorkbookObjects } = deps;
  const unsub = store.subscribe((state, prevState) => {
    const workbookObjects = getWorkbookObjects();
    if (!workbookObjects) return;
    const wb = getWb();
    const prev = prevState.selection.active;
    const next = state.selection.active;
    if (prev.sheet === next.sheet && prev.row === next.row && prev.col === next.col) return;
    const nextPivot = findPivotTableAtCell(wb, next);
    if (!nextPivot) {
      if (workbookObjects.isPivotFieldListOpen()) workbookObjects.close();
      return;
    }
    const prevPivot = findPivotTableAtCell(wb, prev);
    const samePivot =
      prevPivot?.sheetIndex === nextPivot.sheetIndex &&
      prevPivot?.pivotIndex === nextPivot.pivotIndex;
    if (samePivot && workbookObjects.isPivotFieldListOpen()) return;
    workbookObjects.openPivotFieldList(nextPivot.sheetIndex, nextPivot.pivotIndex);
  });
  return { detach: unsub };
}
