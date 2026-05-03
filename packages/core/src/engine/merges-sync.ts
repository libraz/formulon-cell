import { type SpreadsheetStore, mutators } from '../store/store.js';
import type { WorkbookHandle } from './workbook-handle.js';

/** Replace the store's merges for `sheet` with whatever the engine reports.
 *  No-op when `capabilities.merges` is off — the JS-side state stays as-is.
 *  Only operates on `sheet`; merges on other sheets are left untouched so
 *  this can be called per-tab without dropping cross-tab state. */
export function hydrateMergesFromEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  if (!wb.capabilities.merges) return;
  const ranges = wb.getMerges(sheet);
  // Drop existing merges on this sheet, then add the engine's set.
  const state = store.getState();
  for (const anchorKey of state.merges.byAnchor.keys()) {
    const r = state.merges.byAnchor.get(anchorKey);
    if (!r || r.sheet !== sheet) continue;
    mutators.unmergeRange(store, r);
  }
  for (const r of ranges) mutators.mergeRange(store, r);
}
