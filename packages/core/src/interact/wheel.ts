import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { type SpreadsheetStore, mutators } from '../store/store.js';

export interface WheelDeps {
  /** Element to listen on — typically the grid canvas wrapper. */
  grid: HTMLElement;
  store: SpreadsheetStore;
  /** When provided, Ctrl/Cmd+wheel zoom changes are also pushed to the engine
   *  (`setSheetZoom`) so the active-sheet zoom round-trips through .xlsx. */
  wb?: WorkbookHandle;
}

/**
 * Wheel-to-scroll bridge.
 *
 * - **Ctrl/Cmd + wheel** → zoom (delegates to `mutators.setZoom`).
 * - **Shift + wheel**    → horizontal scroll (deltaY redirected to deltaX).
 * - **Plain wheel**      → vertical (deltaY) and horizontal (deltaX) scroll.
 *
 * Trackpads emit small pixel-precise deltas (1–5 px), which would round to
 * zero rows on every event. We accumulate residuals across events so a slow
 * scroll still advances by one row once enough pixels have accumulated.
 */
export function attachWheel(deps: WheelDeps): () => void {
  const { grid, store, wb } = deps;

  let accY = 0;
  let accX = 0;

  const onWheel = (e: WheelEvent): void => {
    if (e.ctrlKey || e.metaKey) {
      e.preventDefault();
      const cur = store.getState().viewport.zoom;
      const step = e.deltaY < 0 ? 0.1 : -0.1;
      mutators.setZoom(store, Math.round((cur + step) * 10) / 10);
      if (wb) {
        const sheet = store.getState().data.sheetIndex;
        const pct = Math.round(store.getState().viewport.zoom * 100);
        wb.setSheetZoom(sheet, pct);
      }
      return;
    }
    const layout = store.getState().layout;
    const rh = Math.max(1, layout.defaultRowHeight);
    const cw = Math.max(1, layout.defaultColWidth);
    const dx = e.shiftKey ? e.deltaY : e.deltaX;
    const dy = e.shiftKey ? 0 : e.deltaY;
    accY += dy;
    accX += dx;
    const dRow = Math.trunc(accY / rh);
    const dCol = Math.trunc(accX / cw);
    if (dRow === 0 && dCol === 0) return;
    accY -= dRow * rh;
    accX -= dCol * cw;
    e.preventDefault();
    mutators.scrollBy(store, dRow, dCol);
  };

  grid.addEventListener('wheel', onWheel, { passive: false });
  return () => grid.removeEventListener('wheel', onWheel);
}
