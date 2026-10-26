import type { WorkbookHandle } from '../engine/workbook-handle.js';
import {
  frozenColsWidth,
  frozenRowsHeight,
  gridOriginX,
  gridOriginY,
  layoutForView,
  mirrorX,
} from '../render/geometry.js';
import { renderScale, steppedZoom } from '../store/render-scale.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { navigationPolicyFor } from './navigation-policy.js';

export interface WheelDeps {
  /** Element to listen on — typically the grid canvas wrapper. */
  grid: HTMLElement;
  store: SpreadsheetStore;
  /** When provided, Ctrl/Cmd+wheel zoom changes are also pushed to the current
   *  engine (`setSheetZoom`) so the active-sheet zoom round-trips through .xlsx. */
  getWb?: () => WorkbookHandle | null;
}

/**
 * Wheel-to-scroll bridge.
 *
 * - **Ctrl/Cmd + wheel** → zoom (delegates to `mutators.setZoom`).
 * - **Shift + wheel**    → horizontal scroll (deltaY redirected to deltaX).
 * - **Plain wheel**      → vertical (deltaY) and horizontal (deltaX) scroll.
 *   Over a split's top or left pane, that pane scrolls on its own axis.
 *
 * Trackpads emit small pixel-precise deltas (1–5 px), which would round to
 * zero rows on every event. We accumulate residuals across events so a slow
 * scroll still advances by one row once enough pixels have accumulated.
 */
/** Whether the pointer is over the top and/or left pane of a split. */
function splitPaneAt(
  grid: HTMLElement,
  state: ReturnType<SpreadsheetStore['getState']>,
  e: WheelEvent,
): { top: boolean; left: boolean } {
  if (state.layout.splitPanes !== true) return { top: false, left: false };
  const layout = layoutForView(state);
  const rect = grid.getBoundingClientRect();
  const x = mirrorX(layout, e.clientX - rect.left);
  const y = e.clientY - rect.top;
  return {
    top:
      layout.freezeRows > 0 && y < gridOriginY(layout) + frozenRowsHeight(layout, state.viewport),
    left:
      layout.freezeCols > 0 && x < gridOriginX(layout) + frozenColsWidth(layout, state.viewport),
  };
}

export function attachWheel(deps: WheelDeps): () => void {
  const { grid, store, getWb } = deps;

  let accY = 0;
  let accX = 0;

  const onWheel = (e: WheelEvent): void => {
    if (e.ctrlKey || e.metaKey) {
      e.preventDefault();
      mutators.setZoom(store, steppedZoom(store.getState().viewport.zoom, e.deltaY < 0 ? 1 : -1));
      // A bounded embedded viewport owns zoom as session presentation state.
      // Do not persist it through the workbook unless the host explicitly
      // mounted an unrestricted navigation policy; this keeps a wheel gesture
      // from becoming an unexpected workbook mutation.
      const navOptions = navigationPolicyFor(store)?.options;
      const restricted = !!(navOptions?.range || navOptions?.selectable);
      const wb = getWb?.();
      if (wb && !restricted) {
        const sheet = store.getState().data.sheetIndex;
        const pct = Math.round(store.getState().viewport.zoom * 100);
        wb.setSheetZoom(sheet, pct);
      }
      return;
    }
    const state = store.getState();
    const layout = state.layout;
    // Step thresholds are screen px, matching the wheel delta.
    const scale = renderScale(layout, state.viewport);
    const rh = Math.max(1, layout.defaultRowHeight * scale);
    const cw = Math.max(1, layout.defaultColWidth * scale);
    // A right-to-left sheet advances its columns leftwards, so a rightward
    // scroll gesture has to walk the column index backwards.
    const dirX = state.ui.rightToLeft === true ? -1 : 1;
    const dx = (e.shiftKey ? e.deltaY : e.deltaX) * dirX;
    const dy = e.shiftKey ? 0 : e.deltaY;
    accY += dy;
    accX += dx;
    const dRow = Math.trunc(accY / rh);
    const dCol = Math.trunc(accX / cw);
    if (dRow === 0 && dCol === 0) return;
    accY -= dRow * rh;
    accX -= dCol * cw;
    e.preventDefault();
    const pane = splitPaneAt(grid, state, e);
    const bandRow = pane.top ? dRow : 0;
    const bandCol = pane.left ? dCol : 0;
    if (bandRow !== 0 || bandCol !== 0) mutators.scrollSplitPane(store, bandRow, bandCol);
    if (dRow - bandRow !== 0 || dCol - bandCol !== 0) {
      mutators.scrollBy(store, dRow - bandRow, dCol - bandCol);
    }
  };

  grid.addEventListener('wheel', onWheel, { passive: false });
  return () => grid.removeEventListener('wheel', onWheel);
}
