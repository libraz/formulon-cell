import { setSheetZoom } from '../commands/view.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import {
  frozenColsWidth,
  frozenRowsHeight,
  gridOriginX,
  gridOriginY,
  layoutForView,
} from '../render/geometry.js';
import { renderScale, steppedZoom } from '../store/render-scale.js';
import { mutators, type SpreadsheetStore, type State } from '../store/store.js';
import { navigationPolicyFor } from './navigation-policy.js';
import { gridLocal } from './pointer-targets.js';

export interface WheelDeps {
  /** Element to listen on — typically the grid canvas wrapper. */
  grid: HTMLElement;
  store: SpreadsheetStore;
  /** When provided, Ctrl/Cmd+wheel zoom changes are also pushed to the current
   *  engine (`setSheetZoom`) so the active-sheet zoom round-trips through .xlsx. */
  getWb?: () => WorkbookHandle | null;
}

/** Whether the pointer is over the top and/or left pane of a split. */
function splitPaneAt(
  grid: HTMLElement,
  state: State,
  e: WheelEvent,
): { top: boolean; left: boolean } {
  if (state.layout.splitPanes !== true) return { top: false, left: false };
  const { layout, x, y } = gridLocal(grid, state, e);
  return {
    top:
      layout.freezeRows > 0 && y < gridOriginY(layout) + frozenRowsHeight(layout, state.viewport),
    left:
      layout.freezeCols > 0 && x < gridOriginX(layout) + frozenColsWidth(layout, state.viewport),
  };
}

/** Where the pointer sits relative to frozen panes (not splits), in screen px. */
function frozenRegionAt(
  grid: HTMLElement,
  state: State,
  e: WheelEvent,
): { blockY: boolean; blockX: boolean } {
  const layout = layoutForView(state);
  if (state.layout.splitPanes === true || (layout.freezeRows === 0 && layout.freezeCols === 0)) {
    return { blockY: false, blockX: false };
  }
  const { x, y } = gridLocal(grid, state, e);
  const originX = gridOriginX(layout);
  const originY = gridOriginY(layout);
  // Headers and the frozen rows take no wheel input; the frozen columns scroll rows only.
  if (x < originX || y < originY) return { blockY: true, blockX: true };
  if (layout.freezeRows > 0 && y < originY + frozenRowsHeight(layout, state.viewport)) {
    return { blockY: true, blockX: true };
  }
  const inFrozenCols =
    layout.freezeCols > 0 && x < originX + frozenColsWidth(layout, state.viewport);
  return { blockY: false, blockX: inFrozenCols };
}

/** Mac line-wheel rates, in rows and columns per line; rows round half down, columns truncate. */
const MAC_ROWS_PER_LINE = 0.7;
const MAC_COLS_PER_LINE = 0.15;

/** Idle time after which the next wheel event starts a new gesture. */
const WHEEL_GESTURE_GAP_MS = 250;

const DOM_DELTA_LINE = 1;
const DOM_DELTA_PAGE = 2;

/** Px per wheel delta unit: one cell per line, `count` cells per page, else raw px. */
const wheelUnit = (deltaMode: number, cell: number, count: number): number =>
  deltaMode === DOM_DELTA_LINE ? cell : deltaMode === DOM_DELTA_PAGE ? cell * count : 1;

/**
 * Wheel-to-scroll bridge.
 *
 * - **Ctrl/Cmd + wheel** → zoom (delegates to `setSheetZoom`).
 * - **Shift + wheel**    → horizontal scroll (deltaY redirected to deltaX).
 * - **Plain wheel**      → vertical (deltaY) and horizontal (deltaX) scroll.
 *   Over a split's top or left pane, that pane scrolls on its own axis.
 *
 * Trackpads emit small pixel-precise deltas (1–5 px), which would round to
 * zero rows on every event. We accumulate residuals across events so a slow
 * scroll still advances by one row once enough pixels have accumulated.
 */
export function attachWheel(deps: WheelDeps): () => void {
  const { grid, store, getWb } = deps;

  let accY = 0;
  let accX = 0;
  let lastStamp = Number.NEGATIVE_INFINITY;
  let lastMode = 0;
  let lastSignY = 0;
  let lastSignX = 0;

  const onWheel = (e: WheelEvent): void => {
    if (e.ctrlKey || e.metaKey) {
      e.preventDefault();
      // A bounded embedded viewport owns zoom as session presentation state.
      // Do not persist it through the workbook unless the host explicitly
      // mounted an unrestricted navigation policy; this keeps a wheel gesture
      // from becoming an unexpected workbook mutation.
      const navOptions = navigationPolicyFor(store)?.options;
      const restricted = !!(navOptions?.range || navOptions?.selectable);
      setSheetZoom(
        store,
        steppedZoom(store.getState().viewport.zoom, e.deltaY < 0 ? 1 : -1),
        restricted ? undefined : (getWb?.() ?? undefined),
      );
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
    // Line deltas step one row/column per line, page deltas a page of them.
    const unitY = wheelUnit(e.deltaMode, rh, state.viewport.rowCount);
    const unitX = wheelUnit(e.deltaMode, cw, state.viewport.colCount);
    const mac =
      e.deltaMode === DOM_DELTA_LINE &&
      (grid.closest<HTMLElement>('.fc-host') ?? grid).dataset.fcPlatform === 'mac';
    const frozen = mac ? frozenRegionAt(grid, state, e) : { blockY: false, blockX: false };
    if (frozen.blockY && frozen.blockX) return;
    const rowRate = mac ? MAC_ROWS_PER_LINE : 1;
    const colRate = mac ? MAC_COLS_PER_LINE : 1;
    const dx = frozen.blockX ? 0 : (e.shiftKey ? e.deltaY : e.deltaX) * unitX * dirX * colRate;
    const dy = e.shiftKey || frozen.blockY ? 0 : e.deltaY * unitY * rowRate;
    // A pause, a unit change or a reversal ends the gesture; its partial row is not carried over.
    if (e.timeStamp - lastStamp > WHEEL_GESTURE_GAP_MS || e.deltaMode !== lastMode) {
      accY = 0;
      accX = 0;
    }
    lastStamp = e.timeStamp;
    lastMode = e.deltaMode;
    // A rounded Mac step can leave a remainder against the direction, so compare the deltas.
    if (dy * lastSignY < 0) accY = 0;
    if (dx * lastSignX < 0) accX = 0;
    if (dy !== 0) lastSignY = Math.sign(dy);
    if (dx !== 0) lastSignX = Math.sign(dx);
    accY += dy;
    accX += dx;
    const dRow = mac ? -Math.round(-accY / rh) : Math.trunc(accY / rh);
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
