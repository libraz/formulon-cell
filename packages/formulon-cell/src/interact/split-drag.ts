import { captureLayoutSnapshot, type LayoutSnapshot } from '../commands/slice-history.js';
import { MAX_COL, MAX_ROW } from '../engine/address.js';
import { syncLayoutToEngine } from '../engine/layout-sync.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import {
  colWidth,
  frozenColsWidth,
  frozenRowsHeight,
  gridOriginX,
  gridOriginY,
  layoutForView,
  mirrorX,
  rowHeight,
} from '../render/geometry.js';
import { SPLIT_BAR_PX } from '../render/grid/lines.js';
import { frozenColStart, frozenRowStart } from '../store/freeze-band.js';
import { mutators, type SpreadsheetStore, type State } from '../store/store.js';

export interface SplitDragDeps {
  grid: HTMLElement;
  store: SpreadsheetStore;
  getWb: () => WorkbookHandle | null;
}

interface BarHit {
  row: boolean;
  col: boolean;
}

/** Rows or columns from `origin` whose far edge lies nearest `target` px. */
const countToward = (
  target: number,
  origin: number,
  limit: number,
  size: (i: number) => number,
): number => {
  let px = 0;
  let count = 0;
  while (origin + count < limit) {
    const s = size(origin + count);
    if (px + s / 2 > target) break;
    px += s;
    count += 1;
  }
  return count;
};

/**
 * Split-bar dragging: a bar follows the pointer, snapping to row and column
 * boundaries, and an axis dragged back into the headers loses its split. A
 * double-click on a bar removes that split. The engine is written once, on
 * release. Listens in the capture phase so a grab never reaches the grid.
 */
export function attachSplitDrag(deps: SplitDragDeps): () => void {
  const { grid, store, getWb } = deps;
  let drag:
    | (BarHit & { before: LayoutSnapshot; sheet: number; origin: { row: number; col: number } })
    | null = null;

  const local = (e: MouseEvent, state: State) => {
    const rect = grid.getBoundingClientRect();
    const layout = layoutForView(state);
    return { layout, x: mirrorX(layout, e.clientX - rect.left), y: e.clientY - rect.top };
  };

  const barAt = (e: MouseEvent): BarHit | null => {
    const state = store.getState();
    if (state.layout.splitPanes !== true) return null;
    const { layout, x, y } = local(e, state);
    const barY = gridOriginY(layout) + frozenRowsHeight(layout, state.viewport);
    const barX = gridOriginX(layout) + frozenColsWidth(layout, state.viewport);
    const row = layout.freezeRows > 0 && Math.abs(y - barY) <= SPLIT_BAR_PX;
    const col = layout.freezeCols > 0 && Math.abs(x - barX) <= SPLIT_BAR_PX;
    return row || col ? { row, col } : null;
  };

  const cursorFor = (hit: BarHit): string =>
    hit.row && hit.col ? 'move' : hit.row ? 'row-resize' : 'col-resize';

  const paneOrigin = (state: State) => ({
    row: frozenRowStart(state.layout),
    col: frozenColStart(state.layout),
  });

  const setSplit = (rows: number, cols: number, origin: { row: number; col: number }): void => {
    mutators.setFreezePanes(store, rows, cols, origin, true);
  };

  const commit = (before: LayoutSnapshot, sheet: number): void => {
    const wb = getWb();
    if (!wb) return;
    const state = store.getState();
    syncLayoutToEngine(wb, state.layout, sheet, before, captureLayoutSnapshot(state));
  };

  const onDown = (e: PointerEvent): void => {
    if (e.button !== 0) return;
    const hit = barAt(e);
    if (!hit) return;
    e.preventDefault();
    e.stopPropagation();
    const state = store.getState();
    // The origin survives an axis dragged to zero, which would otherwise reset it.
    drag = {
      ...hit,
      before: captureLayoutSnapshot(state),
      sheet: state.data.sheetIndex,
      origin: paneOrigin(state),
    };
    try {
      grid.setPointerCapture(e.pointerId);
    } catch {
      /* synthetic events have no active pointer to capture */
    }
  };

  const onMove = (e: PointerEvent): void => {
    if (!drag) {
      const hit = barAt(e);
      if (!hit) return;
      grid.style.cursor = cursorFor(hit);
      e.stopPropagation();
      return;
    }
    e.stopPropagation();
    const state = store.getState();
    const { layout, x, y } = local(e, state);
    const { viewport } = state;
    const rows = drag.row
      ? countToward(y - gridOriginY(layout), drag.origin.row, MAX_ROW + 1, (r) =>
          rowHeight(layout, r, viewport),
        )
      : layout.freezeRows;
    const cols = drag.col
      ? countToward(x - gridOriginX(layout), drag.origin.col, MAX_COL + 1, (c) =>
          colWidth(layout, c, viewport),
        )
      : layout.freezeCols;
    if (rows !== layout.freezeRows || cols !== layout.freezeCols) setSplit(rows, cols, drag.origin);
  };

  const onUp = (e: PointerEvent): void => {
    if (!drag) return;
    e.stopPropagation();
    const { before, sheet } = drag;
    drag = null;
    grid.style.cursor = '';
    commit(before, sheet);
  };

  const onDblClick = (e: MouseEvent): void => {
    const hit = barAt(e);
    if (!hit) return;
    e.preventDefault();
    e.stopPropagation();
    const state = store.getState();
    const before = captureLayoutSnapshot(state);
    setSplit(
      hit.row ? 0 : state.layout.freezeRows,
      hit.col ? 0 : state.layout.freezeCols,
      paneOrigin(state),
    );
    commit(before, state.data.sheetIndex);
  };

  grid.addEventListener('pointerdown', onDown, true);
  grid.addEventListener('pointermove', onMove, true);
  grid.addEventListener('pointerup', onUp, true);
  grid.addEventListener('pointercancel', onUp, true);
  grid.addEventListener('dblclick', onDblClick, true);
  return () => {
    grid.removeEventListener('pointerdown', onDown, true);
    grid.removeEventListener('pointermove', onMove, true);
    grid.removeEventListener('pointerup', onUp, true);
    grid.removeEventListener('pointercancel', onUp, true);
    grid.removeEventListener('dblclick', onDblClick, true);
  };
}
