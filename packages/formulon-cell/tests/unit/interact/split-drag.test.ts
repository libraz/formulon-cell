import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachSplitDrag } from '../../../src/interact/split-drag.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const fire = (el: HTMLElement, type: string, x: number, y: number): PointerEvent => {
  const e = new PointerEvent(type, {
    clientX: x,
    clientY: y,
    button: 0,
    bubbles: true,
    cancelable: true,
    pointerId: 1,
  });
  el.dispatchEvent(e);
  return e;
};

describe('attachSplitDrag', () => {
  let grid: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let detach: () => void;

  beforeEach(() => {
    grid = document.createElement('div');
    document.body.appendChild(grid);
    store = createSpreadsheetStore();
    store.setState((s) => ({ ...s, viewport: { ...s.viewport, widthPx: 800 } }));
    wb = {
      capabilities: { freeze: true, sheetSplit: true },
      setSheetSplit: vi.fn(() => true),
      setSheetFreeze: vi.fn(() => true),
    } as unknown as WorkbookHandle;
    // Split below row 5 and right of column B: bars at y = 20 + 100, x = 32 + 128.
    mutators.setFreezePanes(store, 5, 2, { row: 0, col: 0 }, true);
    detach = attachSplitDrag({ grid, store, getWb: () => wb });
  });

  afterEach(() => {
    detach();
    grid.remove();
  });

  it('drags the horizontal bar to the nearest row boundary and saves on release', () => {
    const down = fire(grid, 'pointerdown', 400, 121);
    expect(down.defaultPrevented).toBe(true);
    fire(grid, 'pointermove', 400, 228);
    expect(store.getState().layout).toMatchObject({ freezeRows: 10, freezeCols: 2 });
    expect(wb.setSheetSplit).not.toHaveBeenCalled();
    fire(grid, 'pointerup', 400, 228);
    expect(wb.setSheetSplit).toHaveBeenCalledWith(0, 2400, 3300);
  });

  it('drags the vertical bar', () => {
    fire(grid, 'pointerdown', 159, 400);
    fire(grid, 'pointermove', 290, 400);
    fire(grid, 'pointerup', 290, 400);
    expect(store.getState().layout).toMatchObject({ freezeRows: 5, freezeCols: 4 });
  });

  it('removes the split on an axis dragged back to the headers', () => {
    fire(grid, 'pointerdown', 400, 120);
    fire(grid, 'pointermove', 400, 5);
    fire(grid, 'pointerup', 400, 5);
    expect(store.getState().layout).toMatchObject({
      freezeRows: 0,
      freezeCols: 2,
      splitPanes: true,
    });
  });

  it('keeps the pane origin when a bar is dragged into the headers and back', () => {
    mutators.setFreezePanes(store, 5, 2, { row: 3, col: 0 }, true);
    fire(grid, 'pointerdown', 400, 120);
    fire(grid, 'pointermove', 400, 5);
    fire(grid, 'pointermove', 400, 228);
    fire(grid, 'pointerup', 400, 228);
    expect(store.getState().layout).toMatchObject({ freezeRows: 10, freezeOriginRow: 3 });
  });

  it('removes the split on a double-click on its bar', () => {
    grid.dispatchEvent(new MouseEvent('dblclick', { clientX: 400, clientY: 120, bubbles: true }));
    expect(store.getState().layout).toMatchObject({ freezeRows: 0, freezeCols: 2 });
  });

  it('leaves clicks away from the bars and frozen panes to the grid', () => {
    expect(fire(grid, 'pointerdown', 400, 300).defaultPrevented).toBe(false);
    mutators.setFreezePanes(store, 5, 2, { row: 0, col: 0 });
    expect(fire(grid, 'pointerdown', 400, 120).defaultPrevented).toBe(false);
  });
});
