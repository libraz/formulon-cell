// @vitest-environment node

import { describe, expect, it } from 'vitest';
import { removeSheet } from '../../../src/commands/sheet-mutate.js';
import {
  hydrateViewStateFromEngine,
  nearestVisibleSheet,
  syncViewStateToEngine,
} from '../../../src/engine/view-state-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const reopen = async (wb: WorkbookHandle) => {
  const reloaded = await WorkbookHandle.loadBytes(wb.save());
  const store = createSpreadsheetStore();
  hydrateViewStateFromEngine(reloaded, store, { followActiveSheet: true });
  return { reloaded, store };
};

describe('saved view state', () => {
  it('reopens on the active sheet with its selection and scroll', async () => {
    const wb = await WorkbookHandle.createDefault();
    wb.addSheet();
    wb.addSheet();
    const store = createSpreadsheetStore();
    mutators.setSheetIndex(store, 1);
    store.setState((s) => {
      const a = { sheet: 1, row: 4, col: 2 };
      return {
        ...s,
        selection: {
          active: a,
          anchor: a,
          range: { sheet: 1, r0: 4, c0: 2, r1: 6, c1: 3 },
          extraRanges: [{ sheet: 1, r0: 10, c0: 0, r1: 10, c1: 0 }],
        },
        viewport: { ...s.viewport, rowStart: 20, colStart: 5 },
      };
    });
    wb.attachStore(store);

    const { reloaded, store: opened } = await reopen(wb);
    const s = opened.getState();
    expect(s.data.sheetIndex).toBe(1);
    expect(s.selection.active).toEqual({ sheet: 1, row: 4, col: 2 });
    expect(s.selection.range).toMatchObject({ r0: 4, c0: 2, r1: 6, c1: 3 });
    expect(s.selection.extraRanges).toMatchObject([{ r0: 10, c0: 0, r1: 10, c1: 0 }]);
    expect([s.viewport.rowStart, s.viewport.colStart]).toEqual([20, 5]);
    expect(reloaded.activeSheet()).toBe(1);
  });

  it('keeps a visited sheet it was switched away from', async () => {
    const wb = await WorkbookHandle.createDefault();
    wb.addSheet();
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 7, col: 1 });
    store.setState((s) => ({ ...s, viewport: { ...s.viewport, rowStart: 9 } }));
    mutators.setSheetIndex(store, 1);
    wb.attachStore(store);

    const { store: opened } = await reopen(wb);
    expect(opened.getState().data.sheetIndex).toBe(1);
    mutators.setSheetIndex(opened, 0);
    const s = opened.getState();
    expect(s.selection.active).toEqual({ sheet: 0, row: 7, col: 1 });
    expect(s.viewport.rowStart).toBe(9);
  });

  it('restores the frozen band origin and the scrolling pane start', async () => {
    const wb = await WorkbookHandle.createDefault();
    wb.setSheetFreeze(0, 2, 1);
    const store = createSpreadsheetStore();
    mutators.setFreezePanes(store, 2, 1, { row: 3, col: 0 });
    store.setState((s) => ({ ...s, viewport: { ...s.viewport, rowStart: 12, colStart: 4 } }));
    wb.attachStore(store);

    const { store: opened } = await reopen(wb);
    const s = opened.getState();
    expect(s.layout.freezeOriginRow).toBe(3);
    expect(s.layout.freezeOriginCol).toBe(0);
    expect([s.viewport.rowStart, s.viewport.colStart]).toEqual([12, 4]);
  });

  it('opens on the next visible sheet when the active one is hidden', async () => {
    const wb = await WorkbookHandle.createDefault();
    wb.addSheet();
    wb.addSheet();
    const store = createSpreadsheetStore();
    mutators.setSheetIndex(store, 1);
    syncViewStateToEngine(wb, store);
    wb.setSheetTabHidden(1, true);
    expect(wb.activeSheet()).toBe(2);
  });
});

describe('sheet structure changes', () => {
  it('keeps each sheet memory with its sheet and loads the right neighbour on remove', async () => {
    const wb = await WorkbookHandle.createDefault();
    wb.addSheet();
    wb.addSheet();
    const store = createSpreadsheetStore();
    wb.attachStore(store);
    mutators.setActive(store, { sheet: 0, row: 1, col: 1 });
    mutators.setSheetIndex(store, 2);
    mutators.setActive(store, { sheet: 2, row: 8, col: 3 });
    mutators.setSheetIndex(store, 1);
    mutators.setActive(store, { sheet: 1, row: 5, col: 0 });

    expect(removeSheet(store, wb, 1)).toBe(true);
    // The old sheet 2 slid into index 1 and opens with its own selection.
    let s = store.getState();
    expect(s.data.sheetIndex).toBe(1);
    expect(s.selection.active).toEqual({ sheet: 1, row: 8, col: 3 });
    mutators.setSheetIndex(store, 0);
    s = store.getState();
    expect(s.selection.active).toEqual({ sheet: 0, row: 1, col: 1 });
  });
});

describe('nearestVisibleSheet', () => {
  it('prefers the right neighbour, then the nearest on the left', () => {
    expect(nearestVisibleSheet(4, new Set([1]), 1)).toBe(2);
    expect(nearestVisibleSheet(4, new Set([2, 3]), 2)).toBe(1);
    expect(nearestVisibleSheet(2, new Set([0, 1]), 0)).toBe(-1);
  });
});
