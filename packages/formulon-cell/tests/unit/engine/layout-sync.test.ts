import { describe, expect, it } from 'vitest';
import { captureLayoutSnapshot } from '../../../src/commands/history.js';
import {
  hydrateLayoutFromEngine,
  syncLayoutSizesToEngine,
  syncLayoutToEngine,
} from '../../../src/engine/layout-sync.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

interface FakeColLayout {
  first: number;
  last: number;
  width: number;
  hidden: boolean;
  outlineLevel: number;
}
interface FakeRowLayout {
  row: number;
  height: number;
  hidden: boolean;
  outlineLevel: number;
}

interface SetColCall {
  sheet: number;
  first: number;
  last: number;
  width: number;
}
interface SetRowCall {
  sheet: number;
  row: number;
  height: number;
}
interface SetFreezeCall {
  sheet: number;
  freezeRows: number;
  freezeCols: number;
}
interface SetColHiddenCall {
  sheet: number;
  first: number;
  last: number;
  hidden: boolean;
}
interface SetRowHiddenCall {
  sheet: number;
  row: number;
  hidden: boolean;
}
interface SetColOutlineCall {
  sheet: number;
  first: number;
  last: number;
  level: number;
}
interface SetRowOutlineCall {
  sheet: number;
  row: number;
  level: number;
}

interface FakeWb {
  wb: WorkbookHandle;
  colCalls: SetColCall[];
  rowCalls: SetRowCall[];
  freezeCalls: SetFreezeCall[];
  colHiddenCalls: SetColHiddenCall[];
  rowHiddenCalls: SetRowHiddenCall[];
  colOutlineCalls: SetColOutlineCall[];
  rowOutlineCalls: SetRowOutlineCall[];
}

const makeFake = (opts: {
  colRowSize: boolean;
  freeze?: boolean;
  sheetZoom?: boolean;
  hiddenRowsCols?: boolean;
  outlines?: boolean;
  cols?: FakeColLayout[];
  rows?: FakeRowLayout[];
  view?: { zoomScale: number; freezeRows: number; freezeCols: number; tabHidden: boolean } | null;
}): FakeWb => {
  const colCalls: SetColCall[] = [];
  const rowCalls: SetRowCall[] = [];
  const freezeCalls: SetFreezeCall[] = [];
  const colHiddenCalls: SetColHiddenCall[] = [];
  const rowHiddenCalls: SetRowHiddenCall[] = [];
  const colOutlineCalls: SetColOutlineCall[] = [];
  const rowOutlineCalls: SetRowOutlineCall[] = [];
  const fake = {
    capabilities: {
      colRowSize: opts.colRowSize,
      freeze: opts.freeze ?? false,
      sheetZoom: opts.sheetZoom ?? false,
      hiddenRowsCols: opts.hiddenRowsCols ?? false,
      outlines: opts.outlines ?? false,
    },
    getColumnLayouts: () => opts.cols ?? [],
    getRowLayouts: () => opts.rows ?? [],
    getSheetView: () => opts.view ?? null,
    setColumnWidth: (sheet: number, first: number, last: number, width: number) => {
      colCalls.push({ sheet, first, last, width });
      return opts.colRowSize;
    },
    setRowHeight: (sheet: number, row: number, height: number) => {
      rowCalls.push({ sheet, row, height });
      return opts.colRowSize;
    },
    setSheetFreeze: (sheet: number, freezeRows: number, freezeCols: number) => {
      freezeCalls.push({ sheet, freezeRows, freezeCols });
      return opts.freeze ?? false;
    },
    setColumnHidden: (sheet: number, first: number, last: number, hidden: boolean) => {
      colHiddenCalls.push({ sheet, first, last, hidden });
      return opts.hiddenRowsCols ?? false;
    },
    setRowHidden: (sheet: number, row: number, hidden: boolean) => {
      rowHiddenCalls.push({ sheet, row, hidden });
      return opts.hiddenRowsCols ?? false;
    },
    setColumnOutline: (sheet: number, first: number, last: number, level: number) => {
      colOutlineCalls.push({ sheet, first, last, level });
      return opts.outlines ?? false;
    },
    setRowOutline: (sheet: number, row: number, level: number) => {
      rowOutlineCalls.push({ sheet, row, level });
      return opts.outlines ?? false;
    },
  };
  return {
    wb: fake as unknown as WorkbookHandle,
    colCalls,
    rowCalls,
    freezeCalls,
    colHiddenCalls,
    rowHiddenCalls,
    colOutlineCalls,
    rowOutlineCalls,
  };
};

describe('hydrateLayoutFromEngine', () => {
  it('seeds colWidths/rowHeights from engine snapshots', () => {
    const store = createSpreadsheetStore();
    const { wb } = makeFake({
      colRowSize: true,
      cols: [{ first: 0, last: 2, width: 120, hidden: false, outlineLevel: 0 }],
      rows: [{ row: 5, height: 40, hidden: false, outlineLevel: 0 }],
    });
    hydrateLayoutFromEngine(wb, store, 0);
    const s = store.getState();
    expect(s.layout.colWidths.get(0)).toBe(120);
    expect(s.layout.colWidths.get(1)).toBe(120);
    expect(s.layout.colWidths.get(2)).toBe(120);
    expect(s.layout.colWidths.get(3)).toBeUndefined();
    expect(s.layout.rowHeights.get(5)).toBe(40);
  });

  it('hydrates hidden flags and outline levels from the same vectors', () => {
    const store = createSpreadsheetStore();
    const { wb } = makeFake({
      colRowSize: true,
      cols: [{ first: 4, last: 4, width: 0, hidden: true, outlineLevel: 2 }],
      rows: [{ row: 3, height: 0, hidden: true, outlineLevel: 1 }],
    });
    hydrateLayoutFromEngine(wb, store, 0);
    const s = store.getState();
    expect(s.layout.hiddenCols.has(4)).toBe(true);
    expect(s.layout.outlineCols.get(4)).toBe(2);
    expect(s.layout.hiddenRows.has(3)).toBe(true);
    expect(s.layout.outlineRows.get(3)).toBe(1);
    // width=0 / height=0 must NOT seed a size override.
    expect(s.layout.colWidths.has(4)).toBe(false);
    expect(s.layout.rowHeights.has(3)).toBe(false);
  });

  it('is a no-op under the stub (capability off, empty engine arrays)', () => {
    const store = createSpreadsheetStore();
    const { wb } = makeFake({ colRowSize: false });
    hydrateLayoutFromEngine(wb, store, 0);
    const s = store.getState();
    expect(s.layout.colWidths.size).toBe(0);
    expect(s.layout.rowHeights.size).toBe(0);
  });

  it('seeds freezeRows/freezeCols and zoom from getSheetView', () => {
    const store = createSpreadsheetStore();
    const { wb } = makeFake({
      colRowSize: false,
      sheetZoom: true,
      view: { zoomScale: 150, freezeRows: 1, freezeCols: 2, tabHidden: false },
    });
    hydrateLayoutFromEngine(wb, store, 0);
    const s = store.getState();
    expect(s.layout.freezeRows).toBe(1);
    expect(s.layout.freezeCols).toBe(2);
    expect(s.viewport.zoom).toBeCloseTo(1.5);
  });

  it('leaves zoom at 1.0 when engine reports the default 100%', () => {
    const store = createSpreadsheetStore();
    const initialZoom = store.getState().viewport.zoom;
    const { wb } = makeFake({
      colRowSize: false,
      sheetZoom: true,
      view: { zoomScale: 100, freezeRows: 0, freezeCols: 0, tabHidden: false },
    });
    hydrateLayoutFromEngine(wb, store, 0);
    expect(store.getState().viewport.zoom).toBe(initialZoom);
  });

  it('clamps engine zoom into the store-supported range', () => {
    const store = createSpreadsheetStore();
    const { wb } = makeFake({
      colRowSize: false,
      sheetZoom: true,
      view: { zoomScale: 400, freezeRows: 0, freezeCols: 0, tabHidden: false },
    });
    hydrateLayoutFromEngine(wb, store, 0);
    expect(store.getState().viewport.zoom).toBe(4);
  });
});

describe('syncLayoutSizesToEngine', () => {
  it('writes only the columns whose width changed', () => {
    const store = createSpreadsheetStore();
    const { wb, colCalls, rowCalls } = makeFake({ colRowSize: true });
    const before = captureLayoutSnapshot(store.getState());
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, colWidths: new Map([[2, 200]]) },
    }));
    const after = captureLayoutSnapshot(store.getState());
    syncLayoutSizesToEngine(wb, store.getState().layout, 0, before, after);
    expect(colCalls).toEqual([{ sheet: 0, first: 2, last: 2, width: 200 }]);
    expect(rowCalls).toEqual([]);
  });

  it('writes the default width when an override is removed', () => {
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, colWidths: new Map([[1, 180]]) },
    }));
    const before = captureLayoutSnapshot(store.getState());
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, colWidths: new Map() },
    }));
    const after = captureLayoutSnapshot(store.getState());
    const { wb, colCalls } = makeFake({ colRowSize: true });
    syncLayoutSizesToEngine(wb, store.getState().layout, 0, before, after);
    expect(colCalls).toEqual([
      { sheet: 0, first: 1, last: 1, width: store.getState().layout.defaultColWidth },
    ]);
  });

  it('skips engine calls when capability is off', () => {
    const store = createSpreadsheetStore();
    const before = captureLayoutSnapshot(store.getState());
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        colWidths: new Map([[0, 150]]),
        rowHeights: new Map([[0, 40]]),
      },
    }));
    const after = captureLayoutSnapshot(store.getState());
    const { wb, colCalls, rowCalls } = makeFake({ colRowSize: false });
    syncLayoutSizesToEngine(wb, store.getState().layout, 0, before, after);
    expect(colCalls).toEqual([]);
    expect(rowCalls).toEqual([]);
  });

  it('writes both column and row deltas in a single pass', () => {
    const store = createSpreadsheetStore();
    const before = captureLayoutSnapshot(store.getState());
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        colWidths: new Map([[0, 150]]),
        rowHeights: new Map([[3, 50]]),
      },
    }));
    const after = captureLayoutSnapshot(store.getState());
    const { wb, colCalls, rowCalls } = makeFake({ colRowSize: true });
    syncLayoutSizesToEngine(wb, store.getState().layout, 1, before, after);
    expect(colCalls).toEqual([{ sheet: 1, first: 0, last: 0, width: 150 }]);
    expect(rowCalls).toEqual([{ sheet: 1, row: 3, height: 50 }]);
  });
});

describe('syncLayoutToEngine — hidden flags', () => {
  it('emits setColumnHidden / setRowHidden for added entries', () => {
    const store = createSpreadsheetStore();
    const before = captureLayoutSnapshot(store.getState());
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        hiddenCols: new Set([3]),
        hiddenRows: new Set([7]),
      },
    }));
    const after = captureLayoutSnapshot(store.getState());
    const { wb, colHiddenCalls, rowHiddenCalls } = makeFake({
      colRowSize: false,
      hiddenRowsCols: true,
    });
    syncLayoutToEngine(wb, store.getState().layout, 0, before, after);
    expect(colHiddenCalls).toEqual([{ sheet: 0, first: 3, last: 3, hidden: true }]);
    expect(rowHiddenCalls).toEqual([{ sheet: 0, row: 7, hidden: true }]);
  });

  it('emits hidden=false when an entry is removed', () => {
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, hiddenCols: new Set([2]) },
    }));
    const before = captureLayoutSnapshot(store.getState());
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, hiddenCols: new Set() },
    }));
    const after = captureLayoutSnapshot(store.getState());
    const { wb, colHiddenCalls } = makeFake({ colRowSize: false, hiddenRowsCols: true });
    syncLayoutToEngine(wb, store.getState().layout, 0, before, after);
    expect(colHiddenCalls).toEqual([{ sheet: 0, first: 2, last: 2, hidden: false }]);
  });

  it('skips hidden sync when capability is off', () => {
    const store = createSpreadsheetStore();
    const before = captureLayoutSnapshot(store.getState());
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, hiddenCols: new Set([0]), hiddenRows: new Set([0]) },
    }));
    const after = captureLayoutSnapshot(store.getState());
    const { wb, colHiddenCalls, rowHiddenCalls } = makeFake({
      colRowSize: false,
      hiddenRowsCols: false,
    });
    syncLayoutToEngine(wb, store.getState().layout, 0, before, after);
    expect(colHiddenCalls).toEqual([]);
    expect(rowHiddenCalls).toEqual([]);
  });
});

describe('syncLayoutToEngine — outline levels', () => {
  it('emits setColumnOutline / setRowOutline on level changes', () => {
    const store = createSpreadsheetStore();
    const before = captureLayoutSnapshot(store.getState());
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        outlineCols: new Map([[1, 2]]),
        outlineRows: new Map([[4, 1]]),
      },
    }));
    const after = captureLayoutSnapshot(store.getState());
    const { wb, colOutlineCalls, rowOutlineCalls } = makeFake({
      colRowSize: false,
      outlines: true,
    });
    syncLayoutToEngine(wb, store.getState().layout, 2, before, after);
    expect(colOutlineCalls).toEqual([{ sheet: 2, first: 1, last: 1, level: 2 }]);
    expect(rowOutlineCalls).toEqual([{ sheet: 2, row: 4, level: 1 }]);
  });

  it('emits level=0 to clear an outline entry', () => {
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, outlineRows: new Map([[5, 3]]) },
    }));
    const before = captureLayoutSnapshot(store.getState());
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, outlineRows: new Map() },
    }));
    const after = captureLayoutSnapshot(store.getState());
    const { wb, rowOutlineCalls } = makeFake({ colRowSize: false, outlines: true });
    syncLayoutToEngine(wb, store.getState().layout, 0, before, after);
    expect(rowOutlineCalls).toEqual([{ sheet: 0, row: 5, level: 0 }]);
  });
});

describe('syncLayoutToEngine — freeze panes', () => {
  it('emits setSheetFreeze when freeze counts change', () => {
    const store = createSpreadsheetStore();
    const before = captureLayoutSnapshot(store.getState());
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, freezeRows: 2, freezeCols: 3 },
    }));
    const after = captureLayoutSnapshot(store.getState());
    const { wb, freezeCalls } = makeFake({ colRowSize: false, freeze: true });
    syncLayoutToEngine(wb, store.getState().layout, 0, before, after);
    expect(freezeCalls).toEqual([{ sheet: 0, freezeRows: 2, freezeCols: 3 }]);
  });

  it('does not emit setSheetFreeze when freeze is unchanged', () => {
    const store = createSpreadsheetStore();
    const before = captureLayoutSnapshot(store.getState());
    const after = captureLayoutSnapshot(store.getState());
    const { wb, freezeCalls } = makeFake({ colRowSize: false, freeze: true });
    syncLayoutToEngine(wb, store.getState().layout, 0, before, after);
    expect(freezeCalls).toEqual([]);
  });
});
