import { describe, expect, it } from 'vitest';
import {
  activateSheetView,
  applySheetView,
  captureSheetView,
  deleteSheetView,
  findSheetView,
  removeSheetView,
  type SheetView,
  type SheetViewSnapshotInput,
  saveSheetView,
  upsertSheetView,
} from '../../../src/commands/sheet-views.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const baseSnap: SheetViewSnapshotInput = {
  sheet: 0,
  filterRange: null,
  freezeRows: 0,
  freezeCols: 0,
  hiddenRows: new Set(),
  hiddenCols: new Set(),
};

describe('captureSheetView', () => {
  it('captures only the fields that differ from defaults', () => {
    const v = captureSheetView('v1', 'Mine', baseSnap);
    expect(v).toEqual({ id: 'v1', name: 'Mine', sheet: 0 });
  });

  it('records freeze when either axis is engaged', () => {
    const v = captureSheetView('v1', 'X', { ...baseSnap, freezeRows: 1 });
    expect(v.freeze).toEqual({ rows: 1, cols: 0 });
  });

  it('records sorted hiddenRows/Cols', () => {
    const v = captureSheetView('v1', 'X', {
      ...baseSnap,
      hiddenRows: new Set([3, 1, 2]),
      hiddenCols: new Set([5, 4]),
    });
    expect(v.hiddenRows).toEqual([1, 2, 3]);
    expect(v.hiddenCols).toEqual([4, 5]);
  });

  it('captures filterRange and sort verbatim', () => {
    const range = { sheet: 0, r0: 0, c0: 0, r1: 5, c1: 3 };
    const sort = { column: 'B', direction: 'asc' as const };
    const v = captureSheetView('v1', 'X', { ...baseSnap, filterRange: range, sort });
    expect(v.filterRange).toEqual(range);
    expect(v.sort).toEqual(sort);
  });
});

describe('applySheetView', () => {
  it('flattens absent fields into null/zero defaults', () => {
    const view: SheetView = { id: 'v1', name: 'X', sheet: 0 };
    expect(applySheetView(view)).toEqual({
      sheet: 0,
      filterRange: null,
      freezeRows: 0,
      freezeCols: 0,
      hiddenRows: [],
      hiddenCols: [],
      sort: null,
    });
  });

  it('round-trips capture → apply', () => {
    const snap: SheetViewSnapshotInput = {
      sheet: 0,
      filterRange: { sheet: 0, r0: 0, c0: 0, r1: 9, c1: 2 },
      freezeRows: 1,
      freezeCols: 0,
      hiddenRows: new Set([2]),
      hiddenCols: new Set(),
      sort: { column: 'A', direction: 'desc' },
    };
    const v = captureSheetView('v1', 'My View', snap);
    const patch = applySheetView(v);
    expect(patch).toEqual({
      sheet: 0,
      filterRange: snap.filterRange,
      freezeRows: 1,
      freezeCols: 0,
      hiddenRows: [2],
      hiddenCols: [],
      sort: snap.sort,
    });
  });
});

describe('upsertSheetView / removeSheetView / findSheetView', () => {
  const a: SheetView = { id: 'a', name: 'A', sheet: 0 };
  const b: SheetView = { id: 'b', name: 'B', sheet: 1 };

  it('upsert appends a new view', () => {
    expect(upsertSheetView([a], b)).toEqual([a, b]);
  });

  it('upsert replaces an existing view by id', () => {
    const next = { ...a, name: 'A-renamed' };
    expect(upsertSheetView([a, b], next)).toEqual([b, next]);
  });

  it('remove returns the same array reference when id misses', () => {
    const arr = [a, b];
    expect(removeSheetView(arr, 'missing')).toBe(arr);
  });

  it('remove drops the matching view', () => {
    expect(removeSheetView([a, b], 'a')).toEqual([b]);
  });

  it('findSheetView returns the matching record or null', () => {
    expect(findSheetView([a, b], 'b')).toBe(b);
    expect(findSheetView([a, b], 'missing')).toBeNull();
  });
});

describe('store-backed sheet views', () => {
  it('saves the current view settings into the store', () => {
    const store = createSpreadsheetStore();
    const range = { sheet: 0, r0: 0, c0: 0, r1: 10, c1: 2 };
    mutators.setFilterRange(store, range);
    mutators.setFreezePanes(store, 1, 2);
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        hiddenRows: new Set([4, 2]),
        hiddenCols: new Set([3]),
      },
    }));

    const view = saveSheetView(store, 'v1', 'Review', { column: 'B', direction: 'asc' });

    expect(view).toMatchObject({
      id: 'v1',
      name: 'Review',
      sheet: 0,
      filterRange: range,
      freeze: { rows: 1, cols: 2 },
      hiddenRows: [2, 4],
      hiddenCols: [3],
      sort: { column: 'B', direction: 'asc' },
    });
    expect(store.getState().sheetViews.views).toEqual([view]);
  });

  it('activates a stored view on the current sheet', () => {
    const store = createSpreadsheetStore();
    saveSheetView(store, 'v1', 'Base');
    mutators.setFilterRange(store, { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 1 });
    mutators.setFreezePanes(store, 2, 1);
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, hiddenRows: new Set([6]), hiddenCols: new Set([2]) },
    }));
    saveSheetView(store, 'v2', 'Filtered');

    mutators.setFilterRange(store, null);
    mutators.setFreezePanes(store, 0, 0);
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, hiddenRows: new Set(), hiddenCols: new Set() },
    }));

    const result = activateSheetView(store, 'v2');
    const state = store.getState();

    expect(result.ok).toBe(true);
    expect(state.sheetViews.activeViewId).toBe('v2');
    expect(state.ui.filterRange).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 4, c1: 1 });
    expect(state.layout.freezeRows).toBe(2);
    expect(state.layout.freezeCols).toBe(1);
    expect([...state.layout.hiddenRows]).toEqual([6]);
    expect([...state.layout.hiddenCols]).toEqual([2]);
  });

  it('rejects missing or other-sheet views and deletes registered views', () => {
    const store = createSpreadsheetStore();
    mutators.upsertSheetView(store, { id: 'v2', name: 'Other', sheet: 1 });

    expect(activateSheetView(store, 'missing')).toEqual({ ok: false, reason: 'not-found' });
    expect(activateSheetView(store, 'v2')).toEqual({ ok: false, reason: 'different-sheet' });

    deleteSheetView(store, 'v2');
    expect(store.getState().sheetViews.views).toEqual([]);
  });
});
