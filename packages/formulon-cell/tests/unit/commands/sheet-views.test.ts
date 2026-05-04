import { describe, expect, it } from 'vitest';
import {
  applySheetView,
  captureSheetView,
  findSheetView,
  removeSheetView,
  type SheetView,
  type SheetViewSnapshotInput,
  upsertSheetView,
} from '../../../src/commands/sheet-views.js';

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
