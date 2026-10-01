import { describe, expect, it } from 'vitest';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  attachNavigationPolicy,
  clampNavigationAddr,
  isNavigationAddrAllowed,
  navigationBoundsFor,
  nextTabStop,
} from '../../../src/interact/navigation-policy.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const workbook = (): WorkbookHandle => ({ sheetCount: 1 }) as WorkbookHandle;

describe('navigation policy', () => {
  it('clamps the existing selection when a fixed region is installed', () => {
    const store = createSpreadsheetStore();
    const handle = attachNavigationPolicy(store, workbook, {
      range: { sheet: 0, r0: 4, c0: 1, r1: 7, c1: 3 },
    });

    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 4, col: 1 });
    expect(store.getState().selection.anchor).toEqual({ sheet: 0, row: 4, col: 1 });
    expect(store.getState().selection.range).toEqual({
      sheet: 0,
      r0: 4,
      c0: 1,
      r1: 4,
      c1: 1,
    });

    handle.setOptions({ range: { sheet: 0, r0: 6, c0: 2, r1: 7, c1: 3 } });
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 6, col: 2 });
    expect(store.getState().selection.range).toEqual({
      sheet: 0,
      r0: 6,
      c0: 2,
      r1: 6,
      c1: 2,
    });
    handle.dispose();
  });

  it('keeps a nonzero viewport rectangle bounded while preserving outside data', () => {
    const store = createSpreadsheetStore();
    const outside = { sheet: 0, row: 40, col: 40 };
    mutators.setCell(store, outside, { kind: 'number', value: 99 }, '=1+2');

    const handle = attachNavigationPolicy(store, workbook, {
      range: { sheet: 0, r0: 5, c0: 3, r1: 7, c1: 5 },
    });
    mutators.setViewportSize(store, 2, 2);
    expect(navigationBoundsFor(store)).toEqual({ sheet: 0, r0: 5, c0: 3, r1: 7, c1: 5 });
    expect(store.getState().viewport.rowStart).toBe(5);
    expect(store.getState().viewport.colStart).toBe(3);

    mutators.scrollBy(store, 100, 100);
    expect(store.getState().viewport.rowStart).toBe(6);
    expect(store.getState().viewport.colStart).toBe(4);
    mutators.setActive(store, { sheet: 0, row: 100, col: 100 });
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 7, col: 5 });
    expect(store.getState().data.cells.get('0:40:40')).toEqual({
      value: { kind: 'number', value: 99 },
      formula: '=1+2',
    });
    handle.dispose();
  });

  it('bounds row, column, all, and disjoint selection mutators', () => {
    const store = createSpreadsheetStore();
    const handle = attachNavigationPolicy(store, workbook, {
      range: { sheet: 0, r0: 4, c0: 2, r1: 6, c1: 5 },
    });

    mutators.selectRow(store, 100);
    expect(store.getState().selection.range).toEqual({
      sheet: 0,
      r0: 6,
      c0: 2,
      r1: 6,
      c1: 5,
    });
    mutators.selectCols(store, 0, 100);
    expect(store.getState().selection.range).toEqual({
      sheet: 0,
      r0: 4,
      c0: 2,
      r1: 6,
      c1: 5,
    });
    mutators.selectAll(store);
    expect(store.getState().selection.range).toEqual({
      sheet: 0,
      r0: 4,
      c0: 2,
      r1: 6,
      c1: 5,
    });
    mutators.addExtraRange(store, {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 99,
      c1: 99,
    });
    expect(store.getState().selection.extraRanges).toEqual([]);
    handle.dispose();
  });

  it('traverses row-major stops forward and backward', () => {
    const store = createSpreadsheetStore();
    const handle = attachNavigationPolicy(store, workbook, {
      range: { sheet: 0, r0: 2, c0: 2, r1: 3, c1: 4 },
      tabNavigation: 'normal',
    });
    expect(nextTabStop(store, { sheet: 0, row: 2, col: 2 }, false)).toEqual({
      sheet: 0,
      row: 2,
      col: 3,
    });
    expect(nextTabStop(store, { sheet: 0, row: 3, col: 2 }, true)).toEqual({
      sheet: 0,
      row: 2,
      col: 4,
    });
    expect(nextTabStop(store, { sheet: 0, row: 3, col: 4 }, false)).toBeNull();
    handle.dispose();
  });

  it('rejects partial merged exposure and removes restrictions dynamically', () => {
    const store = createSpreadsheetStore();
    mutators.mergeRange(store, { sheet: 0, r0: 2, c0: 2, r1: 3, c1: 3 });
    expect(() =>
      attachNavigationPolicy(store, workbook, {
        range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 3 },
      }),
    ).toThrow(/partially exposes merged/);

    const handle = attachNavigationPolicy(store, workbook, {
      range: { sheet: 0, r0: 2, c0: 2, r1: 3, c1: 3 },
    });
    expect(isNavigationAddrAllowed(store, { sheet: 0, row: 0, col: 0 })).toBe(false);
    expect(clampNavigationAddr(store, { sheet: 0, row: 0, col: 0 })).toEqual({
      sheet: 0,
      row: 2,
      col: 2,
    });
    handle.setOptions();
    expect(navigationBoundsFor(store)).toBeUndefined();
    expect(isNavigationAddrAllowed(store, { sheet: 0, row: 0, col: 0 })).toBe(true);
  });
});
