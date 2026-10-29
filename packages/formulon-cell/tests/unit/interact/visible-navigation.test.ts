import { describe, expect, it } from 'vitest';
import {
  nextVisibleIndex,
  stepVisibleWithMerge,
  visibleAxisCount,
} from '../../../src/interact/visible-navigation.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('visible navigation', () => {
  it('skips contiguous hidden rows and columns in either direction', () => {
    const layout = createSpreadsheetStore().getState().layout;
    layout.hiddenRows = new Set([1, 2, 7]);
    layout.hiddenCols = new Set([1, 2, 7]);

    expect(nextVisibleIndex(layout, 'row', 0, 1, 0, 9)).toBe(3);
    expect(nextVisibleIndex(layout, 'row', 3, -1, 0, 9)).toBe(0);
    expect(nextVisibleIndex(layout, 'col', 0, 1, 0, 9)).toBe(3);
    expect(nextVisibleIndex(layout, 'col', 3, -1, 0, 9)).toBe(0);
    expect(nextVisibleIndex(layout, 'row', 7, 1, 0, 9)).toBe(8);
  });

  it('uses shown rows for default-hidden sheets without scanning hidden row limits', () => {
    const layout = createSpreadsheetStore().getState().layout;
    layout.defaultRowsHidden = true;
    layout.shownRows = new Set([2, 50, 900_000]);
    layout.hiddenRows = new Set([50, 100]);

    expect(nextVisibleIndex(layout, 'row', 0, 1, 0, 1_048_575)).toBe(2);
    expect(nextVisibleIndex(layout, 'row', 2, 1, 0, 1_048_575)).toBe(900_000);
    expect(nextVisibleIndex(layout, 'row', 900_000, 1, 0, 1_048_575)).toBeNull();
    expect(nextVisibleIndex(layout, 'row', 0, 1, 0, 1)).toBeNull();
    expect(visibleAxisCount(layout, 'row', 0, 1_000_000)).toBe(2);
  });

  it('matches Excel when entering a merge through a hidden anchor or a visible body', () => {
    const store = createSpreadsheetStore();
    mutators.mergeRange(store, { sheet: 0, r0: 1, r1: 3, c0: 0, c1: 0 });
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, hiddenRows: new Set([1, 2]) },
    }));

    expect(
      stepVisibleWithMerge(store.getState(), { sheet: 0, row: 0, col: 0 }, 1, 0, 20, 20),
    ).toEqual({
      sheet: 0,
      row: 4,
      col: 0,
    });
    expect(
      stepVisibleWithMerge(store.getState(), { sheet: 0, row: 4, col: 0 }, -1, 0, 20, 20),
    ).toEqual({
      sheet: 0,
      row: 1,
      col: 0,
    });

    const horizontal = createSpreadsheetStore();
    mutators.mergeRange(horizontal, { sheet: 0, r0: 0, r1: 0, c0: 1, c1: 3 });
    horizontal.setState((s) => ({
      ...s,
      layout: { ...s.layout, hiddenCols: new Set([1, 2]) },
    }));
    expect(
      stepVisibleWithMerge(horizontal.getState(), { sheet: 0, row: 0, col: 0 }, 0, 1, 20, 20),
    ).toEqual({
      sheet: 0,
      row: 0,
      col: 4,
    });
    expect(
      stepVisibleWithMerge(horizontal.getState(), { sheet: 0, row: 0, col: 4 }, 0, -1, 20, 20),
    ).toEqual({
      sheet: 0,
      row: 0,
      col: 1,
    });
  });

  it('skips a merge reached through a hidden gap when its anchor is hidden', () => {
    const vertical = createSpreadsheetStore();
    mutators.mergeRange(vertical, { sheet: 0, r0: 2, r1: 4, c0: 0, c1: 0 });
    vertical.setState((s) => ({
      ...s,
      layout: { ...s.layout, hiddenRows: new Set([1, 2, 3]) },
    }));
    expect(
      stepVisibleWithMerge(vertical.getState(), { sheet: 0, row: 0, col: 0 }, 1, 0, 20, 20),
    ).toEqual({
      sheet: 0,
      row: 5,
      col: 0,
    });

    const horizontal = createSpreadsheetStore();
    mutators.mergeRange(horizontal, { sheet: 0, r0: 0, r1: 0, c0: 2, c1: 4 });
    horizontal.setState((s) => ({
      ...s,
      layout: { ...s.layout, hiddenCols: new Set([1, 2, 3]) },
    }));
    expect(
      stepVisibleWithMerge(horizontal.getState(), { sheet: 0, row: 0, col: 0 }, 0, 1, 20, 20),
    ).toEqual({
      sheet: 0,
      row: 0,
      col: 5,
    });
  });

  it('enters a visible merge body after a hidden gap when moving in reverse across rows', () => {
    const store = createSpreadsheetStore();
    mutators.mergeRange(store, { sheet: 0, r0: 2, r1: 4, c0: 0, c1: 0 });
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, hiddenRows: new Set([1, 2, 3, 5]) },
    }));

    expect(
      stepVisibleWithMerge(store.getState(), { sheet: 0, row: 6, col: 0 }, -1, 0, 20, 20),
    ).toEqual({ sheet: 0, row: 2, col: 0 });
  });

  it('enters a visible merge body after a hidden gap when moving in reverse across columns', () => {
    const store = createSpreadsheetStore();
    mutators.mergeRange(store, { sheet: 0, r0: 0, r1: 0, c0: 2, c1: 4 });
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, hiddenCols: new Set([1, 2, 3, 5]) },
    }));

    expect(
      stepVisibleWithMerge(store.getState(), { sheet: 0, row: 0, col: 6 }, 0, -1, 20, 20),
    ).toEqual({ sheet: 0, row: 0, col: 2 });
  });

  it.each([
    {
      name: 'two row gaps',
      axis: 'row',
      merges: [
        [2, 4],
        [6, 8],
      ],
      hidden: [1, 2, 3, 5, 6, 7],
      start: 0,
      direction: 1,
      expected: 6,
    },
    {
      name: 'two direct row merges',
      axis: 'row',
      merges: [
        [1, 3],
        [4, 6],
      ],
      hidden: [1, 2, 4, 5],
      start: 0,
      direction: 1,
      expected: 4,
    },
    {
      name: 'three row gaps',
      axis: 'row',
      merges: [
        [2, 4],
        [6, 8],
        [10, 12],
      ],
      hidden: [1, 2, 3, 5, 6, 7, 9, 10, 11],
      start: 0,
      direction: 1,
      expected: 6,
    },
    {
      name: 'moving out of the first row merge',
      axis: 'row',
      merges: [
        [2, 4],
        [6, 8],
        [10, 12],
      ],
      hidden: [1, 2, 3, 5, 6, 7, 9, 10, 11],
      start: 4,
      direction: 1,
      expected: 6,
    },
    {
      name: 'visible gap after the first row merge',
      axis: 'row',
      merges: [
        [2, 4],
        [6, 8],
      ],
      hidden: [1, 2, 3, 6, 7],
      start: 0,
      direction: 1,
      expected: 5,
    },
    {
      name: 'fully hidden second row merge',
      axis: 'row',
      merges: [
        [2, 4],
        [6, 8],
      ],
      hidden: [1, 2, 3, 5, 6, 7, 8],
      start: 0,
      direction: 1,
      expected: 9,
    },
    {
      name: 'reverse into a later row merge',
      axis: 'row',
      merges: [
        [2, 4],
        [6, 8],
      ],
      hidden: [1, 2, 3, 5, 6, 7, 9],
      start: 10,
      direction: -1,
      expected: 6,
    },
    {
      name: 'two column gaps',
      axis: 'col',
      merges: [
        [2, 4],
        [6, 8],
      ],
      hidden: [1, 2, 3, 5, 6, 7],
      start: 0,
      direction: 1,
      expected: 6,
    },
    {
      name: 'three column gaps',
      axis: 'col',
      merges: [
        [2, 4],
        [6, 8],
        [10, 12],
      ],
      hidden: [1, 2, 3, 5, 6, 7, 9, 10, 11],
      start: 0,
      direction: 1,
      expected: 6,
    },
  ] as const)(
    'matches staged Excel merge resolution: $name',
    ({ axis, merges, hidden, start, direction, expected }) => {
      const store = createSpreadsheetStore();
      for (const [first, last] of merges) {
        mutators.mergeRange(
          store,
          axis === 'row'
            ? { sheet: 0, r0: first, r1: last, c0: 0, c1: 0 }
            : { sheet: 0, r0: 0, r1: 0, c0: first, c1: last },
        );
      }
      store.setState((s) => ({
        ...s,
        layout:
          axis === 'row'
            ? { ...s.layout, hiddenRows: new Set(hidden) }
            : { ...s.layout, hiddenCols: new Set(hidden) },
      }));
      const from =
        axis === 'row' ? { sheet: 0, row: start, col: 0 } : { sheet: 0, row: 0, col: start };
      const result = stepVisibleWithMerge(
        store.getState(),
        from,
        axis === 'row' ? direction : 0,
        axis === 'col' ? direction : 0,
        30,
        30,
      );

      expect(axis === 'row' ? result.row : result.col).toBe(expected);
    },
  );
});
