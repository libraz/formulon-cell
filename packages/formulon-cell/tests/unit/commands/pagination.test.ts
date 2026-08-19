import { beforeEach, describe, expect, it } from 'vitest';
import {
  bandIndexAt,
  pageLayoutGaps,
  pageNumberOf,
  paginationFor,
  resetPaginationCache,
} from '../../../src/commands/pagination.js';
import { splitAxisIntoBands } from '../../../src/commands/print.js';
import { addrKey } from '../../../src/engine/address.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

/** Put a value in a cell so the used range — and therefore the paginated
 *  extent — reaches it. */
const seed = (store: SpreadsheetStore, row: number, col: number): void => {
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    cells.set(addrKey({ sheet: 0, row, col }), {
      value: { kind: 'number', value: 1 },
      formula: null,
    });
    return { ...s, data: { ...s.data, cells } };
  });
};

describe('splitAxisIntoBands', () => {
  it('closes a band when the next index would overflow the page', () => {
    const bands = splitAxisIntoBands({ from: 0, to: 9, sizeOf: () => 30, budget: 100 });
    expect(bands.map((b) => [b.start, b.end])).toEqual([
      [0, 2],
      [3, 5],
      [6, 8],
      [9, 9],
    ]);
  });

  it('opens a band at a manual break even when the page has room left', () => {
    const bands = splitAxisIntoBands({
      from: 0,
      to: 9,
      sizeOf: () => 10,
      budget: 1000,
      manualBreaks: [4],
    });
    expect(bands).toEqual([
      { start: 0, end: 3, manual: false },
      { start: 4, end: 9, manual: true },
    ]);
  });

  it('ignores manual breaks outside the range being paginated', () => {
    const bands = splitAxisIntoBands({
      from: 3,
      to: 6,
      sizeOf: () => 10,
      budget: 1000,
      manualBreaks: [1, 9],
    });
    expect(bands).toEqual([{ start: 3, end: 6, manual: false }]);
  });

  it('gives an index wider than a whole page its own band rather than looping', () => {
    const bands = splitAxisIntoBands({ from: 0, to: 2, sizeOf: () => 500, budget: 100 });
    expect(bands).toHaveLength(3);
  });

  it('skips hidden indices, which measure zero', () => {
    const hidden = new Set([1, 2]);
    const bands = splitAxisIntoBands({
      from: 0,
      to: 5,
      sizeOf: (i) => (hidden.has(i) ? 0 : 60),
      budget: 100,
    });
    // Rows 1 and 2 cost nothing, so they ride along with row 0; row 3 is the
    // second visible row and would overflow the page.
    expect(bands[0]).toEqual({ start: 0, end: 2, manual: false });
    expect(bands[1]?.start).toBe(3);
  });
});

describe('paginationFor', () => {
  beforeEach(() => resetPaginationCache());

  it('anchors at A1 and reaches past the used range so blank pages exist', () => {
    const store = createSpreadsheetStore();
    seed(store, 4, 3);
    const pagination = paginationFor(store.getState(), 0);

    expect(pagination.origin).toEqual({ row: 0, col: 0 });
    expect(pagination.content).toEqual({ row: 4, col: 3 });
    // Only the first page carries content, but later pages are laid out too.
    expect(pagination.pageCount).toBe(1);
    expect(pagination.rowBands.length).toBeGreaterThan(1);
    expect(pagination.colBands.length).toBeGreaterThan(1);
  });

  it('anchors at the print area when one is set', () => {
    const store = createSpreadsheetStore();
    seed(store, 40, 8);
    mutators.setPageSetup(store, 0, { printArea: 'C3:E9' });
    const pagination = paginationFor(store.getState(), 0);

    expect(pagination.origin).toEqual({ row: 2, col: 2 });
    expect(pagination.content).toEqual({ row: 8, col: 4 });
  });

  it('honours a manual row break', () => {
    const store = createSpreadsheetStore();
    seed(store, 20, 0);
    mutators.setPageSetup(store, 0, { manualPageBreakRows: [5] });
    const pagination = paginationFor(store.getState(), 0);

    expect(pagination.rowBands[0]).toEqual({ start: 0, end: 4, manual: false });
    expect(pagination.rowBands[1]?.start).toBe(5);
    expect(pagination.rowBands[1]?.manual).toBe(true);
  });

  it('reuses the memoised result until an input slice changes', () => {
    const store = createSpreadsheetStore();
    seed(store, 4, 3);
    const first = paginationFor(store.getState(), 0);
    expect(paginationFor(store.getState(), 0)).toBe(first);

    mutators.setPageSetup(store, 0, { orientation: 'landscape' });
    expect(paginationFor(store.getState(), 0)).not.toBe(first);
  });

  it('rounds the reach so scrolling one row does not invalidate the cache', () => {
    const store = createSpreadsheetStore();
    seed(store, 4, 3);
    const first = paginationFor(store.getState(), 0, { throughRow: 10, throughCol: 10 });
    expect(paginationFor(store.getState(), 0, { throughRow: 11, throughCol: 10 })).toBe(first);
  });
});

describe('bandIndexAt / pageNumberOf', () => {
  beforeEach(() => resetPaginationCache());

  it('maps an index to the band that holds it', () => {
    const bands = [
      { start: 0, end: 3, manual: false },
      { start: 4, end: 9, manual: true },
    ];
    expect(bandIndexAt(bands, 0)).toBe(0);
    expect(bandIndexAt(bands, 4)).toBe(1);
    expect(bandIndexAt(bands, 99)).toBe(-1);
  });

  it('numbers pages down-then-over by default and leaves blank paper unnumbered', () => {
    const store = createSpreadsheetStore();
    // Two row bands and two column bands, all carrying content.
    seed(store, 120, 30);
    mutators.setPageSetup(store, 0, { manualPageBreakRows: [10], manualPageBreakCols: [4] });
    const pagination = paginationFor(store.getState(), 0);

    expect(pageNumberOf(pagination, 0, 0)).toBe(1);
    expect(pageNumberOf(pagination, 1, 0)).toBe(2);
    expect(pageNumberOf(pagination, 0, 1, 'overThenDown')).toBeGreaterThan(1);
    expect(pageNumberOf(pagination, 0, pagination.colBands.length - 1)).toBe(0);
  });
});

describe('pageLayoutGaps', () => {
  beforeEach(() => resetPaginationCache());

  it('opens one margin before the first page and both plus a gutter after it', () => {
    const store = createSpreadsheetStore();
    seed(store, 200, 0);
    const state = store.getState();
    const pagination = paginationFor(state, 0);
    const { pageGapRows } = pageLayoutGaps(state, 0);

    const first = pagination.rowBands[0];
    const second = pagination.rowBands[1];
    expect(first && pageGapRows.get(first.start)).toBeCloseTo(pagination.marginPx.top, 5);
    expect(second && pageGapRows.get(second.start)).toBeGreaterThan(
      pagination.marginPx.top + pagination.marginPx.bottom,
    );
  });

  it('scales the gutters with the viewport zoom', () => {
    const store = createSpreadsheetStore();
    seed(store, 200, 0);
    const base = pageLayoutGaps(store.getState(), 0).pageGapRows.get(0);

    mutators.setZoom(store, 2);
    resetPaginationCache();
    const zoomed = pageLayoutGaps(store.getState(), 0).pageGapRows.get(0);

    expect(base).toBeGreaterThan(0);
    expect(zoomed).toBeCloseTo((base ?? 0) * 2, 5);
  });
});
