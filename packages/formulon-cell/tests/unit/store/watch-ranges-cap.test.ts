import { describe, expect, it } from 'vitest';

import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('addWatchRanges cap', () => {
  it('applies the cap to the combined size of several ranges', () => {
    const store = createSpreadsheetStore();
    const ranges = [0, 1, 2].map((i) => ({
      sheet: 0,
      r0: i * 100,
      c0: 0,
      r1: i * 100 + 99,
      c1: 99,
    }));
    // 3 x 10_000 cells: each range fits on its own, the total does not.
    mutators.addWatchRanges(store, ranges);
    expect(store.getState().watch.watches.length).toBe(0);
  });

  it('counts overlapping ranges once against the cap', () => {
    const store = createSpreadsheetStore();
    const range = { sheet: 0, r0: 0, c0: 0, r1: 59, c1: 99 };
    mutators.addWatchRanges(store, [range, range]);
    expect(store.getState().watch.watches.length).toBe(6_000);
  });

  it('counts already watched cells toward the cap', () => {
    const store = createSpreadsheetStore();
    mutators.addWatchRanges(store, [{ sheet: 0, r0: 0, c0: 0, r1: 99, c1: 99 }]);
    expect(store.getState().watch.watches.length).toBe(10_000);
    mutators.addWatchRanges(store, [{ sheet: 0, r0: 200, c0: 0, r1: 200, c1: 0 }]);
    expect(store.getState().watch.watches.length).toBe(10_000);
  });
});
