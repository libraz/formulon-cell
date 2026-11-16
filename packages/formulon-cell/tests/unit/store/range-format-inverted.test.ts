import { describe, expect, it } from 'vitest';

import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('setRangeFormat with an inverted range', () => {
  it('formats the normalized rectangle', () => {
    const store = createSpreadsheetStore();
    mutators.setRangeFormat(store, { sheet: 0, r0: 2, c0: 2, r1: 0, c1: 0 }, { bold: true });
    expect(store.getState().format.formats.size).toBe(9);
  });

  it('applies the size guard to an inverted huge range', () => {
    const store = createSpreadsheetStore();
    mutators.setRangeFormat(store, { sheet: 0, r0: 999, c0: 999, r1: 0, c1: 0 }, { bold: true });
    expect(store.getState().format.formats.size).toBe(0);
  });
});
