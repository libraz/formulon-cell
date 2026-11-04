import { describe, expect, it } from 'vitest';

import { directCellFormat, sameAddr, sameFormatTarget } from '../../../src/store/direct-format.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('store/direct-format', () => {
  it('compares sheet, row, and column when checking address equality', () => {
    expect(sameAddr({ sheet: 0, row: 1, col: 2 }, { sheet: 0, row: 1, col: 2 })).toBe(true);
    expect(sameAddr({ sheet: 0, row: 1, col: 2 }, { sheet: 1, row: 1, col: 2 })).toBe(false);
    expect(sameAddr({ sheet: 0, row: 1, col: 2 }, { sheet: 0, row: 9, col: 2 })).toBe(false);
    expect(sameAddr({ sheet: 0, row: 1, col: 2 }, { sheet: 0, row: 1, col: 9 })).toBe(false);
  });

  it('returns the stored format, or undefined for an unformatted cell', () => {
    const store = createSpreadsheetStore();
    const addr = { sheet: 0, row: 1, col: 1 };
    mutators.setCellFormat(store, addr, { bold: true, fill: '#ffff00' });

    expect(directCellFormat(store.getState(), addr)).toEqual({ bold: true, fill: '#ffff00' });
    expect(directCellFormat(store.getState(), { sheet: 0, row: 2, col: 2 })).toBeUndefined();
  });

  it('resolves a merged body through its anchor', () => {
    const store = createSpreadsheetStore();
    const anchor = { sheet: 0, row: 0, col: 0 };
    const body = { sheet: 0, row: 0, col: 1 };
    const other = { sheet: 0, row: 0, col: 3 };
    mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });
    mutators.setCellFormat(store, anchor, { numFmt: { kind: 'text' }, bold: true });

    expect(sameFormatTarget(store.getState(), body, anchor)).toBe(true);
    expect(sameFormatTarget(store.getState(), body, other)).toBe(false);
    expect(directCellFormat(store.getState(), body)).toEqual({
      numFmt: { kind: 'text' },
      bold: true,
    });
  });
});
