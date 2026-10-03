import { describe, expect, it } from 'vitest';

import {
  formatWithPending,
  sameAddr,
  sameFormatTarget,
} from '../../../src/store/pending-format.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('store/pending-format', () => {
  it('compares sheet, row, and column when checking address equality', () => {
    expect(sameAddr({ sheet: 0, row: 1, col: 2 }, { sheet: 0, row: 1, col: 2 })).toBe(true);
    expect(sameAddr({ sheet: 0, row: 1, col: 2 }, { sheet: 1, row: 1, col: 2 })).toBe(false);
    expect(sameAddr({ sheet: 0, row: 1, col: 2 }, { sheet: 0, row: 9, col: 2 })).toBe(false);
    expect(sameAddr({ sheet: 0, row: 1, col: 2 }, { sheet: 0, row: 1, col: 9 })).toBe(false);
  });

  it('returns the stored format when there is no pending format for that address', () => {
    const store = createSpreadsheetStore();
    const addr = { sheet: 0, row: 1, col: 1 };
    mutators.setCellFormat(store, addr, { bold: true, fill: '#ffff00' });
    mutators.setPendingFormat(store, {
      addr: { sheet: 0, row: 1, col: 2 },
      format: { italic: true },
    });

    expect(formatWithPending(store.getState(), addr)).toEqual({
      bold: true,
      fill: '#ffff00',
    });
  });

  it('overlays pending format onto stored format for the matching address', () => {
    const store = createSpreadsheetStore();
    const addr = { sheet: 0, row: 1, col: 1 };
    mutators.setCellFormat(store, addr, { bold: false, fill: '#ffff00', align: 'left' });
    mutators.setPendingFormat(store, {
      addr,
      format: { bold: true, italic: true },
    });

    expect(formatWithPending(store.getState(), addr)).toEqual({
      bold: true,
      fill: '#ffff00',
      align: 'left',
      italic: true,
    });
  });

  it('returns pending-only format for cells without stored format', () => {
    const store = createSpreadsheetStore();
    const addr = { sheet: 0, row: 2, col: 2 };
    mutators.setPendingFormat(store, {
      addr,
      format: { underline: true },
    });

    expect(formatWithPending(store.getState(), addr)).toEqual({ underline: true });
  });

  it('resolves stored and pending formats through a merged anchor', () => {
    const store = createSpreadsheetStore();
    const anchor = { sheet: 0, row: 0, col: 0 };
    const body = { sheet: 0, row: 0, col: 1 };
    const other = { sheet: 0, row: 0, col: 3 };
    mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });
    mutators.setCellFormat(store, anchor, { numFmt: { kind: 'text' }, bold: false });
    mutators.setPendingFormat(store, { addr: body, format: { bold: true } });

    expect(sameFormatTarget(store.getState(), body, anchor)).toBe(true);
    expect(sameFormatTarget(store.getState(), body, other)).toBe(false);
    expect(formatWithPending(store.getState(), anchor)).toEqual({
      numFmt: { kind: 'text' },
      bold: true,
    });
    expect(formatWithPending(store.getState(), body)).toEqual({
      numFmt: { kind: 'text' },
      bold: true,
    });
  });

  it('ignores pending formats whose merge target is unrelated', () => {
    const store = createSpreadsheetStore();
    const anchor = { sheet: 0, row: 0, col: 0 };
    const body = { sheet: 0, row: 0, col: 1 };
    const otherAnchor = { sheet: 1, row: 0, col: 0 };
    mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    mutators.setCellFormat(store, anchor, { fill: '#fff' });
    mutators.setPendingFormat(store, { addr: otherAnchor, format: { bold: true } });

    expect(formatWithPending(store.getState(), body)).toEqual({ fill: '#fff' });
    expect(sameFormatTarget(store.getState(), body, otherAnchor)).toBe(false);
  });
});
