import { describe, expect, it } from 'vitest';
import { applyCellStyle, CELL_STYLES, getCellStyle } from '../../../src/commands/cell-styles.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/index.js';

describe('CELL_STYLES', () => {
  it('contains the spreadsheet-flavored presets', () => {
    const ids = CELL_STYLES.map((s) => s.id);
    expect(ids).toContain('normal');
    expect(ids).toContain('heading1');
    expect(ids).toContain('good');
    expect(ids).toContain('checkCell');
    expect(ids).toContain('explanatoryText');
    expect(ids).toContain('accent1');
    expect(ids).toContain('accent6_20');
    expect(ids).toContain('currency');
  });
});

describe('getCellStyle', () => {
  it('returns the matching def', () => {
    expect(getCellStyle('good')?.format.fill).toBe('#c6efce');
    expect(getCellStyle('accent1')?.format.fill).toBe('#4472c4');
    expect(getCellStyle('accent4_20')?.format.fill).toBe('#fff2cc');
  });

  it('returns undefined for unknown ids', () => {
    // @ts-expect-error — testing invalid id rejection
    expect(getCellStyle('not-a-style')).toBeUndefined();
  });
});

describe('applyCellStyle', () => {
  it('writes the named-style fields onto every cell in range', () => {
    const store = createSpreadsheetStore();
    applyCellStyle(store, null, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 }, 'good');
    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.fill).toBe('#c6efce');
    expect(fmt?.color).toBe('#006100');
    expect(fmt?.cellStyle).toBe('good');
    const fmtCorner = store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 1 }));
    expect(fmtCorner?.fill).toBe('#c6efce');
    expect(fmtCorner?.cellStyle).toBe('good');
  });

  it('clears every format field for the "normal" preset', () => {
    const store = createSpreadsheetStore();
    applyCellStyle(store, null, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, 'good');
    applyCellStyle(store, null, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, 'normal');
    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    // Either the entry is gone, or all visible style fields are undefined.
    if (fmt) {
      expect(fmt.fill).toBeUndefined();
      expect(fmt.color).toBeUndefined();
      expect(fmt.bold).toBeUndefined();
      expect(fmt.cellStyle).toBeUndefined();
    }
  });

  it('applies Excel-style accent and explanatory presets', () => {
    const store = createSpreadsheetStore();
    applyCellStyle(store, null, { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 }, 'accent5_20');
    let fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 2 }));
    expect(fmt).toMatchObject({ color: '#1f4e79', fill: '#ddebf7' });

    applyCellStyle(store, null, { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 }, 'explanatoryText');
    fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 2 }));
    expect(fmt).toMatchObject({ color: '#7f7f7f', italic: true });
  });

  it('no-ops on unknown style id', () => {
    const store = createSpreadsheetStore();
    const before = store.getState();
    // @ts-expect-error — testing invalid id rejection
    applyCellStyle(store, null, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, 'not-a-style');
    expect(store.getState()).toBe(before);
  });
});
