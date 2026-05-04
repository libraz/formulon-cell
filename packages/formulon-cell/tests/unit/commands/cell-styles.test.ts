import { describe, expect, it } from 'vitest';
import { applyCellStyle, CELL_STYLES, getCellStyle } from '../../../src/commands/cell-styles.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/index.js';

describe('CELL_STYLES', () => {
  it('contains the Excel-flavored presets', () => {
    const ids = CELL_STYLES.map((s) => s.id);
    expect(ids).toContain('normal');
    expect(ids).toContain('heading1');
    expect(ids).toContain('good');
    expect(ids).toContain('currency');
  });
});

describe('getCellStyle', () => {
  it('returns the matching def', () => {
    expect(getCellStyle('good')?.format.fill).toBe('#c6efce');
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
    const fmtCorner = store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 1 }));
    expect(fmtCorner?.fill).toBe('#c6efce');
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
    }
  });

  it('no-ops on unknown style id', () => {
    const store = createSpreadsheetStore();
    const before = store.getState();
    // @ts-expect-error — testing invalid id rejection
    applyCellStyle(store, null, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, 'not-a-style');
    expect(store.getState()).toBe(before);
  });
});
