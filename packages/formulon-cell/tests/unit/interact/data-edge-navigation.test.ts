import { describe, expect, it, vi } from 'vitest';
import { addrKey } from '../../../src/engine/address.js';
import { jumpToDataEdge } from '../../../src/interact/data-edge-navigation.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import type { CellFormat } from '../../../src/store/types.js';

const setCellKeys = (store: ReturnType<typeof createSpreadsheetStore>, keys: string[]): void => {
  store.setState((state) => {
    const cells = new Map(state.data.cells);
    for (const key of keys) cells.set(key, { value: { kind: 'number', value: 1 }, formula: null });
    return { ...state, data: { ...state.data, cells } };
  });
};

const setFormat = (
  store: ReturnType<typeof createSpreadsheetStore>,
  row: number,
  col: number,
  format: CellFormat,
): void => {
  store.setState((state) => {
    const formats = new Map(state.format.formats);
    formats.set(addrKey({ sheet: 0, row, col }), format);
    return { ...state, format: { ...state.format, formats } };
  });
};

describe('data-edge navigation', () => {
  it('jumps to the end of a populated run, then across a gap to the next run', () => {
    const store = createSpreadsheetStore();
    setCellKeys(
      store,
      [1, 2, 3, 4, 7].map((col) => addrKey({ sheet: 0, row: 1, col })),
    );
    const origin = { sheet: 0, row: 1, col: 1 };

    expect(jumpToDataEdge(store.getState(), origin, 0, 1)).toEqual({
      sheet: 0,
      row: 1,
      col: 4,
    });
    expect(jumpToDataEdge(store.getState(), { sheet: 0, row: 1, col: 4 }, 0, 1)).toEqual({
      sheet: 0,
      row: 1,
      col: 7,
    });
  });

  it('uses format-only cells and ignores empty format entries', () => {
    const store = createSpreadsheetStore();
    setFormat(store, 2, 0, {});
    setFormat(store, 5, 0, { hyperlink: 'https://example.test' });

    expect(jumpToDataEdge(store.getState(), { sheet: 0, row: 0, col: 0 }, 1, 0)).toEqual({
      sheet: 0,
      row: 5,
      col: 0,
    });
  });

  it('keeps the nearest reverse-direction sparse target with bounded direct probes', () => {
    const store = createSpreadsheetStore();
    setCellKeys(
      store,
      [850, 900].map((row) => addrKey({ sheet: 0, row, col: 0 })),
    );
    const cells = store.getState().data.cells;
    const has = vi.spyOn(cells, 'has');
    const keys = vi.spyOn(cells, 'keys');

    expect(jumpToDataEdge(store.getState(), { sheet: 0, row: 1000, col: 0 }, -1, 0)).toEqual({
      sheet: 0,
      row: 900,
      col: 0,
    });
    expect(has.mock.calls.length).toBeLessThanOrEqual(35);
    expect(keys).toHaveBeenCalledTimes(1);
  });
});
