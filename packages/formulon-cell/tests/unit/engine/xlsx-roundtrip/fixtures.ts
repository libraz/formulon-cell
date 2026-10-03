import { addrKey, type WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import type { SpreadsheetStore } from '../../../../src/store/store.js';

export const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';

/** A two-column header plus two data rows — the smallest grid `formatAsTable`
 * can name every column of. */
export const seedTableCells = (wb: WorkbookHandle): void => {
  wb.setText({ sheet: 0, row: 0, col: 0 }, 'Region');
  wb.setText({ sheet: 0, row: 0, col: 1 }, 'Amount');
  wb.setText({ sheet: 0, row: 1, col: 0 }, 'East');
  wb.setNumber({ sheet: 0, row: 1, col: 1 }, 10);
  wb.setText({ sheet: 0, row: 2, col: 0 }, 'West');
  wb.setNumber({ sheet: 0, row: 2, col: 1 }, 20);
};

export const seedStoreText = (
  store: SpreadsheetStore,
  row: number,
  col: number,
  value: string,
): void => {
  store.setState((state) => {
    const cells = new Map(state.data.cells);
    cells.set(addrKey({ sheet: 0, row, col }), { value: { kind: 'text', value }, formula: null });
    return { ...state, data: { ...state.data, cells } };
  });
};
