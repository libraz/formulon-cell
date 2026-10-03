import { addrKey, WorkbookHandle } from '../../../../../src/engine/workbook-handle.js';
import type { SpreadsheetStore } from '../../../../../src/store/store.js';

export const newWb = (): Promise<WorkbookHandle> =>
  WorkbookHandle.createDefault({ preferStub: true });

export const seedAndMirror = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  cells: Array<{ row: number; col: number; value: number | string; formula?: string }>,
): void => {
  store.setState((s) => {
    const map = new Map(s.data.cells);
    for (const c of cells) {
      const addr = { sheet: 0, row: c.row, col: c.col };
      if (c.formula) {
        wb.setFormula(addr, c.formula);
        map.set(addrKey(addr), {
          value:
            typeof c.value === 'number'
              ? { kind: 'number', value: c.value }
              : { kind: 'text', value: c.value },
          formula: c.formula,
        });
      } else if (typeof c.value === 'number') {
        wb.setNumber(addr, c.value);
        map.set(addrKey(addr), { value: { kind: 'number', value: c.value }, formula: null });
      } else {
        wb.setText(addr, c.value);
        map.set(addrKey(addr), { value: { kind: 'text', value: c.value }, formula: null });
      }
    }
    return { ...s, data: { ...s.data, cells: map } };
  });
  wb.recalc();
};

export const setActive = (store: SpreadsheetStore, row: number, col: number, sheet = 0): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      active: { sheet, row, col },
      anchor: { sheet, row, col },
      range: { sheet, r0: row, c0: col, r1: row, c1: col },
    },
  }));
};

export const setSelection = (
  store: SpreadsheetStore,
  r0: number,
  c0: number,
  r1: number,
  c1: number,
  sheet = 0,
): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      active: { sheet, row: r0, col: c0 },
      anchor: { sheet, row: r0, col: c0 },
      range: { sheet, r0, c0, r1, c1 },
      extraRanges: [],
    },
  }));
};

export const num = (wb: WorkbookHandle, row: number, col: number): number => {
  const v = wb.getValue({ sheet: 0, row, col });
  return v.kind === 'number' ? v.value : Number.NaN;
};

export function assertSnap<T>(s: T | null): asserts s is T {
  if (s === null) throw new Error('expected snapshot');
}
