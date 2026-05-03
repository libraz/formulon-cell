import { beforeEach, describe, expect, it } from 'vitest';
import { pasteSpecial } from '../../../../src/commands/clipboard/paste-special.js';
import { captureSnapshot } from '../../../../src/commands/clipboard/snapshot.js';
import { WorkbookHandle, addrKey } from '../../../../src/engine/workbook-handle.js';
import {
  type SpreadsheetStore,
  createSpreadsheetStore,
  mutators,
} from '../../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const seedAndMirror = (
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

const setActive = (store: SpreadsheetStore, row: number, col: number): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      active: { sheet: 0, row, col },
      anchor: { sheet: 0, row, col },
      range: { sheet: 0, r0: row, c0: col, r1: row, c1: col },
    },
  }));
};

const num = (wb: WorkbookHandle, row: number, col: number): number => {
  const v = wb.getValue({ sheet: 0, row, col });
  return v.kind === 'number' ? v.value : Number.NaN;
};

describe('pasteSpecial', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  it('writes values via the "values" mode', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 10 },
      { row: 0, col: 1, value: 'hi' },
    ]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    expect(snap).not.toBeNull();
    setActive(store, 5, 5);
    const got = pasteSpecial(store.getState(), store, wb, snap!, {
      what: 'values',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    wb.recalc();
    expect(got?.writtenRange).toEqual({ sheet: 0, r0: 5, c0: 5, r1: 5, c1: 6 });
    expect(num(wb, 5, 5)).toBe(10);
    expect(wb.getValue({ sheet: 0, row: 5, col: 6 })).toEqual({ kind: 'text', value: 'hi' });
  });

  it('"formulas" mode pastes formulas, not values', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 5, formula: '=2+3' }]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setActive(store, 3, 3);
    pasteSpecial(store.getState(), store, wb, snap!, {
      what: 'formulas',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    expect(wb.cellFormula({ sheet: 0, row: 3, col: 3 })).toBe('=2+3');
  });

  it('arithmetic operations combine src and dest numerics', () => {
    // Dest cell pre-existing value.
    seedAndMirror(store, wb, [
      { row: 5, col: 5, value: 100 },
      { row: 0, col: 0, value: 7 },
    ]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setActive(store, 5, 5);
    pasteSpecial(store.getState(), store, wb, snap!, {
      what: 'values',
      operation: 'add',
      skipBlanks: false,
      transpose: false,
    });
    wb.recalc();
    expect(num(wb, 5, 5)).toBe(107);
  });

  it('divide by zero produces NaN result, which is skipped', () => {
    seedAndMirror(store, wb, [
      { row: 5, col: 5, value: 50 },
      { row: 0, col: 0, value: 0 },
    ]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setActive(store, 5, 5);
    pasteSpecial(store.getState(), store, wb, snap!, {
      what: 'values',
      operation: 'divide',
      skipBlanks: false,
      transpose: false,
    });
    wb.recalc();
    // Source value of 0 → divide-by-zero → result skipped, dest unchanged.
    expect(num(wb, 5, 5)).toBe(50);
  });

  it('skipBlanks leaves destination cells untouched when source is blank', () => {
    // Source range has one numeric and one blank cell.
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 9 }]);
    // Destination has a value at the would-be-blank position.
    seedAndMirror(store, wb, [{ row: 5, col: 6, value: 99 }]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    setActive(store, 5, 5);
    pasteSpecial(store.getState(), store, wb, snap!, {
      what: 'values',
      operation: 'none',
      skipBlanks: true,
      transpose: false,
    });
    wb.recalc();
    expect(num(wb, 5, 5)).toBe(9);
    // Untouched.
    expect(num(wb, 5, 6)).toBe(99);
  });

  it('transpose swaps rows and cols', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 0, col: 1, value: 2 },
      { row: 0, col: 2, value: 3 },
    ]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });
    setActive(store, 5, 5);
    const got = pasteSpecial(store.getState(), store, wb, snap!, {
      what: 'values',
      operation: 'none',
      skipBlanks: false,
      transpose: true,
    });
    wb.recalc();
    // 1x3 → 3x1
    expect(got?.writtenRange).toEqual({ sheet: 0, r0: 5, c0: 5, r1: 7, c1: 5 });
    expect(num(wb, 5, 5)).toBe(1);
    expect(num(wb, 6, 5)).toBe(2);
    expect(num(wb, 7, 5)).toBe(3);
  });

  it('"formats" mode copies cell format and skips values', () => {
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 1 }]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setActive(store, 5, 5);
    pasteSpecial(store.getState(), store, wb, snap!, {
      what: 'formats',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    wb.recalc();
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 5, col: 5 }))?.bold).toBe(
      true,
    );
    // No value pasted.
    expect(wb.getValue({ sheet: 0, row: 5, col: 5 }).kind).toBe('blank');
  });

  it('"values-and-numfmt" cherry-picks numFmt without bold', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { bold: true, numFmt: { kind: 'fixed', decimals: 2 } },
    );
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 9 }]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setActive(store, 4, 4);
    pasteSpecial(store.getState(), store, wb, snap!, {
      what: 'values-and-numfmt',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    const dest = store.getState().format.formats.get(addrKey({ sheet: 0, row: 4, col: 4 }));
    expect(dest?.numFmt).toEqual({ kind: 'fixed', decimals: 2 });
    expect(dest?.bold).toBeUndefined();
  });

  it('updates active selection to the written range', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 0, col: 1, value: 2 },
    ]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    setActive(store, 7, 8);
    pasteSpecial(store.getState(), store, wb, snap!, {
      what: 'values',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    const sel = store.getState().selection;
    expect(sel.range).toEqual({ sheet: 0, r0: 7, c0: 8, r1: 7, c1: 9 });
  });
});
