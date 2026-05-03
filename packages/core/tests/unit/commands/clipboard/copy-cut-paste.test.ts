import { beforeEach, describe, expect, it } from 'vitest';
import { copy } from '../../../../src/commands/clipboard/copy.js';
import { cut } from '../../../../src/commands/clipboard/cut.js';
import { pasteTSV } from '../../../../src/commands/clipboard/paste.js';
import { WorkbookHandle, addrKey } from '../../../../src/engine/workbook-handle.js';
import { type SpreadsheetStore, createSpreadsheetStore } from '../../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const seedAndMirror = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  cells: Array<{ row: number; col: number; value: number | string | boolean }>,
): void => {
  store.setState((s) => {
    const map = new Map(s.data.cells);
    for (const c of cells) {
      const addr = { sheet: 0, row: c.row, col: c.col };
      if (typeof c.value === 'number') {
        wb.setNumber(addr, c.value);
        map.set(addrKey(addr), { value: { kind: 'number', value: c.value }, formula: null });
      } else if (typeof c.value === 'boolean') {
        wb.setBool(addr, c.value);
        map.set(addrKey(addr), { value: { kind: 'bool', value: c.value }, formula: null });
      } else {
        wb.setText(addr, c.value);
        map.set(addrKey(addr), { value: { kind: 'text', value: c.value }, formula: null });
      }
    }
    return { ...s, data: { ...s.data, cells: map } };
  });
  wb.recalc();
};

const setRange = (
  store: SpreadsheetStore,
  r0: number,
  c0: number,
  r1: number,
  c1: number,
): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      ...s.selection,
      active: { sheet: 0, row: r0, col: c0 },
      anchor: { sheet: 0, row: r0, col: c0 },
      range: { sheet: 0, r0, c0, r1, c1 },
    },
  }));
};

describe('copy', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  it('returns null for an inverted range', () => {
    store.setState((s) => ({
      ...s,
      selection: {
        ...s.selection,
        range: { sheet: 0, r0: 5, c0: 5, r1: 4, c1: 4 },
      },
    }));
    expect(copy(store.getState())).toBeNull();
  });

  it('returns null for an over-sized range (>1M cells)', () => {
    setRange(store, 0, 0, 1_048_575, 16_383);
    expect(copy(store.getState())).toBeNull();
  });

  it('emits TSV in row-major order with display strings', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 0, col: 1, value: 'two' },
      { row: 1, col: 0, value: true },
      { row: 1, col: 1, value: false },
    ]);
    setRange(store, 0, 0, 1, 1);
    const got = copy(store.getState());
    expect(got?.tsv).toBe('1\ttwo\r\nTRUE\tFALSE');
    expect(got?.range).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
  });

  it('emits empty cells for missing source positions', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 'x' }]);
    setRange(store, 0, 0, 0, 2);
    expect(copy(store.getState())?.tsv).toBe('x\t\t');
  });
});

describe('cut', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  it('produces the same TSV as copy and blanks the source cells', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 0, col: 1, value: 2 },
    ]);
    setRange(store, 0, 0, 0, 1);
    const got = cut(store.getState(), wb);
    expect(got?.tsv).toBe('1\t2');
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 }).kind).toBe('blank');
  });

  it('returns null when the underlying copy fails (inverted range)', () => {
    store.setState((s) => ({
      ...s,
      selection: {
        ...s.selection,
        range: { sheet: 0, r0: 5, c0: 5, r1: 4, c1: 4 },
      },
    }));
    expect(cut(store.getState(), wb)).toBeNull();
  });
});

describe('pasteTSV', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  it('returns null on empty payload', () => {
    expect(pasteTSV(store.getState(), wb, '')).toBeNull();
  });

  it('writes coerced values starting at the active cell', () => {
    setRange(store, 1, 2, 1, 2);
    const got = pasteTSV(store.getState(), wb, 'foo\t42\r\nTRUE\t=A1');
    wb.recalc();
    expect(got?.writtenRange).toEqual({ sheet: 0, r0: 1, c0: 2, r1: 2, c1: 3 });
    const v00 = wb.getValue({ sheet: 0, row: 1, col: 2 });
    expect(v00.kind === 'text' && v00.value).toBe('foo');
    const v01 = wb.getValue({ sheet: 0, row: 1, col: 3 });
    expect(v01.kind === 'number' && v01.value).toBe(42);
    const v10 = wb.getValue({ sheet: 0, row: 2, col: 2 });
    expect(v10.kind === 'bool' && v10.value).toBe(true);
    // Formula cell.
    expect(wb.cellFormula({ sheet: 0, row: 2, col: 3 })).toBe('=A1');
  });

  it('reports the widest row in writtenRange when rows have unequal column counts', () => {
    setRange(store, 0, 0, 0, 0);
    const got = pasteTSV(store.getState(), wb, 'a\nb\tc\td');
    expect(got?.writtenRange).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 2 });
  });
});
