import { beforeEach, describe, expect, it } from 'vitest';
import { fillDestFor, fillRange } from '../../../src/commands/fill.js';
import type { Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const num = (wb: WorkbookHandle, sheet: number, row: number, col: number): number => {
  const v = wb.getValue({ sheet, row, col });
  return v.kind === 'number' ? v.value : Number.NaN;
};

const text = (wb: WorkbookHandle, sheet: number, row: number, col: number): string => {
  const v = wb.getValue({ sheet, row, col });
  return v.kind === 'text' ? v.value : '';
};

/**
 * Mirror writes on `wb` back into the store's data map so fillRange can read
 * the source cells through state. The real app does this via the change-event
 * subscription on WorkbookHandle, but we keep tests engine-only here.
 */
const seedAndMirror = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  cells: Array<{ row: number; col: number; value: number | string }>,
): void => {
  store.setState((s) => {
    const map = new Map(s.data.cells);
    for (const c of cells) {
      const key = `${0}:${c.row}:${c.col}`;
      if (typeof c.value === 'number') {
        wb.setNumber({ sheet: 0, row: c.row, col: c.col }, c.value);
        map.set(key, { value: { kind: 'number', value: c.value }, formula: null });
      } else {
        wb.setText({ sheet: 0, row: c.row, col: c.col }, c.value);
        map.set(key, { value: { kind: 'text', value: c.value }, formula: null });
      }
    }
    return { ...s, data: { ...s.data, cells: map } };
  });
  wb.recalc();
};

describe('fillDestFor', () => {
  it('returns the source unchanged when the target sits inside it', () => {
    const src: Range = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 };
    const got = fillDestFor(src, { row: 1, col: 1 });
    expect(got).toEqual(src);
  });

  it('extends down when row delta dominates', () => {
    const src: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    const got = fillDestFor(src, { row: 5, col: 1 });
    expect(got).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 5, c1: 1 });
  });

  it('extends up when target is above the source', () => {
    const src: Range = { sheet: 0, r0: 5, c0: 0, r1: 6, c1: 1 };
    const got = fillDestFor(src, { row: 0, col: 1 });
    expect(got).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 6, c1: 1 });
  });

  it('extends right when column delta dominates', () => {
    const src: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 };
    const got = fillDestFor(src, { row: 0, col: 5 });
    expect(got).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 5 });
  });

  it('extends left when target is left of the source', () => {
    const src: Range = { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 6 };
    const got = fillDestFor(src, { row: 0, col: 1 });
    expect(got).toEqual({ sheet: 0, r0: 0, c0: 1, r1: 0, c1: 6 });
  });
});

describe('fillRange', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  it('returns false when src and dest are identical', () => {
    const range: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    expect(fillRange(store.getState(), wb, range, range)).toBe(false);
  });

  it('extrapolates a numeric arithmetic series downward', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 10 },
      { row: 1, col: 0, value: 20 },
    ]);
    const src: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 };
    const dest: Range = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 };
    expect(fillRange(store.getState(), wb, src, dest)).toBe(true);
    wb.recalc();
    expect(num(wb, 0, 2, 0)).toBe(30);
    expect(num(wb, 0, 3, 0)).toBe(40);
    expect(num(wb, 0, 4, 0)).toBe(50);
  });

  it('copies a single numeric source cell', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 7 }]);
    const src: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };
    const dest: Range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };
    fillRange(store.getState(), wb, src, dest);
    wb.recalc();
    expect(num(wb, 0, 1, 0)).toBe(7);
    expect(num(wb, 0, 2, 0)).toBe(7);
    expect(num(wb, 0, 3, 0)).toBe(7);
  });

  it('increments trailing integers in text labels', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 'Item 1' }]);
    const src: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };
    const dest: Range = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 };
    fillRange(store.getState(), wb, src, dest);
    wb.recalc();
    expect(text(wb, 0, 1, 0)).toBe('Item 2');
    expect(text(wb, 0, 2, 0)).toBe('Item 3');
  });

  it('increments full-width trailing digits in text labels', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: '項目１' }]);
    const src: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };
    const dest: Range = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 };
    fillRange(store.getState(), wb, src, dest);
    wb.recalc();
    expect(text(wb, 0, 1, 0)).toBe('項目２');
    expect(text(wb, 0, 2, 0)).toBe('項目３');
  });

  it('preserves zero padding when incrementing text labels', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 'No. ００１' }]);
    const src: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };
    const dest: Range = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 };
    fillRange(store.getState(), wb, src, dest);
    wb.recalc();
    expect(text(wb, 0, 1, 0)).toBe('No. ００２');
    expect(text(wb, 0, 2, 0)).toBe('No. ００３');
  });

  it('extends right with a numeric series', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 0, col: 1, value: 3 },
    ]);
    const src: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 };
    const dest: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 3 };
    fillRange(store.getState(), wb, src, dest);
    wb.recalc();
    expect(num(wb, 0, 0, 2)).toBe(5);
    expect(num(wb, 0, 0, 3)).toBe(7);
  });

  it('extends up by reversing the projection', () => {
    seedAndMirror(store, wb, [
      { row: 5, col: 0, value: 50 },
      { row: 6, col: 0, value: 60 },
    ]);
    const src: Range = { sheet: 0, r0: 5, c0: 0, r1: 6, c1: 0 };
    const dest: Range = { sheet: 0, r0: 3, c0: 0, r1: 6, c1: 0 };
    fillRange(store.getState(), wb, src, dest);
    wb.recalc();
    // Going upward, the step is -10; row 4 = 40, row 3 = 30.
    expect(num(wb, 0, 4, 0)).toBe(40);
    expect(num(wb, 0, 3, 0)).toBe(30);
  });

  it('cycles short English weekday names (Mon → Tue → ...) past the end of the list', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 'Mon' },
      { row: 1, col: 0, value: 'Tue' },
    ]);
    const src: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 };
    const dest: Range = { sheet: 0, r0: 0, c0: 0, r1: 7, c1: 0 };
    fillRange(store.getState(), wb, src, dest);
    wb.recalc();
    expect(text(wb, 0, 2, 0)).toBe('Wed');
    expect(text(wb, 0, 5, 0)).toBe('Sat');
    expect(text(wb, 0, 6, 0)).toBe('Sun');
    expect(text(wb, 0, 7, 0)).toBe('Mon');
  });

  it('cycles English month names', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 'Jan' },
      { row: 1, col: 0, value: 'Feb' },
    ]);
    const src: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 };
    const dest: Range = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 };
    fillRange(store.getState(), wb, src, dest);
    wb.recalc();
    expect(text(wb, 0, 2, 0)).toBe('Mar');
    expect(text(wb, 0, 3, 0)).toBe('Apr');
    expect(text(wb, 0, 4, 0)).toBe('May');
  });

  it('preserves source casing for English custom lists', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 'jan' },
      { row: 1, col: 0, value: 'feb' },
    ]);
    const src: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 };
    const dest: Range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };
    fillRange(store.getState(), wb, src, dest);
    wb.recalc();
    expect(text(wb, 0, 2, 0)).toBe('mar');
    expect(text(wb, 0, 3, 0)).toBe('apr');
  });

  it('preserves uppercase source casing for English custom lists', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 'MON' }]);
    const src: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };
    const dest: Range = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 };
    fillRange(store.getState(), wb, src, dest);
    wb.recalc();
    expect(text(wb, 0, 1, 0)).toBe('TUE');
    expect(text(wb, 0, 2, 0)).toBe('WED');
  });

  it('cycles Japanese weekday characters (日 → 月 → ...)', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: '日' }]);
    const src: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };
    const dest: Range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };
    fillRange(store.getState(), wb, src, dest);
    wb.recalc();
    expect(text(wb, 0, 1, 0)).toBe('月');
    expect(text(wb, 0, 2, 0)).toBe('火');
    expect(text(wb, 0, 3, 0)).toBe('水');
  });

  it('cycles parenthesized Japanese weekday labels', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: '(土)' }]);
    const src: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };
    const dest: Range = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 };
    fillRange(store.getState(), wb, src, dest);
    wb.recalc();
    expect(text(wb, 0, 1, 0)).toBe('(日)');
    expect(text(wb, 0, 2, 0)).toBe('(月)');
  });

  it('cycles roman quarter labels', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 'QIII' }]);
    const src: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };
    const dest: Range = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 };
    fillRange(store.getState(), wb, src, dest);
    wb.recalc();
    expect(text(wb, 0, 1, 0)).toBe('QIV');
    expect(text(wb, 0, 2, 0)).toBe('QI');
  });

  it('cycles Japanese abbreviated weekday labels', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: '金曜' }]);
    const src: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };
    const dest: Range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };
    fillRange(store.getState(), wb, src, dest);
    wb.recalc();
    expect(text(wb, 0, 1, 0)).toBe('土曜');
    expect(text(wb, 0, 2, 0)).toBe('日曜');
    expect(text(wb, 0, 3, 0)).toBe('月曜');
  });

  it('preserves full-width minus signs when incrementing text labels', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: '項目－２' }]);
    const src: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };
    const dest: Range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };
    fillRange(store.getState(), wb, src, dest);
    wb.recalc();
    expect(text(wb, 0, 1, 0)).toBe('項目－１');
    expect(text(wb, 0, 2, 0)).toBe('項目０');
    expect(text(wb, 0, 3, 0)).toBe('項目１');
  });

  it('cycles Japanese month labels', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: '11月' }]);
    const src: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };
    const dest: Range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };
    fillRange(store.getState(), wb, src, dest);
    wb.recalc();
    expect(text(wb, 0, 1, 0)).toBe('12月');
    expect(text(wb, 0, 2, 0)).toBe('1月');
    expect(text(wb, 0, 3, 0)).toBe('2月');
  });

  it('cycles Q1/Q2/Q3/Q4 quarter labels', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 'Q1' }]);
    const src: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };
    const dest: Range = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 };
    fillRange(store.getState(), wb, src, dest);
    wb.recalc();
    expect(text(wb, 0, 1, 0)).toBe('Q2');
    expect(text(wb, 0, 2, 0)).toBe('Q3');
    expect(text(wb, 0, 3, 0)).toBe('Q4');
    // Wraps back to Q1.
    expect(text(wb, 0, 4, 0)).toBe('Q1');
  });

  it('cycles full-width quarter labels', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 'Q４' }]);
    const src: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };
    const dest: Range = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 };
    fillRange(store.getState(), wb, src, dest);
    wb.recalc();
    expect(text(wb, 0, 1, 0)).toBe('Q１');
    expect(text(wb, 0, 2, 0)).toBe('Q２');
  });

  it('copyOnly suppresses series extrapolation and tiles the source', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 1, col: 0, value: 2 },
    ]);
    const src: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 };
    const dest: Range = { sheet: 0, r0: 0, c0: 0, r1: 5, c1: 0 };
    fillRange(store.getState(), wb, src, dest, { copyOnly: true });
    wb.recalc();
    // Tiled: 1, 2, 1, 2, 1, 2 — not 1, 2, 3, 4, 5, 6.
    expect(num(wb, 0, 2, 0)).toBe(1);
    expect(num(wb, 0, 3, 0)).toBe(2);
    expect(num(wb, 0, 4, 0)).toBe(1);
    expect(num(wb, 0, 5, 0)).toBe(2);
  });

  it('shifts relative refs when filling formulas down', () => {
    // Source cell C1 = =A1+B1. Filling down to C4 should produce
    //   C2 = =A2+B2, C3 = =A3+B3, C4 = =A4+B4.
    wb.setFormula({ sheet: 0, row: 0, col: 2 }, '=A1+B1');
    store.setState((s) => {
      const map = new Map(s.data.cells);
      map.set('0:0:2', { value: { kind: 'number', value: 0 }, formula: '=A1+B1' });
      return { ...s, data: { ...s.data, cells: map } };
    });
    const src: Range = { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 };
    const dest: Range = { sheet: 0, r0: 0, c0: 2, r1: 3, c1: 2 };
    fillRange(store.getState(), wb, src, dest);
    expect(wb.cellFormula({ sheet: 0, row: 1, col: 2 })).toBe('=A2+B2');
    expect(wb.cellFormula({ sheet: 0, row: 2, col: 2 })).toBe('=A3+B3');
    expect(wb.cellFormula({ sheet: 0, row: 3, col: 2 })).toBe('=A4+B4');
  });

  it('shifts relative refs when filling formulas right', () => {
    wb.setFormula({ sheet: 0, row: 0, col: 0 }, '=B1+1');
    store.setState((s) => {
      const map = new Map(s.data.cells);
      map.set('0:0:0', { value: { kind: 'number', value: 0 }, formula: '=B1+1' });
      return { ...s, data: { ...s.data, cells: map } };
    });
    const src: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };
    const dest: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 };
    fillRange(store.getState(), wb, src, dest);
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 1 })).toBe('=C1+1');
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 2 })).toBe('=D1+1');
  });

  it('respects $-locked refs when filling down', () => {
    // =$A1 — col anchored, row relative. Fill down: $A2, $A3, $A4.
    wb.setFormula({ sheet: 0, row: 0, col: 1 }, '=$A1+A$1');
    store.setState((s) => {
      const map = new Map(s.data.cells);
      map.set('0:0:1', { value: { kind: 'number', value: 0 }, formula: '=$A1+A$1' });
      return { ...s, data: { ...s.data, cells: map } };
    });
    const src: Range = { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 };
    const dest: Range = { sheet: 0, r0: 0, c0: 1, r1: 2, c1: 1 };
    fillRange(store.getState(), wb, src, dest);
    expect(wb.cellFormula({ sheet: 0, row: 1, col: 1 })).toBe('=$A2+A$1');
    expect(wb.cellFormula({ sheet: 0, row: 2, col: 1 })).toBe('=$A3+A$1');
  });
});
