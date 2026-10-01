import { beforeEach, describe, expect, it, vi } from 'vitest';
import { pasteSpecial } from '../../../../src/commands/clipboard/paste-special.js';
import { captureSnapshot } from '../../../../src/commands/clipboard/snapshot.js';
import { History } from '../../../../src/commands/history.js';
import { addrKey, WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
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

const setActive = (store: SpreadsheetStore, row: number, col: number, sheet = 0): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      active: { sheet, row, col },
      anchor: { sheet, row, col },
      range: { sheet, r0: row, c0: col, r1: row, c1: col },
    },
  }));
};

const setSelection = (
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

const num = (wb: WorkbookHandle, row: number, col: number): number => {
  const v = wb.getValue({ sheet: 0, row, col });
  return v.kind === 'number' ? v.value : Number.NaN;
};

function assertSnap<T>(s: T | null): asserts s is T {
  if (s === null) throw new Error('expected snapshot');
}

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
    assertSnap(snap);
    const got = pasteSpecial(store.getState(), store, wb, snap, {
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

  it('"formulas" mode preserves a formula instead of replacing it with its result', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 5, formula: '=2+3' }]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setActive(store, 3, 3);
    assertSnap(snap);
    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'formulas',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    expect(wb.cellFormula({ sheet: 0, row: 3, col: 3 })).toBe('=2+3');
  });

  it.each(['formulas', 'formulas-and-numfmt'] as const)(
    '%s pastes numeric, text, and boolean constants',
    (what) => {
      seedAndMirror(store, wb, [
        { row: 0, col: 0, value: 7 },
        { row: 0, col: 1, value: 'text' },
      ]);
      wb.setBool({ sheet: 0, row: 0, col: 2 }, true);
      wb.recalc();
      mutators.replaceCells(store, wb.cells(0));
      const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });
      assertSnap(snap);
      setActive(store, 2, 3);
      pasteSpecial(store.getState(), store, wb, snap, {
        what,
        operation: 'none',
        skipBlanks: false,
        transpose: false,
      });
      expect(wb.getValue({ sheet: 0, row: 2, col: 3 })).toEqual({ kind: 'number', value: 7 });
      expect(wb.getValue({ sheet: 0, row: 2, col: 4 })).toEqual({ kind: 'text', value: 'text' });
      expect(wb.getValue({ sheet: 0, row: 2, col: 5 })).toEqual({ kind: 'bool', value: true });
    },
  );

  it('"values" mode pastes a formula cell result instead of the formula', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 5, formula: '=2+3' }]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setActive(store, 3, 3);
    assertSnap(snap);
    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'values',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    wb.recalc();
    expect(wb.cellFormula({ sheet: 0, row: 3, col: 3 })).toBeNull();
    expect(num(wb, 3, 3)).toBe(5);
  });

  it('"formulas" mode shifts relative references from source to destination', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 2, value: 5, formula: '=A1+B$1' }]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 });
    setActive(store, 3, 4);
    assertSnap(snap);
    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'formulas',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    expect(wb.cellFormula({ sheet: 0, row: 3, col: 4 })).toBe('=C4+D$1');
  });

  it('keeps cut formulas verbatim and moves external refs to the pasted range (C-4)', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 10 },
      { row: 0, col: 1, value: 20 },
      { row: 0, col: 2, value: 30, formula: '=A1+B1' },
      { row: 0, col: 3, value: 30, formula: '=C1' },
      { row: 1, col: 0, value: 30, formula: '=A1+B1' },
    ]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 }, 'cut');
    setActive(store, 4, 4);
    assertSnap(snap);
    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });

    expect(wb.cellFormula({ sheet: 0, row: 4, col: 4 })).toBe('=A1+B1');
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 3 })).toBe('=E5');
    expect(wb.cellFormula({ sheet: 0, row: 1, col: 0 })).toBe('=A1+B1');
  });

  it('updates formulas that referenced a cut source cell after paste (C-4)', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 10 },
      { row: 0, col: 1, value: 20, formula: '=A1*2' },
      { row: 0, col: 2, value: 10, formula: '=A1' },
      { row: 1, col: 0, value: 10, formula: '=$A$1' },
    ]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, 'cut');
    setActive(store, 4, 3);
    assertSnap(snap);
    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });

    expect(wb.cellFormula({ sheet: 0, row: 0, col: 1 })).toBe('=D5*2');
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 2 })).toBe('=D5');
    expect(wb.cellFormula({ sheet: 0, row: 1, col: 0 })).toBe('=$D$5');
  });

  it('follows cut formulas across sheets and records undo/redo as one move', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 7 },
      { row: 0, col: 1, value: 7, formula: '=A1' },
      { row: 0, col: 2, value: 7, formula: '=B1' },
    ]);
    const sourceName = wb.sheetName(0);
    expect(wb.addSheet('Target')).toBe(1);
    expect(wb.addSheet('Other')).toBe(2);
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 1 }, { bold: true });
    mutators.setCellFormat(store, { sheet: 1, row: 2, col: 3 }, { italic: true });
    wb.setFormula({ sheet: 1, row: 0, col: 0 }, `=${sourceName}!B1`);
    wb.setFormula({ sheet: 2, row: 0, col: 0 }, `=${sourceName}!B1`);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }, 'cut');
    assertSnap(snap);
    // A sheet switch may leave the active-sheet cache without the source
    // cells. The engine still owns the source payload and must be cleared.
    store.setState((s) => ({
      ...s,
      data: { ...s.data, sheetIndex: 1, cells: new Map() },
    }));
    setActive(store, 2, 3, 1);

    const history = new History();
    wb.attachHistory(history);
    history.begin();
    const result = pasteSpecial(
      store.getState(),
      store,
      wb,
      snap,
      { what: 'all', operation: 'none', skipBlanks: false, transpose: false },
      history,
    );
    history.end();

    expect(result?.writtenRange).toEqual({ sheet: 1, r0: 2, c0: 3, r1: 2, c1: 3 });
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 }).kind).toBe('blank');
    expect(wb.cellFormula({ sheet: 1, row: 2, col: 3 })).toBe(`=${sourceName}!A1`);
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 2 })).toBe('=Target!D3');
    expect(wb.cellFormula({ sheet: 1, row: 0, col: 0 })).toBe('=D3');
    expect(wb.cellFormula({ sheet: 2, row: 0, col: 0 })).toBe('=Target!D3');
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 1 })),
    ).toBeUndefined();
    expect(store.getState().format.formats.get(addrKey({ sheet: 1, row: 2, col: 3 }))).toEqual({
      bold: true,
    });

    expect(history.undo()).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'number', value: 7 });
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 1 })).toBe('=A1');
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 2 })).toBe('=B1');
    expect(wb.cellFormula({ sheet: 1, row: 2, col: 3 })).toBeNull();
    expect(wb.cellFormula({ sheet: 1, row: 0, col: 0 })).toBe(`=${sourceName}!B1`);
    expect(wb.cellFormula({ sheet: 2, row: 0, col: 0 })).toBe(`=${sourceName}!B1`);
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 1 }))).toEqual({
      bold: true,
    });
    expect(store.getState().format.formats.get(addrKey({ sheet: 1, row: 2, col: 3 }))).toEqual({
      italic: true,
    });

    expect(history.redo()).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 }).kind).toBe('blank');
    expect(wb.cellFormula({ sheet: 1, row: 2, col: 3 })).toBe(`=${sourceName}!A1`);
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 2 })).toBe('=Target!D3');
    expect(wb.cellFormula({ sheet: 1, row: 0, col: 0 })).toBe('=D3');
    expect(wb.cellFormula({ sheet: 2, row: 0, col: 0 })).toBe('=Target!D3');
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 1 })),
    ).toBeUndefined();
    expect(store.getState().format.formats.get(addrKey({ sheet: 1, row: 2, col: 3 }))).toEqual({
      bold: true,
    });
  });

  it('rejects cut Paste Special variants before clearing the source', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 7 }]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, 'cut');
    assertSnap(snap);
    setActive(store, 2, 2);

    for (const options of [
      { what: 'values' as const, operation: 'none' as const, skipBlanks: false, transpose: false },
      { what: 'all' as const, operation: 'add' as const, skipBlanks: false, transpose: false },
      { what: 'all' as const, operation: 'none' as const, skipBlanks: true, transpose: false },
      { what: 'all' as const, operation: 'none' as const, skipBlanks: false, transpose: true },
    ]) {
      expect(pasteSpecial(store.getState(), store, wb, snap, options)).toBeNull();
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 7 });
      expect(wb.getValue({ sheet: 0, row: 2, col: 2 }).kind).toBe('blank');
    }
  });

  it('arithmetic operations combine src and dest numerics', () => {
    // Dest cell pre-existing value.
    seedAndMirror(store, wb, [
      { row: 5, col: 5, value: 100 },
      { row: 0, col: 0, value: 7 },
    ]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setActive(store, 5, 5);
    assertSnap(snap);
    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'values',
      operation: 'add',
      skipBlanks: false,
      transpose: false,
    });
    wb.recalc();
    expect(num(wb, 5, 5)).toBe(107);
  });

  it('arithmetic operations use numeric results from formula cells', () => {
    seedAndMirror(store, wb, [
      { row: 5, col: 5, value: 100 },
      { row: 0, col: 0, value: 7, formula: '=3+4' },
    ]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setActive(store, 5, 5);
    assertSnap(snap);
    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'values',
      operation: 'add',
      skipBlanks: false,
      transpose: false,
    });
    wb.recalc();
    expect(num(wb, 5, 5)).toBe(107);
  });

  it('applies the operation by value for a formula source under what:"all" (M-10)', () => {
    // Previously what:"all" routed the formula-source through the formula-paste
    // branch, pasting `=3+4` verbatim and dropping the "add" operation.
    seedAndMirror(store, wb, [
      { row: 5, col: 5, value: 100 },
      { row: 0, col: 0, value: 7, formula: '=3+4' },
    ]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setActive(store, 5, 5);
    assertSnap(snap);
    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'add',
      skipBlanks: false,
      transpose: false,
    });
    wb.recalc();
    expect(num(wb, 5, 5)).toBe(107);
    // The destination must hold a value, not the pasted formula.
    expect(wb.cellFormula({ sheet: 0, row: 5, col: 5 })).toBeNull();
  });

  it('arithmetic operations ignore non-numeric source values', () => {
    seedAndMirror(store, wb, [
      { row: 5, col: 5, value: 100 },
      { row: 0, col: 0, value: 'x' },
    ]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setActive(store, 5, 5);
    assertSnap(snap);
    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'values',
      operation: 'add',
      skipBlanks: false,
      transpose: false,
    });
    wb.recalc();
    expect(num(wb, 5, 5)).toBe(100);
  });

  it('divide by zero writes a static #DIV/0! error value', () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    seedAndMirror(store, wb, [
      { row: 5, col: 5, value: 50 },
      { row: 0, col: 0, value: 0 },
    ]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setActive(store, 5, 5);
    assertSnap(snap);
    const result = pasteSpecial(store.getState(), store, wb, snap, {
      what: 'values',
      operation: 'divide',
      skipBlanks: false,
      transpose: false,
    });
    wb.recalc();
    // Source value of 0 -> divide-by-zero -> static spreadsheet error.
    expect(result?.skippedNonFiniteOperations).toBe(1);
    expect(warn).toHaveBeenCalledWith(expect.stringContaining('static error value'));
    expect(wb.getValue({ sheet: 0, row: 5, col: 5 })).toMatchObject({
      kind: 'error',
      code: 1,
      text: '#DIV/0!',
    });
    warn.mockRestore();
  });

  it('skipBlanks leaves destination cells untouched when source is blank', () => {
    // Source range has one numeric and one blank cell.
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 9 }]);
    // Destination has a value at the would-be-blank position.
    seedAndMirror(store, wb, [{ row: 5, col: 6, value: 99 }]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    setActive(store, 5, 5);
    assertSnap(snap);
    pasteSpecial(store.getState(), store, wb, snap, {
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

  it('All clears a destination value and comment for an unformatted blank source', () => {
    seedAndMirror(store, wb, [{ row: 5, col: 5, value: 99 }]);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 5, col: 5 },
      {
        comment: 'destination note',
        commentAuthor: 'Bob',
      },
    );
    const snap = captureSnapshot(store.getState(), {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 0,
      c1: 0,
    });
    setActive(store, 5, 5);
    assertSnap(snap);

    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });

    expect(wb.getValue({ sheet: 0, row: 5, col: 5 })).toEqual({ kind: 'blank' });
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 5, col: 5 }))?.comment,
    ).toBeUndefined();
  });

  it('skipBlanks preserves a destination value and comment for a blank source', () => {
    seedAndMirror(store, wb, [{ row: 5, col: 5, value: 99 }]);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 5, col: 5 },
      {
        comment: 'destination note',
        commentAuthor: 'Bob',
      },
    );
    const snap = captureSnapshot(store.getState(), {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 0,
      c1: 0,
    });
    setActive(store, 5, 5);
    assertSnap(snap);

    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: true,
      transpose: false,
    });

    expect(wb.getValue({ sheet: 0, row: 5, col: 5 })).toEqual({ kind: 'number', value: 99 });
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 5, col: 5 }))).toEqual({
      comment: 'destination note',
      commentAuthor: 'Bob',
    });
  });

  it('transpose swaps rows and cols', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 0, col: 1, value: 2 },
      { row: 0, col: 2, value: 3 },
    ]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });
    setActive(store, 5, 5);
    assertSnap(snap);
    const got = pasteSpecial(store.getState(), store, wb, snap, {
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

  it('transpose shifts formula references from each original source cell', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 1, formula: '=A1' },
      { row: 0, col: 1, value: 2, formula: '=B1' },
    ]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    setActive(store, 5, 5);
    assertSnap(snap);
    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'formulas',
      operation: 'none',
      skipBlanks: false,
      transpose: true,
    });
    expect(wb.cellFormula({ sheet: 0, row: 5, col: 5 })).toBe('=F6');
    expect(wb.cellFormula({ sheet: 0, row: 6, col: 5 })).toBe('=F7');
  });

  it('"formats" mode copies cell format and skips values', () => {
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 1 }]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setActive(store, 5, 5);
    assertSnap(snap);
    pasteSpecial(store.getState(), store, wb, snap, {
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

  it('"formats" mode clears a stale destination format when the source is unformatted (M-11)', () => {
    // Source A1 carries a value but no format.
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 1 }]);
    // Destination F6 already has bold formatting that must be overwritten.
    mutators.setCellFormat(store, { sheet: 0, row: 5, col: 5 }, { bold: true });
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setActive(store, 5, 5);
    assertSnap(snap);
    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'formats',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    // Excel copies the source's *absence* of formatting, clearing the stale bold.
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 5, col: 5 })),
    ).toBeUndefined();
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
    assertSnap(snap);
    pasteSpecial(store.getState(), store, wb, snap, {
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
    assertSnap(snap);
    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'values',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    const sel = store.getState().selection;
    expect(sel.range).toEqual({ sheet: 0, r0: 7, c0: 8, r1: 7, c1: 9 });
  });

  it('fills an exact multiple selection with a scalar copy', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 7 }]);
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });
    const snap = captureSnapshot(store.getState(), {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 0,
      c1: 0,
    });
    setSelection(store, 3, 4, 5, 6);
    assertSnap(snap);

    const result = pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });

    expect(result?.writtenRange).toEqual({ sheet: 0, r0: 3, c0: 4, r1: 5, c1: 6 });
    for (let row = 3; row <= 5; row += 1) {
      for (let col = 4; col <= 6; col += 1) {
        expect(num(wb, row, col)).toBe(7);
        expect(store.getState().format.formats.get(addrKey({ sheet: 0, row, col }))).toEqual({
          bold: true,
        });
      }
    }
  });

  it('keeps a nonmultiple selection at the active tile footprint', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 0, col: 1, value: 2 },
      { row: 1, col: 0, value: 3 },
      { row: 1, col: 1, value: 4 },
      { row: 5, col: 5, value: 99 },
    ]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    setSelection(store, 3, 3, 5, 5);
    assertSnap(snap);

    const result = pasteSpecial(store.getState(), store, wb, snap, {
      what: 'values',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });

    expect(result?.writtenRange).toEqual({ sheet: 0, r0: 3, c0: 3, r1: 4, c1: 4 });
    expect(num(wb, 3, 3)).toBe(1);
    expect(num(wb, 4, 4)).toBe(4);
    expect(num(wb, 5, 5)).toBe(99);
  });

  it('repeats a transposed tile across an exact multiple selection', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 0, col: 1, value: 2 },
    ]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    setSelection(store, 3, 3, 6, 4);
    assertSnap(snap);

    const result = pasteSpecial(store.getState(), store, wb, snap, {
      what: 'values',
      operation: 'none',
      skipBlanks: false,
      transpose: true,
    });

    expect(result?.writtenRange).toEqual({ sheet: 0, r0: 3, c0: 3, r1: 6, c1: 4 });
    for (let col = 3; col <= 4; col += 1) {
      expect(num(wb, 3, col)).toBe(1);
      expect(num(wb, 4, col)).toBe(2);
      expect(num(wb, 5, col)).toBe(1);
      expect(num(wb, 6, col)).toBe(2);
    }
  });

  it('never repeats a cut into a larger selected footprint', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 0, col: 1, value: 2 },
      { row: 1, col: 0, value: 3 },
      { row: 1, col: 1, value: 4 },
    ]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 }, 'cut');
    setSelection(store, 3, 3, 6, 6);
    assertSnap(snap);

    const result = pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });

    expect(result?.writtenRange).toEqual({ sheet: 0, r0: 3, c0: 3, r1: 4, c1: 4 });
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'blank' });
    expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'blank' });
    expect(wb.getValue({ sheet: 0, row: 3, col: 3 })).toEqual({ kind: 'number', value: 1 });
    expect(wb.getValue({ sheet: 0, row: 4, col: 4 })).toEqual({ kind: 'number', value: 4 });
    expect(wb.getValue({ sheet: 0, row: 5, col: 5 })).toEqual({ kind: 'blank' });
  });

  it('pastes All with source merge topology and preserves the anchor value', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 7 }]);
    const sourceMerge = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    mutators.mergeRange(store, sourceMerge);
    wb.engineAddMerge(0, sourceMerge);
    const snap = captureSnapshot(store.getState(), sourceMerge);
    setActive(store, 3, 3);
    assertSnap(snap);

    const result = pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    wb.recalc();

    expect(result?.writtenRange).toEqual({ sheet: 0, r0: 3, c0: 3, r1: 4, c1: 4 });
    expect(store.getState().merges.byAnchor.get(addrKey({ sheet: 0, row: 3, col: 3 }))).toEqual({
      sheet: 0,
      r0: 3,
      c0: 3,
      r1: 4,
      c1: 4,
    });
    expect(wb.getValue({ sheet: 0, row: 3, col: 3 })).toEqual({ kind: 'number', value: 7 });
  });

  it('pastes a scalar into a selected merge anchor without tearing the merge', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 9 }]);
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });
    const destinationMerge = { sheet: 0, r0: 3, c0: 3, r1: 4, c1: 4 };
    mutators.mergeRange(store, destinationMerge);
    wb.engineAddMerge(0, destinationMerge);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    mutators.setActive(store, { sheet: 0, row: 3, col: 3 });
    assertSnap(snap);

    const result = pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });

    expect(result?.writtenRange).toEqual(destinationMerge);
    expect(store.getState().merges.byAnchor.get(addrKey({ sheet: 0, row: 3, col: 3 }))).toEqual(
      destinationMerge,
    );
    expect(wb.getValue({ sheet: 0, row: 3, col: 3 })).toEqual({ kind: 'number', value: 9 });
    expect(wb.getValue({ sheet: 0, row: 4, col: 4 })).toEqual({ kind: 'blank' });
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 3, col: 3 }))).toEqual({
      bold: true,
    });
    expect(store.getState().format.formats.has(addrKey({ sheet: 0, row: 4, col: 4 }))).toBe(false);
  });

  it('Formats pastes merge topology without values', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 7 }]);
    const sourceMerge = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    mutators.mergeRange(store, sourceMerge);
    wb.engineAddMerge(0, sourceMerge);
    const snap = captureSnapshot(store.getState(), sourceMerge);
    setActive(store, 3, 3);
    assertSnap(snap);

    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'formats',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    wb.recalc();

    expect(store.getState().merges.byAnchor.get(addrKey({ sheet: 0, row: 3, col: 3 }))).toEqual({
      sheet: 0,
      r0: 3,
      c0: 3,
      r1: 4,
      c1: 4,
    });
    expect(wb.getValue({ sheet: 0, row: 3, col: 3 }).kind).toBe('blank');
  });

  it('repeats every source merge into each tile of the destination', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 7 }]);
    const sourceMerge = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    mutators.mergeRange(store, sourceMerge);
    wb.engineAddMerge(0, sourceMerge);
    const snap = captureSnapshot(store.getState(), sourceMerge);
    setSelection(store, 0, 3, 3, 6);
    assertSnap(snap);

    const result = pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });

    expect(result?.writtenRange).toEqual({ sheet: 0, r0: 0, c0: 3, r1: 3, c1: 6 });
    expect([...store.getState().merges.byAnchor.values()]).toEqual([
      sourceMerge,
      { sheet: 0, r0: 0, c0: 3, r1: 1, c1: 4 },
      { sheet: 0, r0: 0, c0: 5, r1: 1, c1: 6 },
      { sheet: 0, r0: 2, c0: 3, r1: 3, c1: 4 },
      { sheet: 0, r0: 2, c0: 5, r1: 3, c1: 6 },
    ]);
    const anchors: [number, number][] = [
      [0, 3],
      [0, 5],
      [2, 3],
      [2, 5],
    ];
    for (const [row, col] of anchors) {
      expect(num(wb, row, col)).toBe(7);
    }
  });

  it('Values does not import source merge topology', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 7 }]);
    const sourceMerge = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    mutators.mergeRange(store, sourceMerge);
    wb.engineAddMerge(0, sourceMerge);
    const snap = captureSnapshot(store.getState(), sourceMerge);
    setActive(store, 3, 3);
    assertSnap(snap);

    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'values',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });

    expect(store.getState().merges.byAnchor.has(addrKey({ sheet: 0, row: 3, col: 3 }))).toBe(false);
    expect(wb.getValue({ sheet: 0, row: 3, col: 3 })).toEqual({ kind: 'number', value: 7 });
  });

  it('rejects a partial destination merge before mutating the grid', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 7 }]);
    const sourceMerge = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    mutators.mergeRange(store, sourceMerge);
    wb.engineAddMerge(0, sourceMerge);
    const destinationMerge = { sheet: 0, r0: 2, c0: 2, r1: 4, c1: 4 };
    mutators.mergeRange(store, destinationMerge);
    wb.engineAddMerge(0, destinationMerge);
    const snap = captureSnapshot(store.getState(), sourceMerge);
    setActive(store, 3, 3);
    assertSnap(snap);

    const result = pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });

    expect(result).toBeNull();
    expect(store.getState().merges.byAnchor.get(addrKey({ sheet: 0, row: 2, col: 2 }))).toEqual(
      destinationMerge,
    );
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 7 });
  });

  it('rejects a merge crossing the boundary of a repeated destination', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 0, col: 1, value: 2 },
      { row: 1, col: 0, value: 3 },
      { row: 1, col: 1, value: 4 },
    ]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    const crossingMerge = { sheet: 0, r0: 3, c0: 6, r1: 4, c1: 7 };
    mutators.mergeRange(store, crossingMerge);
    wb.engineAddMerge(0, crossingMerge);
    setSelection(store, 0, 3, 3, 6);
    assertSnap(snap);

    const result = pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });

    expect(result).toBeNull();
    expect(store.getState().merges.byAnchor.get(addrKey({ sheet: 0, row: 3, col: 6 }))).toEqual(
      crossingMerge,
    );
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 1 });
    expect(wb.getValue({ sheet: 0, row: 3, col: 3 })).toEqual({ kind: 'blank' });
  });

  it('removes a fully-contained destination merge for an unmerged All paste', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 7 }]);
    const destinationMerge = { sheet: 0, r0: 3, c0: 3, r1: 4, c1: 4 };
    mutators.mergeRange(store, destinationMerge);
    wb.engineAddMerge(0, destinationMerge);
    const snap = captureSnapshot(store.getState(), {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 1,
      c1: 1,
    });
    setActive(store, 3, 3);
    assertSnap(snap);

    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });

    expect(store.getState().merges.byAnchor.has(addrKey({ sheet: 0, row: 3, col: 3 }))).toBe(false);
    expect(wb.getValue({ sheet: 0, row: 3, col: 3 })).toEqual({ kind: 'number', value: 7 });
  });

  it('keeps a fully-contained destination merge for a Values paste', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 7 }]);
    const destinationMerge = { sheet: 0, r0: 3, c0: 3, r1: 4, c1: 4 };
    mutators.mergeRange(store, destinationMerge);
    wb.engineAddMerge(0, destinationMerge);
    const snap = captureSnapshot(store.getState(), {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 1,
      c1: 1,
    });
    setActive(store, 3, 3);
    assertSnap(snap);

    expect(() =>
      pasteSpecial(store.getState(), store, wb, snap, {
        what: 'values',
        operation: 'none',
        skipBlanks: false,
        transpose: false,
      }),
    ).not.toThrow();
    expect(store.getState().merges.byAnchor.get(addrKey({ sheet: 0, row: 3, col: 3 }))).toEqual(
      destinationMerge,
    );
  });

  it('defers cut clearing and supports an overlapping move from the snapshot', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 0, col: 1, value: 2 },
    ]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 }, 'cut');
    assertSnap(snap);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 1 });
    setActive(store, 0, 1);

    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    wb.recalc();

    expect(wb.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'number', value: 1 });
    expect(wb.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({ kind: 'number', value: 2 });
  });

  it('moves cut merge topology and source formatting with the payload', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 7 }]);
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });
    const sourceMerge = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    mutators.mergeRange(store, sourceMerge);
    wb.engineAddMerge(0, sourceMerge);
    const snap = captureSnapshot(store.getState(), sourceMerge, 'cut');
    setActive(store, 3, 3);
    assertSnap(snap);

    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    wb.recalc();

    expect(store.getState().merges.byAnchor.has(addrKey({ sheet: 0, row: 0, col: 0 }))).toBe(false);
    expect(store.getState().merges.byAnchor.has(addrKey({ sheet: 0, row: 3, col: 3 }))).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');
    expect(wb.getValue({ sheet: 0, row: 3, col: 3 })).toEqual({ kind: 'number', value: 7 });
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 })),
    ).toBeUndefined();
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 3, col: 3 }))?.bold).toBe(
      true,
    );
  });

  it('undoes a merged copy as one history action', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 7 }]);
    const sourceMerge = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    mutators.mergeRange(store, sourceMerge);
    wb.engineAddMerge(0, sourceMerge);
    const snap = captureSnapshot(store.getState(), sourceMerge);
    assertSnap(snap);
    setActive(store, 3, 3);
    const history = new History();
    wb.attachHistory(history);
    history.begin();
    pasteSpecial(
      store.getState(),
      store,
      wb,
      snap,
      { what: 'all', operation: 'none', skipBlanks: false, transpose: false },
      history,
    );
    history.end();

    expect(history.undo()).toBe(true);
    expect(store.getState().merges.byAnchor.has(addrKey({ sheet: 0, row: 3, col: 3 }))).toBe(false);
    expect(store.getState().merges.byAnchor.get(addrKey({ sheet: 0, row: 0, col: 0 }))).toEqual(
      sourceMerge,
    );
  });

  it('copies All comments and authors with the source payload', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 7 },
      { row: 3, col: 3, value: 11 },
    ]);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        bold: true,
        comment: 'source note',
        commentAuthor: 'Alice',
      },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 3, col: 3 },
      {
        italic: true,
        comment: 'destination note',
        commentAuthor: 'Bob',
      },
    );
    const snap = captureSnapshot(store.getState(), {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 0,
      c1: 0,
    });
    setActive(store, 3, 3);
    assertSnap(snap);

    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });

    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))).toEqual({
      bold: true,
      comment: 'source note',
      commentAuthor: 'Alice',
    });
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 3, col: 3 }))).toEqual({
      bold: true,
      comment: 'source note',
      commentAuthor: 'Alice',
    });
  });

  it('Formats paste preserves destination comment metadata', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 7 }]);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        bold: true,
        comment: 'source note',
        commentAuthor: 'Alice',
      },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 3, col: 3 },
      {
        italic: true,
        comment: 'destination note',
        commentAuthor: 'Bob',
      },
    );
    const snap = captureSnapshot(store.getState(), {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 0,
      c1: 0,
    });
    setActive(store, 3, 3);
    assertSnap(snap);

    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'formats',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });

    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 3, col: 3 }))).toEqual({
      bold: true,
      comment: 'destination note',
      commentAuthor: 'Bob',
    });
  });

  it('cuts All comments and restores them with one undo action', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 7 },
      { row: 3, col: 3, value: 11 },
    ]);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        comment: 'source note',
        commentAuthor: 'Alice',
      },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 3, col: 3 },
      {
        comment: 'destination note',
        commentAuthor: 'Bob',
      },
    );
    const snap = captureSnapshot(
      store.getState(),
      {
        sheet: 0,
        r0: 0,
        c0: 0,
        r1: 0,
        c1: 0,
      },
      'cut',
    );
    setActive(store, 3, 3);
    assertSnap(snap);
    const history = new History();
    history.begin();
    pasteSpecial(
      store.getState(),
      store,
      wb,
      snap,
      { what: 'all', operation: 'none', skipBlanks: false, transpose: false },
      history,
    );
    history.end();

    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.comment,
    ).toBeUndefined();
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 3, col: 3 }))).toEqual({
      comment: 'source note',
      commentAuthor: 'Alice',
    });
    expect(history.undo()).toBe(true);
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))).toEqual({
      comment: 'source note',
      commentAuthor: 'Alice',
    });
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 3, col: 3 }))).toEqual({
      comment: 'destination note',
      commentAuthor: 'Bob',
    });
    expect(history.redo()).toBe(true);
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.comment,
    ).toBeUndefined();
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 3, col: 3 }))).toEqual({
      comment: 'source note',
      commentAuthor: 'Alice',
    });
  });
});
