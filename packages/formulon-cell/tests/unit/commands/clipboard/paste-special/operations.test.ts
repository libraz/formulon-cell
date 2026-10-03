import { beforeEach, describe, expect, it, vi } from 'vitest';
import { pasteSpecial } from '../../../../../src/commands/clipboard/paste-special.js';
import { captureSnapshot } from '../../../../../src/commands/clipboard/snapshot.js';
import type { WorkbookHandle } from '../../../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../../../src/store/store.js';
import { assertSnap, newWb, num, seedAndMirror, setActive } from './fixtures.js';

describe('pasteSpecial', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
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
});
