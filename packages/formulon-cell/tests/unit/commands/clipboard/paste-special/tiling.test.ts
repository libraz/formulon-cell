import { beforeEach, describe, expect, it } from 'vitest';
import { pasteSpecial } from '../../../../../src/commands/clipboard/paste-special.js';
import { captureSnapshot } from '../../../../../src/commands/clipboard/snapshot.js';
import { addrKey, type WorkbookHandle } from '../../../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../../src/store/store.js';
import { assertSnap, newWb, num, seedAndMirror, setActive, setSelection } from './fixtures.js';

describe('pasteSpecial', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
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
});
