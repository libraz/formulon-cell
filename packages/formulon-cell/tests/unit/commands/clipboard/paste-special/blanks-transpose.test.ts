import { beforeEach, describe, expect, it } from 'vitest';
import { pasteSpecial } from '../../../../../src/commands/clipboard/paste-special.js';
import { captureSnapshot } from '../../../../../src/commands/clipboard/snapshot.js';
import { addrKey, type WorkbookHandle } from '../../../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../../src/store/store.js';
import { assertSnap, newWb, num, seedAndMirror, setActive } from './fixtures.js';

describe('pasteSpecial', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
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
});
