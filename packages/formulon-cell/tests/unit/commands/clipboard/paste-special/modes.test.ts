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
});
