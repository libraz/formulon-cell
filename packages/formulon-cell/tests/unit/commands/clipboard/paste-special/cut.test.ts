import { beforeEach, describe, expect, it } from 'vitest';
import { copy } from '../../../../../src/commands/clipboard/copy.js';
import { pasteSpecial } from '../../../../../src/commands/clipboard/paste-special.js';
import {
  captureSnapshot,
  captureSnapshotFromCopyResult,
} from '../../../../../src/commands/clipboard/snapshot.js';
import { History } from '../../../../../src/commands/history.js';
import { addrKey, type WorkbookHandle } from '../../../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../../src/store/store.js';
import { assertSnap, newWb, seedAndMirror, setActive, setSelection } from './fixtures.js';

describe('pasteSpecial', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
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

  it('moves a whole-column cut with logical references intact', () => {
    seedAndMirror(store, wb, [
      { row: 5, col: 2, value: 7, formula: '=C20' },
      { row: 6, col: 2, value: 11, formula: '=A7' },
      { row: 0, col: 4, value: 7, formula: '=C20' },
    ]);
    setSelection(store, 0, 2, 1_048_575, 2);
    const copied = copy(store.getState());
    const snap = copied ? captureSnapshotFromCopyResult(store.getState(), copied, 'cut') : null;
    assertSnap(snap);
    setActive(store, 0, 3);

    const result = pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    wb.recalc();

    expect(result?.writtenRange).toEqual({ sheet: 0, r0: 0, c0: 3, r1: 1_048_575, c1: 3 });
    expect(wb.cellFormula({ sheet: 0, row: 5, col: 2 })).toBeNull();
    expect(wb.cellFormula({ sheet: 0, row: 5, col: 3 })).toBe('=D20');
    expect(wb.cellFormula({ sheet: 0, row: 6, col: 3 })).toBe('=A7');
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 4 })).toBe('=D20');
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
});
