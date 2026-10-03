import { beforeEach, describe, expect, it } from 'vitest';
import { pasteSpecial } from '../../../../../src/commands/clipboard/paste-special.js';
import { captureSnapshot } from '../../../../../src/commands/clipboard/snapshot.js';
import { History } from '../../../../../src/commands/history.js';
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
});
