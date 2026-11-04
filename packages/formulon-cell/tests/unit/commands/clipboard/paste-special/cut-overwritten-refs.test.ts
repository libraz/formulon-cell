import { beforeEach, describe, expect, it } from 'vitest';
import { pasteSpecial } from '../../../../../src/commands/clipboard/paste-special.js';
import { captureSnapshot } from '../../../../../src/commands/clipboard/snapshot.js';
import { addrKey, type WorkbookHandle } from '../../../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../../src/store/store.js';
import { assertSnap, newWb, seedAndMirror, setActive } from './fixtures.js';

const ALL = { what: 'all', operation: 'none', skipBlanks: false, transpose: false } as const;

describe('cut paste over referenced cells', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  const cutTo = (
    src: { r0: number; c0: number; r1: number; c1: number },
    row: number,
    col: number,
  ): void => {
    const snap = captureSnapshot(store.getState(), { sheet: 0, ...src }, 'cut');
    setActive(store, row, col);
    assertSnap(snap);
    pasteSpecial(store.getState(), store, wb, snap, ALL);
  };
  const f = (row: number, col: number): string | null => wb.cellFormula({ sheet: 0, row, col });

  it('turns references to overwritten cells into #REF!', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 5 },
      { row: 3, col: 3, value: 100 },
      { row: 3, col: 4, value: 101 },
      { row: 4, col: 3, value: 102 },
      { row: 4, col: 4, value: 103 },
      { row: 8, col: 0, value: 0, formula: '=D4*2' },
      { row: 8, col: 1, value: 0, formula: '=SUM(D4:E5)' },
      { row: 8, col: 2, value: 0, formula: '=SUM(D4:E4)' },
      { row: 8, col: 3, value: 0, formula: '=SUM(D4:F5)' },
    ]);
    cutTo({ r0: 0, c0: 0, r1: 1, c1: 1 }, 3, 3);
    expect(f(8, 0)).toBe('=#REF!*2');
    expect(f(8, 1)).toBe('=SUM(#REF!)');
    expect(f(8, 2)).toBe('=SUM(#REF!)');
    // A range the destination covers only partly is left alone.
    expect(f(8, 3)).toBe('=SUM(D4:F5)');
  });

  it('keeps references into the moved source following the move', () => {
    seedAndMirror(store, wb, [
      { row: 1, col: 1, value: 1 },
      { row: 1, col: 2, value: 2 },
      { row: 2, col: 1, value: 3 },
      { row: 2, col: 2, value: 4 },
      { row: 6, col: 0, value: 0, formula: '=C2*2' },
      { row: 6, col: 1, value: 0, formula: '=SUM(B2:C3)' },
      { row: 6, col: 2, value: 0, formula: '=D4*2' },
    ]);
    // Destination C3:D4 overlaps the source B2:C3.
    cutTo({ r0: 1, c0: 1, r1: 2, c1: 2 }, 2, 2);
    expect(f(6, 0)).toBe('=D3*2');
    expect(f(6, 1)).toBe('=SUM(C3:D4)');
    expect(f(6, 2)).toBe('=#REF!*2');
  });

  it('qualifies #REF! for a reference to cells overwritten on another sheet', () => {
    wb.addSheet('Sheet2');
    seedAndMirror(store, wb, [
      { row: 1, col: 1, value: 14 },
      { row: 3, col: 0, value: 0, formula: '=Sheet2!D5*2' },
      { row: 3, col: 1, value: 0, formula: '=SUM(Sheet2!D5)' },
    ]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 }, 'cut');
    assertSnap(snap);
    store.setState((s) => ({
      ...s,
      selection: { ...s.selection, active: { sheet: 1, row: 4, col: 3 } },
    }));
    pasteSpecial(store.getState(), store, wb, snap, ALL);
    expect(f(3, 0)).toBe('=Sheet2!#REF!*2');
    expect(f(3, 1)).toBe('=SUM(Sheet2!#REF!)');
  });

  it('drops a merge cut and pasted back onto itself', () => {
    seedAndMirror(store, wb, [{ row: 4, col: 1, value: 1 }]);
    mutators.mergeRange(store, { sheet: 0, r0: 4, c0: 1, r1: 4, c1: 2 });
    cutTo({ r0: 4, c0: 1, r1: 4, c1: 2 }, 4, 1);
    expect(store.getState().merges.byAnchor.has(addrKey({ sheet: 0, row: 4, col: 1 }))).toBe(false);
  });
});
