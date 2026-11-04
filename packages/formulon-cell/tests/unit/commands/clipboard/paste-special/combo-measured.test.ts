import { beforeEach, describe, expect, it } from 'vitest';
import { pasteSpecial } from '../../../../../src/commands/clipboard/paste-special.js';
import { captureSnapshot } from '../../../../../src/commands/clipboard/snapshot.js';
import { transposeFormulaRefs } from '../../../../../src/commands/clipboard/transpose-refs.js';
import { addrKey, type WorkbookHandle } from '../../../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../../src/store/store.js';
import { assertSnap, newWb, seedAndMirror, setSelection } from './fixtures.js';

// Expectations are the Excel results measured through the object model for
// copy/paste combinations (transpose, merged targets, formats tiling).

const opts = (over: Partial<Parameters<typeof pasteSpecial>[4]> = {}) => ({
  what: 'all' as const,
  operation: 'none' as const,
  skipBlanks: false,
  transpose: false,
  ...over,
});

const merges = (store: SpreadsheetStore) =>
  [...store.getState().merges.byAnchor.values()].map((m) => [m.r0, m.c0, m.r1, m.c1]);
const numFmt = (store: SpreadsheetStore, row: number, col: number) =>
  store.getState().format.formats.get(addrKey({ sheet: 0, row, col }))?.numFmt;

describe('transposed paste of relative references', () => {
  it('swaps the relative offset of a reference outside the copied block', () => {
    // A6 = D6+1 sits 3 columns right of its cell; transposed it sits 3 rows below.
    const block = { sheet: 0, r0: 5, c0: 0, r1: 5, c1: 2 };
    expect(
      transposeFormulaRefs(
        '=D6+1',
        block,
        { row: 5, col: 0 },
        { row: 25, col: 5 },
        { row: 25, col: 5 },
      ),
    ).toBe('=F29+1');
    expect(
      transposeFormulaRefs(
        '=D6*2',
        block,
        { row: 5, col: 1 },
        { row: 26, col: 5 },
        { row: 25, col: 5 },
      ),
    ).toBe('=F29*2');
  });

  it('keeps absolute parts and swaps only the relative axes', () => {
    const block = { sheet: 0, r0: 9, c0: 0, r1: 11, c1: 0 };
    expect(
      transposeFormulaRefs(
        '=D11+$D$2',
        block,
        { row: 10, col: 0 },
        { row: 19, col: 6 },
        { row: 19, col: 5 },
      ),
    ).toBe('=G23+$D$2');
  });
});

describe('paste onto merged targets', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  it('Formats of a single cell onto a merge unmerges it and fills the footprint', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 1 }]);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { numFmt: { kind: 'fixed', decimals: 2 } },
    );
    mutators.mergeRange(store, { sheet: 0, r0: 13, c0: 5, r1: 13, c1: 6 });
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    assertSnap(snap);
    setSelection(store, 13, 5, 13, 5);
    expect(
      pasteSpecial(store.getState(), store, wb, snap, opts({ what: 'formats' })),
    ).not.toBeNull();
    expect(merges(store)).toEqual([]);
    expect(numFmt(store, 13, 5)).toEqual({ kind: 'fixed', decimals: 2 });
    expect(numFmt(store, 13, 6)).toEqual({ kind: 'fixed', decimals: 2 });
  });

  it('All tiles across a selection that cuts merges, expanding it to the whole merges', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 0, col: 1, value: 2 },
      { row: 1, col: 0, value: 3 },
      { row: 1, col: 1, value: 4 },
    ]);
    // F14:G14, F16:F17, H15:I16 (zero-based rows 13.., cols 5..).
    for (const m of [
      { sheet: 0, r0: 13, c0: 5, r1: 13, c1: 6 },
      { sheet: 0, r0: 15, c0: 5, r1: 16, c1: 5 },
      { sheet: 0, r0: 14, c0: 7, r1: 15, c1: 8 },
    ]) {
      mutators.mergeRange(store, m);
    }
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    assertSnap(snap);
    setSelection(store, 13, 5, 14, 8);
    const result = pasteSpecial(store.getState(), store, wb, snap, opts());
    expect(result?.writtenRange).toEqual({ sheet: 0, r0: 13, c0: 5, r1: 16, c1: 8 });
    expect(merges(store)).toEqual([]);
    expect(wb.getValue({ sheet: 0, row: 15, col: 7 })).toEqual({ kind: 'number', value: 1 });
    expect(wb.getValue({ sheet: 0, row: 16, col: 8 })).toEqual({ kind: 'number', value: 4 });
  });

  it('Values that would cut a merge are refused and change nothing', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 1, col: 0, value: 2 },
      { row: 2, col: 0, value: 3 },
    ]);
    mutators.mergeRange(store, { sheet: 0, r0: 13, c0: 5, r1: 13, c1: 6 });
    mutators.mergeRange(store, { sheet: 0, r0: 15, c0: 5, r1: 16, c1: 5 });
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 });
    assertSnap(snap);
    setSelection(store, 13, 5, 16, 6);
    expect(pasteSpecial(store.getState(), store, wb, snap, opts({ what: 'values' }))).toBeNull();
    expect(merges(store)).toHaveLength(2);
  });
});

describe('Formats paste tiling', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  it('wraps the source across a selection that is not a multiple of it', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 1 }]);
    const fmts = [
      { kind: 'fixed' as const, decimals: 2 },
      { kind: 'fixed' as const, decimals: 0 },
      { kind: 'percent' as const, decimals: 1 },
    ];
    fmts.forEach((numFmt, col) => {
      mutators.setCellFormat(store, { sheet: 0, row: 0, col }, { numFmt });
    });
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });
    assertSnap(snap);
    setSelection(store, 7, 5, 8, 8);
    const result = pasteSpecial(store.getState(), store, wb, snap, opts({ what: 'formats' }));
    expect(result?.writtenRange).toEqual({ sheet: 0, r0: 7, c0: 5, r1: 8, c1: 8 });
    for (const row of [7, 8]) {
      expect(numFmt(store, row, 5)).toEqual(fmts[0]);
      expect(numFmt(store, row, 7)).toEqual(fmts[2]);
      expect(numFmt(store, row, 8)).toEqual(fmts[0]);
    }
  });

  it('crops a merged source to a smaller selection without adding a merge', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 1 }]);
    mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    assertSnap(snap);
    setSelection(store, 1, 5, 2, 5);
    const result = pasteSpecial(store.getState(), store, wb, snap, opts({ what: 'formats' }));
    expect(result?.writtenRange).toEqual({ sheet: 0, r0: 1, c0: 5, r1: 2, c1: 5 });
    expect(merges(store)).toEqual([[0, 0, 0, 1]]);
  });
});
