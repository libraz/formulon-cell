import { beforeEach, describe, expect, it } from 'vitest';
import { copy } from '../../../../../src/commands/clipboard/copy.js';
import { pasteSpecial } from '../../../../../src/commands/clipboard/paste-special.js';
import { captureSnapshotFromCopyResult } from '../../../../../src/commands/clipboard/snapshot.js';
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

  it('pastes a whole-column copy to a single first-row cell, preserving offsets, tail clearing, and width', () => {
    seedAndMirror(store, wb, [
      { row: 1, col: 3, value: 99 },
      { row: 4, col: 0, value: 7 },
    ]);
    mutators.setCellFormat(store, { sheet: 0, row: 4, col: 0 }, { bold: true });
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        colWidths: new Map([
          [0, 25],
          [3, 10],
        ]),
      },
    }));
    setSelection(store, 0, 0, 1_048_575, 0);
    const copied = copy(store.getState());
    expect(copied?.logicalRange).toEqual({
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 1_048_575,
      c1: 0,
    });
    const snap = copied ? captureSnapshotFromCopyResult(store.getState(), copied) : null;
    assertSnap(snap);
    setActive(store, 0, 3);
    const got = pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    wb.recalc();
    expect(got?.writtenRange).toEqual({ sheet: 0, r0: 0, c0: 3, r1: 1_048_575, c1: 3 });
    expect(num(wb, 4, 3)).toBe(7);
    expect(wb.getValue({ sheet: 0, row: 1, col: 3 })).toEqual({ kind: 'blank' });
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 4, col: 3 })),
    ).toMatchObject({
      bold: true,
    });
    expect(store.getState().layout.colWidths.get(3)).toBe(25);
  });

  it('pastes a whole-row copy to a single first-column cell, preserving offsets, tail clearing, and height', () => {
    seedAndMirror(store, wb, [
      { row: 3, col: 0, value: 99 },
      { row: 0, col: 4, value: 7 },
    ]);
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        rowHeights: new Map([
          [0, 35],
          [3, 20],
        ]),
      },
    }));
    setSelection(store, 0, 0, 0, 16_383);
    const copied = copy(store.getState());
    const snap = copied ? captureSnapshotFromCopyResult(store.getState(), copied) : null;
    assertSnap(snap);
    setActive(store, 3, 0);
    const got = pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    wb.recalc();
    expect(got?.writtenRange).toEqual({ sheet: 0, r0: 3, c0: 0, r1: 3, c1: 16_383 });
    expect(num(wb, 3, 4)).toBe(7);
    expect(wb.getValue({ sheet: 0, row: 3, col: 0 })).toEqual({ kind: 'blank' });
    expect(store.getState().layout.rowHeights.get(3)).toBe(35);
  });

  it('uses only the source-width footprint for a nonmultiple whole-column selection', () => {
    seedAndMirror(store, wb, [
      { row: 4, col: 0, value: 1 },
      { row: 4, col: 1, value: 2 },
      { row: 4, col: 5, value: 9 },
    ]);
    setSelection(store, 0, 0, 1_048_575, 1);
    const copied = copy(store.getState());
    const snap = copied ? captureSnapshotFromCopyResult(store.getState(), copied) : null;
    assertSnap(snap);
    setSelection(store, 0, 3, 0, 5);
    const got = pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    wb.recalc();
    expect(got?.writtenRange).toEqual({ sheet: 0, r0: 0, c0: 3, r1: 1_048_575, c1: 4 });
    expect(num(wb, 4, 3)).toBe(1);
    expect(num(wb, 4, 4)).toBe(2);
    expect(num(wb, 4, 5)).toBe(9);
  });

  it('rejects a whole-band paste whose anchor is outside the first row/column', () => {
    seedAndMirror(store, wb, [{ row: 4, col: 0, value: 7 }]);
    setSelection(store, 0, 0, 1_048_575, 0);
    const copied = copy(store.getState());
    const snap = copied ? captureSnapshotFromCopyResult(store.getState(), copied) : null;
    assertSnap(snap);
    setActive(store, 2, 5);
    expect(
      pasteSpecial(store.getState(), store, wb, snap, {
        what: 'all',
        operation: 'none',
        skipBlanks: false,
        transpose: false,
      }),
    ).toBeNull();
    expect(wb.getValue({ sheet: 0, row: 2, col: 5 })).toEqual({ kind: 'blank' });
  });

  it('repeats an exact-multiple whole-column copy with merges, formulas, widths, and tail clearing', () => {
    seedAndMirror(store, wb, [
      { row: 5, col: 2, value: 1 },
      { row: 6, col: 2, value: 4, formula: '=A1+3' },
      { row: 7, col: 7, value: 99 },
    ]);
    const sourceMerge = { sheet: 0, r0: 5, c0: 2, r1: 5, c1: 3 };
    mutators.mergeRange(store, sourceMerge);
    wb.engineAddMerge(0, sourceMerge);
    mutators.setCellFormat(store, { sheet: 0, row: 5, col: 2 }, { bold: true });
    mutators.setColWidth(store, 2, 40);
    mutators.setColWidth(store, 3, 50);
    setSelection(store, 0, 2, 1_048_575, 3);
    const copied = copy(store.getState());
    const snap = copied ? captureSnapshotFromCopyResult(store.getState(), copied) : null;
    assertSnap(snap);

    setSelection(store, 0, 5, 1_048_575, 8);
    const result = pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    wb.recalc();

    expect(result?.writtenRange).toEqual({ sheet: 0, r0: 0, c0: 5, r1: 1_048_575, c1: 8 });
    expect(wb.getValue({ sheet: 0, row: 5, col: 5 })).toEqual({ kind: 'number', value: 1 });
    expect(wb.getValue({ sheet: 0, row: 5, col: 7 })).toEqual({ kind: 'number', value: 1 });
    expect(wb.cellFormula({ sheet: 0, row: 6, col: 5 })).toBe('=D1+3');
    expect(wb.cellFormula({ sheet: 0, row: 6, col: 7 })).toBe('=F1+3');
    expect(wb.getValue({ sheet: 0, row: 7, col: 7 })).toEqual({ kind: 'blank' });
    expect(store.getState().merges.byAnchor.get(addrKey({ sheet: 0, row: 5, col: 5 }))).toEqual({
      sheet: 0,
      r0: 5,
      c0: 5,
      r1: 5,
      c1: 6,
    });
    expect(store.getState().merges.byAnchor.get(addrKey({ sheet: 0, row: 5, col: 7 }))).toEqual({
      sheet: 0,
      r0: 5,
      c0: 7,
      r1: 5,
      c1: 8,
    });
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 5, col: 5 }))).toEqual({
      bold: true,
    });
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 5, col: 7 }))).toEqual({
      bold: true,
    });
    expect(store.getState().layout.colWidths.get(5)).toBe(40);
    expect(store.getState().layout.colWidths.get(6)).toBe(50);
    expect(store.getState().layout.colWidths.get(7)).toBe(40);
    expect(store.getState().layout.colWidths.get(8)).toBe(50);
  });

  it('repeats an exact-multiple whole-row copy with merges, formulas, heights, and tail clearing', () => {
    seedAndMirror(store, wb, [
      { row: 5, col: 3, value: 1 },
      { row: 6, col: 3, value: 4, formula: '=A1+3' },
      { row: 10, col: 7, value: 99 },
    ]);
    const sourceMerge = { sheet: 0, r0: 5, c0: 3, r1: 5, c1: 4 };
    mutators.mergeRange(store, sourceMerge);
    wb.engineAddMerge(0, sourceMerge);
    mutators.setCellFormat(store, { sheet: 0, row: 5, col: 3 }, { italic: true });
    mutators.setRowHeight(store, 5, 35);
    mutators.setRowHeight(store, 6, 40);
    setSelection(store, 5, 0, 6, 16_383);
    const copied = copy(store.getState());
    const snap = copied ? captureSnapshotFromCopyResult(store.getState(), copied) : null;
    assertSnap(snap);

    setSelection(store, 8, 0, 11, 16_383);
    const result = pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    wb.recalc();

    expect(result?.writtenRange).toEqual({ sheet: 0, r0: 8, c0: 0, r1: 11, c1: 16_383 });
    expect(wb.getValue({ sheet: 0, row: 8, col: 3 })).toEqual({ kind: 'number', value: 1 });
    expect(wb.getValue({ sheet: 0, row: 10, col: 3 })).toEqual({ kind: 'number', value: 1 });
    expect(wb.cellFormula({ sheet: 0, row: 9, col: 3 })).toBe('=A4+3');
    expect(wb.cellFormula({ sheet: 0, row: 11, col: 3 })).toBe('=A6+3');
    expect(wb.getValue({ sheet: 0, row: 10, col: 7 })).toEqual({ kind: 'blank' });
    expect(store.getState().merges.byAnchor.get(addrKey({ sheet: 0, row: 8, col: 3 }))).toEqual({
      sheet: 0,
      r0: 8,
      c0: 3,
      r1: 8,
      c1: 4,
    });
    expect(store.getState().merges.byAnchor.get(addrKey({ sheet: 0, row: 10, col: 3 }))).toEqual({
      sheet: 0,
      r0: 10,
      c0: 3,
      r1: 10,
      c1: 4,
    });
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 8, col: 3 }))).toEqual({
      italic: true,
    });
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 10, col: 3 }))).toEqual({
      italic: true,
    });
    expect(store.getState().layout.rowHeights.get(8)).toBe(35);
    expect(store.getState().layout.rowHeights.get(9)).toBe(40);
    expect(store.getState().layout.rowHeights.get(10)).toBe(35);
    expect(store.getState().layout.rowHeights.get(11)).toBe(40);
  });
});
