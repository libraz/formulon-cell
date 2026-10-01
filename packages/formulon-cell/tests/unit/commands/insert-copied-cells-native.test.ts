// @vitest-environment node

import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { copy } from '../../../src/commands/clipboard/copy.js';
import { insertCopiedBand } from '../../../src/commands/clipboard/insert-copied-cells.js';
import { captureSnapshotFromCopyResult } from '../../../src/commands/clipboard/snapshot.js';
import { History } from '../../../src/commands/history.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const MAX_ROW = 1_048_575;
const MAX_COL = 16_383;

const colRange = (col: number, sheet = 0) => ({
  sheet,
  r0: 0,
  c0: col,
  r1: MAX_ROW,
  c1: col,
});

const rowRange = (row: number, sheet = 0) => ({
  sheet,
  r0: row,
  c0: 0,
  r1: row,
  c1: MAX_COL,
});

describe('native Insert Copied Cells whole-band parity', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await WorkbookHandle.createDefault();
    expect(wb.isStub).toBe(false);
  });

  afterEach(() => wb.dispose());

  it('moves a cut whole column with stable formulas, notes, and one undo step', () => {
    wb.setNumber({ sheet: 0, row: 0, col: 3 }, 7);
    wb.setFormula({ sheet: 0, row: 1, col: 3 }, '=D1');
    wb.setNumber({ sheet: 0, row: 0, col: 1 }, 2);
    wb.setNumber({ sheet: 0, row: 0, col: 2 }, 3);
    wb.setNumber({ sheet: 0, row: 0, col: 4 }, 5);
    wb.setFormula({ sheet: 0, row: 2, col: 0 }, '=D1');
    wb.setFormula({ sheet: 0, row: 2, col: 5 }, '=$D$1');
    expect(wb.setCommentEntry(0, 100, 3, 'engine-author', 'far note')).toBe(true);
    wb.recalc();
    mutators.replaceCells(store, wb.cells(0));
    mutators.setColWidth(store, 3, 40);
    mutators.selectCol(store, 3);
    const cut = copy(store.getState());
    const snapshot = cut && captureSnapshotFromCopyResult(store.getState(), cut, 'cut');
    expect(snapshot).not.toBeNull();
    if (!snapshot) return;
    expect(snapshot.mode).toBe('cut');
    mutators.setCopyRange(store, colRange(3), 'cut');
    mutators.selectCol(store, 1);

    const history = new History();
    wb.attachHistory(history);
    expect(insertCopiedBand(store, wb, history, snapshot)?.writtenRange).toEqual(colRange(1));
    wb.recalc();

    expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'number', value: 7 });
    expect(wb.cellFormula({ sheet: 0, row: 1, col: 1 })).toBe('=B1');
    expect(wb.cellFormula({ sheet: 0, row: 2, col: 0 })).toBe('=B1');
    expect(wb.cellFormula({ sheet: 0, row: 2, col: 5 })).toBe('=$B$1');
    expect(wb.getValue({ sheet: 0, row: 0, col: 3 })).toEqual({ kind: 'number', value: 3 });
    expect(wb.getComment(0, 100, 1)).toEqual({ author: 'engine-author', text: 'far note' });
    expect(wb.getComment(0, 100, 3)).toBeNull();
    expect(store.getState().layout.colWidths.get(1)).toBe(40);
    expect(store.getState().ui.copyRange).toBeNull();
    expect(history.canUndo()).toBe(true);

    expect(history.undo()).toBe(true);
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 0, col: 3 })).toEqual({ kind: 'number', value: 7 });
    expect(wb.cellFormula({ sheet: 0, row: 1, col: 3 })).toBe('=D1');
    expect(wb.cellFormula({ sheet: 0, row: 2, col: 0 })).toBe('=D1');
    expect(wb.cellFormula({ sheet: 0, row: 2, col: 5 })).toBe('=$D$1');
    expect(wb.getComment(0, 100, 3)).toEqual({ author: 'engine-author', text: 'far note' });
    expect(wb.getComment(0, 100, 1)).toBeNull();
    expect(history.redo()).toBe(true);
    wb.recalc();
    expect(wb.cellFormula({ sheet: 0, row: 1, col: 1 })).toBe('=B1');
    expect(wb.getComment(0, 100, 1)).toEqual({ author: 'engine-author', text: 'far note' });
  });

  it.each([false, true])(
    'restores the cut marquee and selection when insertion fails after source deletion (multi-range: %s)',
    (multiRange) => {
      wb.setNumber({ sheet: 0, row: 0, col: 3 }, 7);
      wb.setNumber({ sheet: 0, row: 0, col: 4 }, 8);
      wb.recalc();
      mutators.replaceCells(store, wb.cells(0));
      mutators.selectCol(store, 3);
      const copied = copy(store.getState());
      const snapshot = copied && captureSnapshotFromCopyResult(store.getState(), copied, 'cut');
      expect(snapshot).not.toBeNull();
      if (!snapshot) return;
      mutators.setCopyRange(store, colRange(3), 'cut');
      mutators.selectCol(store, 1);
      const beforeSelection = store.getState().selection;
      const beforeCopyRange = store.getState().ui.copyRange;
      if (multiRange) {
        store.setState((state) => ({ ...state, ui: { ...state.ui, copyRanges: [colRange(3)] } }));
      }
      const beforeCopyRanges = store.getState().ui.copyRanges;
      const beforeCopyRevision = store.getState().ui.copyRevision;

      // The native delete succeeds first. `shiftAnchoredRanges` then clears the
      // cut marquee; use that real store transition to protect the sheet before
      // the following insert, forcing the command's transaction abort path.
      let triggered = false;
      const unsubscribe = store.subscribe((state) => {
        if (!triggered && state.ui.copyMode !== 'cut') {
          triggered = true;
          mutators.setSheetProtected(store, 0, true);
        }
      });
      const history = new History();
      wb.attachHistory(history);
      const result = insertCopiedBand(store, wb, history, snapshot);
      unsubscribe();

      expect(result).toBeNull();
      expect(triggered).toBe(true);
      expect(wb.getValue({ sheet: 0, row: 0, col: 3 })).toEqual({ kind: 'number', value: 7 });
      expect(wb.getValue({ sheet: 0, row: 0, col: 4 })).toEqual({ kind: 'number', value: 8 });
      expect(store.getState().selection).toEqual(beforeSelection);
      expect(store.getState().ui.copyRange).toEqual(beforeCopyRange);
      expect(store.getState().ui.copyRanges).toEqual(beforeCopyRanges);
      expect(store.getState().ui.copyMode).toBe('cut');
      expect(store.getState().ui.copyRevision).toBe(beforeCopyRevision);
      expect(history.canUndo()).toBe(false);
      mutators.setSheetProtected(store, 0, false);
    },
  );

  it('moves a cut whole row before the source and rewrites row references', () => {
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 11);
    wb.setFormula({ sheet: 0, row: 3, col: 0 }, '=A1');
    wb.setFormula({ sheet: 0, row: 4, col: 0 }, '=A4');
    wb.setFormula({ sheet: 0, row: 0, col: 1 }, '=A$4');
    wb.setFormula({ sheet: 0, row: 6, col: 2 }, '=A$4');
    wb.recalc();
    mutators.replaceCells(store, wb.cells(0));
    mutators.setRowHeight(store, 3, 35);
    mutators.selectRow(store, 3);
    const cut = copy(store.getState());
    const snapshot = cut && captureSnapshotFromCopyResult(store.getState(), cut, 'cut');
    expect(snapshot).not.toBeNull();
    if (!snapshot) return;
    mutators.setCopyRange(store, rowRange(3), 'cut');
    mutators.selectRow(store, 1);

    const history = new History();
    wb.attachHistory(history);
    expect(insertCopiedBand(store, wb, history, snapshot)?.writtenRange).toEqual(rowRange(1));
    wb.recalc();
    expect(wb.cellFormula({ sheet: 0, row: 1, col: 0 })).toBe('=A1');
    expect(wb.cellFormula({ sheet: 0, row: 4, col: 0 })).toBe('=A2');
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 1 })).toBe('=A$2');
    expect(wb.cellFormula({ sheet: 0, row: 6, col: 2 })).toBe('=A$2');
    expect(wb.getValue({ sheet: 0, row: 3, col: 0 }).kind).toBe('blank');
    expect(store.getState().layout.rowHeights.get(1)).toBe(35);
    expect(history.undo()).toBe(true);
    wb.recalc();
    expect(wb.cellFormula({ sheet: 0, row: 3, col: 0 })).toBe('=A1');
    expect(wb.cellFormula({ sheet: 0, row: 4, col: 0 })).toBe('=A4');
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 1 })).toBe('=A$4');
    expect(wb.cellFormula({ sheet: 0, row: 6, col: 2 })).toBe('=A$4');
    expect(history.redo()).toBe(true);
    wb.recalc();
    expect(wb.cellFormula({ sheet: 0, row: 1, col: 0 })).toBe('=A1');
  });

  it('cuts a whole column across sheets and follows external references', () => {
    const targetSheet = wb.addSheet('Target');
    const sourceName = wb.sheetName(0);
    const targetName = wb.sheetName(targetSheet);
    wb.setNumber({ sheet: 0, row: 0, col: 1 }, 7);
    wb.setFormula({ sheet: 0, row: 1, col: 1 }, '=B1');
    wb.setFormula({ sheet: 0, row: 0, col: 2 }, `=${sourceName}!B1`);
    wb.setNumber({ sheet: targetSheet, row: 0, col: 4 }, 99);
    wb.recalc();
    mutators.replaceCells(store, wb.cells(0));
    mutators.selectCol(store, 1);
    const cut = copy(store.getState());
    const snapshot = cut && captureSnapshotFromCopyResult(store.getState(), cut, 'cut');
    expect(snapshot).not.toBeNull();
    if (!snapshot) return;
    mutators.setCopyRange(store, colRange(1), 'cut');
    mutators.setSheetIndex(store, targetSheet);
    mutators.replaceCells(store, wb.cells(targetSheet));
    mutators.selectCol(store, 3);

    const history = new History();
    wb.attachHistory(history);
    expect(insertCopiedBand(store, wb, history, snapshot)?.writtenRange).toEqual(
      colRange(3, targetSheet),
    );
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 }).kind).toBe('blank');
    expect(wb.getValue({ sheet: targetSheet, row: 0, col: 3 })).toEqual({
      kind: 'number',
      value: 7,
    });
    expect(wb.cellFormula({ sheet: targetSheet, row: 1, col: 3 })).toBe('=D1');
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 2 })).toBe(`=${targetName}!D1`);
    expect(history.undo()).toBe(true);
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'number', value: 7 });
    expect(wb.cellFormula({ sheet: 0, row: 1, col: 1 })).toBe('=B1');
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 2 })).toBe(`=${sourceName}!B1`);
    expect(wb.getValue({ sheet: targetSheet, row: 0, col: 3 }).kind).toBe('blank');
    expect(history.redo()).toBe(true);
    wb.recalc();
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 2 })).toBe(`=${targetName}!D1`);
  });

  it('rejects a cut source that partially spans a merge and removes a crossed destination merge', () => {
    const partialSourceMerge = { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 2 };
    mutators.mergeRange(store, partialSourceMerge);
    expect(wb.engineAddMerge(0, partialSourceMerge)).toBe(true);
    wb.setText({ sheet: 0, row: 1, col: 1 }, 'partial');
    wb.recalc();
    mutators.replaceCells(store, wb.cells(0));
    mutators.selectCol(store, 1);
    const partialCut = copy(store.getState());
    const partialSnapshot =
      partialCut && captureSnapshotFromCopyResult(store.getState(), partialCut, 'cut');
    expect(partialSnapshot).not.toBeNull();
    if (!partialSnapshot) return;
    mutators.setCopyRange(store, colRange(1), 'cut');
    mutators.selectCol(store, 4);
    const rejectedHistory = new History();
    wb.attachHistory(rejectedHistory);
    expect(insertCopiedBand(store, wb, rejectedHistory, partialSnapshot)).toBeNull();
    expect(wb.getMerges(0)).toContainEqual(partialSourceMerge);
    expect(rejectedHistory.canUndo()).toBe(false);

    mutators.unmergeRange(store, partialSourceMerge);
    expect(wb.engineClearMerges(0)).toBe(true);
    const destinationMerge = { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 3 };
    mutators.mergeRange(store, destinationMerge);
    expect(wb.engineAddMerge(0, destinationMerge)).toBe(true);
    wb.setText({ sheet: 0, row: 0, col: 0 }, 'move');
    wb.recalc();
    mutators.replaceCells(store, wb.cells(0));
    mutators.selectCol(store, 0);
    const cut = copy(store.getState());
    const snapshot = cut && captureSnapshotFromCopyResult(store.getState(), cut, 'cut');
    expect(snapshot).not.toBeNull();
    if (!snapshot) return;
    mutators.setCopyRange(store, colRange(0), 'cut');
    mutators.selectCol(store, 3);
    const history = new History();
    wb.attachHistory(history);
    expect(insertCopiedBand(store, wb, history, snapshot)).not.toBeNull();
    expect(wb.getMerges(0)).toEqual([]);
    expect(history.undo()).toBe(true);
    expect(wb.getMerges(0)).toContainEqual(destinationMerge);
  });

  it('copies a whole column before the source with formulas, width, and one undo step', () => {
    wb.setNumber({ sheet: 0, row: 4, col: 4 }, 7);
    wb.setFormula({ sheet: 0, row: 4, col: 3 }, '=$E5');
    wb.recalc();
    mutators.setColWidth(store, 3, 40);
    mutators.setColWidth(store, 1, 10);
    mutators.replaceCells(store, wb.cells(0));
    mutators.selectCol(store, 3);
    const copied = copy(store.getState());
    expect(copied?.logicalRange).toEqual(colRange(3));
    const snapshot = copied && captureSnapshotFromCopyResult(store.getState(), copied);
    expect(snapshot?.colWidths?.get(0)).toBe(40);
    expect(snapshot).not.toBeNull();
    if (!snapshot) return;
    mutators.setCopyRange(store, colRange(3));
    mutators.selectCol(store, 1);

    const history = new History();
    wb.attachHistory(history);
    const result = insertCopiedBand(store, wb, history, snapshot);

    expect(result?.writtenRange).toEqual(colRange(1));
    wb.recalc();
    expect(wb.cellFormula({ sheet: 0, row: 4, col: 1 })).toBe('=$F5');
    expect(wb.cellFormula({ sheet: 0, row: 4, col: 4 })).toBe('=$F5');
    expect(store.getState().layout.colWidths.get(1)).toBe(40);
    expect(history.canUndo()).toBe(true);
    expect(history.undo()).toBe(true);
    wb.recalc();
    expect(wb.cellFormula({ sheet: 0, row: 4, col: 3 })).toBe('=$E5');
    expect(wb.getValue({ sheet: 0, row: 4, col: 1 }).kind).toBe('blank');
    expect(history.redo()).toBe(true);
    wb.recalc();
    expect(wb.cellFormula({ sheet: 0, row: 4, col: 1 })).toBe('=$F5');
    expect(wb.cellFormula({ sheet: 0, row: 4, col: 4 })).toBe('=$F5');
  });

  it('refreshes a source column after an insertion to its right', () => {
    wb.setFormula({ sheet: 0, row: 0, col: 1 }, '=$F1');
    wb.setNumber({ sheet: 0, row: 0, col: 5 }, 9);
    wb.recalc();
    mutators.replaceCells(store, wb.cells(0));
    mutators.selectCol(store, 1);
    const copied = copy(store.getState());
    const snapshot = copied && captureSnapshotFromCopyResult(store.getState(), copied);
    expect(snapshot).not.toBeNull();
    if (!snapshot) return;
    mutators.setCopyRange(store, colRange(1));
    mutators.selectCol(store, 3);
    const history = new History();
    wb.attachHistory(history);

    expect(insertCopiedBand(store, wb, history, snapshot)).not.toBeNull();
    wb.recalc();
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 1 })).toBe('=$G1');
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 3 })).toBe('=$G1');
    expect(history.undo()).toBe(true);
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 1 })).toBe('=$F1');
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 3 })).toBeNull();
  });

  it('copies a whole row before the source with absolute references and height', () => {
    wb.setNumber({ sheet: 0, row: 4, col: 9 }, 11);
    wb.setFormula({ sheet: 0, row: 3, col: 9 }, '=J$5');
    wb.recalc();
    mutators.setRowHeight(store, 3, 35);
    mutators.replaceCells(store, wb.cells(0));
    mutators.selectRow(store, 3);
    const copied = copy(store.getState());
    const snapshot = copied && captureSnapshotFromCopyResult(store.getState(), copied);
    expect(snapshot?.rowHeights?.get(0)).toBe(35);
    expect(snapshot).not.toBeNull();
    if (!snapshot) return;
    mutators.setCopyRange(store, rowRange(3));
    mutators.selectRow(store, 1);

    const history = new History();
    wb.attachHistory(history);
    expect(insertCopiedBand(store, wb, history, snapshot)?.writtenRange).toEqual(rowRange(1));
    wb.recalc();
    expect(wb.cellFormula({ sheet: 0, row: 1, col: 9 })).toBe('=J$6');
    expect(wb.cellFormula({ sheet: 0, row: 4, col: 9 })).toBe('=J$6');
    expect(store.getState().layout.rowHeights.get(1)).toBe(35);
    expect(history.undo()).toBe(true);
    wb.recalc();
    expect(wb.cellFormula({ sheet: 0, row: 3, col: 9 })).toBe('=J$5');
    expect(wb.getValue({ sheet: 0, row: 1, col: 9 }).kind).toBe('blank');
    expect(history.redo()).toBe(true);
    wb.recalc();
    expect(wb.cellFormula({ sheet: 0, row: 1, col: 9 })).toBe('=J$6');
  });

  it('refreshes a source row after an insertion to its right', () => {
    wb.setFormula({ sheet: 0, row: 1, col: 0 }, '=A$6');
    wb.setNumber({ sheet: 0, row: 5, col: 0 }, 9);
    wb.recalc();
    mutators.replaceCells(store, wb.cells(0));
    mutators.selectRow(store, 1);
    const copied = copy(store.getState());
    const snapshot = copied && captureSnapshotFromCopyResult(store.getState(), copied);
    expect(snapshot).not.toBeNull();
    if (!snapshot) return;
    mutators.setCopyRange(store, rowRange(1));
    mutators.selectRow(store, 3);
    const history = new History();
    wb.attachHistory(history);

    expect(insertCopiedBand(store, wb, history, snapshot)).not.toBeNull();
    wb.recalc();
    expect(wb.cellFormula({ sheet: 0, row: 1, col: 0 })).toBe('=A$7');
    expect(wb.cellFormula({ sheet: 0, row: 3, col: 0 })).toBe('=A$7');
    expect(history.undo()).toBe(true);
    expect(wb.cellFormula({ sheet: 0, row: 1, col: 0 })).toBe('=A$6');
    expect(wb.cellFormula({ sheet: 0, row: 3, col: 0 })).toBeNull();
  });

  it('restores a displaced destination value when the single insert is undone', () => {
    wb.setNumber({ sheet: 0, row: 1, col: 4 }, 7);
    wb.setText({ sheet: 0, row: 3, col: 0 }, 'old');
    wb.recalc();
    mutators.replaceCells(store, wb.cells(0));
    mutators.selectRow(store, 1);
    const copied = copy(store.getState());
    const snapshot = copied && captureSnapshotFromCopyResult(store.getState(), copied);
    expect(snapshot).not.toBeNull();
    if (!snapshot) return;
    mutators.setCopyRange(store, rowRange(1));
    mutators.selectRow(store, 3);
    const history = new History();
    wb.attachHistory(history);

    expect(insertCopiedBand(store, wb, history, snapshot)).not.toBeNull();
    expect(wb.getValue({ sheet: 0, row: 3, col: 0 }).kind).toBe('blank');
    expect(wb.getValue({ sheet: 0, row: 4, col: 0 })).toEqual({ kind: 'text', value: 'old' });
    expect(history.undo()).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 3, col: 0 })).toEqual({ kind: 'text', value: 'old' });
    expect(wb.getValue({ sheet: 0, row: 4, col: 0 }).kind).toBe('blank');
  });

  it('rejects an insertion strictly inside the copied band before mutating the sheet', () => {
    wb.setText({ sheet: 0, row: 0, col: 0 }, 'A');
    wb.setText({ sheet: 0, row: 0, col: 1 }, 'B');
    mutators.replaceCells(store, wb.cells(0));
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: MAX_ROW, c1: 1 });
    const copied = copy(store.getState());
    expect(copied).not.toBeNull();
    const snapshot = copied && captureSnapshotFromCopyResult(store.getState(), copied);
    expect(snapshot).not.toBeNull();
    if (!snapshot) return;
    expect(snapshot.logicalRange).toEqual({ sheet: 0, r0: 0, c0: 0, r1: MAX_ROW, c1: 1 });
    mutators.setCopyRange(store, { sheet: 0, r0: 0, c0: 0, r1: MAX_ROW, c1: 1 });
    mutators.selectCol(store, 1);
    const history = new History();
    wb.attachHistory(history);

    expect(insertCopiedBand(store, wb, history, snapshot)).toBeNull();
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'A' });
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'text', value: 'B' });
    expect(history.canUndo()).toBe(false);
  });

  it('rejects malformed merge metadata before a structural insert', () => {
    wb.setText({ sheet: 0, row: 0, col: 0 }, 'keep');
    mutators.replaceCells(store, wb.cells(0));
    mutators.selectCol(store, 1);
    const snapshot = {
      mode: 'copy' as const,
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      logicalRange: colRange(0),
      rows: 1,
      cols: 1,
      cells: [[{ value: { kind: 'text' as const, value: 'x' }, formula: null, format: undefined }]],
      merges: [{ r0: 0, c0: 0, r1: 1, c1: 0 }],
    };
    const history = new History();
    wb.attachHistory(history);

    expect(insertCopiedBand(store, wb, history, snapshot)).toBeNull();
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'keep' });
    expect(history.canUndo()).toBe(false);

    const outsideLogical = { ...snapshot, range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 } };
    expect(insertCopiedBand(store, wb, history, outsideLogical)).toBeNull();
    const missingMatrix = { ...snapshot, merges: undefined, cells: [] };
    expect(insertCopiedBand(store, wb, history, missingMatrix)).toBeNull();
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'keep' });
    expect(history.canUndo()).toBe(false);
  });

  it('rejects a tail comment that would overflow a native row insert', () => {
    expect(wb.setCommentEntry(0, MAX_ROW, 0, 'author', 'tail')).toBe(true);
    mutators.selectRow(store, MAX_ROW - 1);
    const snapshot = {
      mode: 'copy' as const,
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      logicalRange: rowRange(0),
      rows: 1,
      cols: 1,
      cells: [[{ value: { kind: 'blank' as const }, formula: null, format: undefined }]],
    };
    const history = new History();
    wb.attachHistory(history);

    expect(insertCopiedBand(store, wb, history, snapshot)).toBeNull();
    expect(wb.getComment(0, MAX_ROW, 0)).toEqual({ author: 'author', text: 'tail' });
    expect(history.canUndo()).toBe(false);

    expect(wb.setCommentEntry(0, 0, MAX_COL, 'column-author', 'column-tail')).toBe(true);
    mutators.selectCol(store, MAX_COL - 1);
    const colSnapshot = {
      mode: 'copy' as const,
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      logicalRange: colRange(0),
      rows: 1,
      cols: 1,
      cells: [[{ value: { kind: 'blank' as const }, formula: null, format: undefined }]],
    };
    const colHistory = new History();
    wb.attachHistory(colHistory);
    expect(insertCopiedBand(store, wb, colHistory, colSnapshot)).toBeNull();
    expect(wb.getComment(0, 0, MAX_COL)).toEqual({
      author: 'column-author',
      text: 'column-tail',
    });
    expect(colHistory.canUndo()).toBe(false);
  });

  it('rejects a band that cannot fit the structural axis without clamping', () => {
    const twoRows = {
      mode: 'copy' as const,
      range: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 },
      logicalRange: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: MAX_COL },
      rows: 2,
      cols: 1,
      cells: [
        [{ value: { kind: 'text' as const, value: 'a' }, formula: null, format: undefined }],
        [{ value: { kind: 'text' as const, value: 'b' }, formula: null, format: undefined }],
      ],
    };
    mutators.selectRow(store, MAX_ROW - 1);
    const rowHistory = new History();
    wb.attachHistory(rowHistory);
    expect(insertCopiedBand(store, wb, rowHistory, twoRows)).toBeNull();
    expect(rowHistory.canUndo()).toBe(false);

    const twoCols = {
      ...twoRows,
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      logicalRange: { sheet: 0, r0: 0, c0: 0, r1: MAX_ROW, c1: 1 },
      rows: 1,
      cols: 2,
      cells: [
        [
          { value: { kind: 'text' as const, value: 'a' }, formula: null, format: undefined },
          { value: { kind: 'text' as const, value: 'b' }, formula: null, format: undefined },
        ],
      ],
    };
    mutators.selectCol(store, MAX_COL - 1);
    const colHistory = new History();
    wb.attachHistory(colHistory);
    expect(insertCopiedBand(store, wb, colHistory, twoCols)).toBeNull();
    expect(colHistory.canUndo()).toBe(false);
  });

  it('uses source-sheet merge and width metadata for a cross-sheet insertion', () => {
    const targetSheet = wb.addSheet('Target');
    const sourceMerge = { sheet: 0, r0: 3, c0: 3, r1: 4, c1: 3 };
    mutators.mergeRange(store, sourceMerge);
    expect(wb.engineAddMerge(0, sourceMerge)).toBe(true);
    wb.setText({ sheet: 0, row: 3, col: 3 }, 'source');
    mutators.setColWidth(store, 3, 40);
    mutators.replaceCells(store, wb.cells(0));
    mutators.selectCol(store, 3);
    const copied = copy(store.getState());
    const snapshot = copied && captureSnapshotFromCopyResult(store.getState(), copied);
    expect(snapshot?.merges).toContainEqual({ r0: 0, c0: 0, r1: 1, c1: 0 });
    expect(snapshot?.colWidths?.get(0)).toBe(40);
    expect(snapshot).not.toBeNull();
    if (!snapshot) return;

    mutators.setSheetIndex(store, targetSheet);
    mutators.replaceCells(store, wb.cells(targetSheet));
    mutators.selectCol(store, 1);
    const history = new History();
    wb.attachHistory(history);
    expect(insertCopiedBand(store, wb, history, snapshot)?.writtenRange).toEqual(
      colRange(1, targetSheet),
    );

    expect(wb.getMerges(0)).toEqual([sourceMerge]);
    expect(wb.getMerges(targetSheet)).toContainEqual({
      sheet: targetSheet,
      r0: 3,
      c0: 1,
      r1: 4,
      c1: 1,
    });
    expect(store.getState().layout.colWidths.get(1)).toBe(40);
    expect(history.undo()).toBe(true);
    expect(wb.getMerges(0)).toEqual([sourceMerge]);
    expect(wb.getMerges(targetSheet)).toEqual([]);
  });

  it('moves cross-sheet references when the target band is inserted and restores them on undo', () => {
    const targetSheet = wb.addSheet('Target');
    const otherSheet = wb.addSheet('Other');
    const targetName = wb.sheetName(targetSheet);
    wb.setNumber({ sheet: targetSheet, row: 0, col: 1 }, 5);
    wb.setFormula({ sheet: 0, row: 0, col: 0 }, `=${targetName}!B1`);
    wb.setFormula({ sheet: otherSheet, row: 0, col: 0 }, `=${targetName}!B1`);
    wb.recalc();
    mutators.replaceCells(store, wb.cells(0));
    mutators.selectCol(store, 0);
    const copied = copy(store.getState());
    const snapshot = copied && captureSnapshotFromCopyResult(store.getState(), copied);
    expect(snapshot).not.toBeNull();
    if (!snapshot) return;

    mutators.setSheetIndex(store, targetSheet);
    mutators.replaceCells(store, wb.cells(targetSheet));
    mutators.selectCol(store, 1);
    const history = new History();
    wb.attachHistory(history);
    expect(insertCopiedBand(store, wb, history, snapshot)).not.toBeNull();
    wb.recalc();
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe(`=${targetName}!C1`);
    expect(wb.cellFormula({ sheet: otherSheet, row: 0, col: 0 })).toBe(`=${targetName}!C1`);
    expect(wb.cellFormula({ sheet: targetSheet, row: 0, col: 1 })).toBe(`=${targetName}!D1`);
    expect(wb.getValue({ sheet: targetSheet, row: 0, col: 2 })).toEqual({
      kind: 'number',
      value: 5,
    });

    expect(history.undo()).toBe(true);
    wb.recalc();
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe(`=${targetName}!B1`);
    expect(wb.cellFormula({ sheet: otherSheet, row: 0, col: 0 })).toBe(`=${targetName}!B1`);
    expect(wb.getValue({ sheet: targetSheet, row: 0, col: 1 })).toEqual({
      kind: 'number',
      value: 5,
    });
    expect(wb.cellFormula({ sheet: targetSheet, row: 0, col: 2 })).toBeNull();
  });
});
