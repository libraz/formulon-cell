import { beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../../src/commands/history.js';
import {
  deleteCols,
  deleteRows,
  insertCols,
  insertRows,
} from '../../../../src/commands/structure.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import { cellText, newWb } from './fixtures.js';

describe('H-2: structure edits re-point merges / conditional formats / filter', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  it('shifts a merge down when rows are inserted above it', () => {
    mutators.mergeRange(store, { sheet: 0, r0: 3, c0: 1, r1: 4, c1: 2 });
    insertRows(store, wb, null, 1, 2);
    const anchors = Array.from(store.getState().merges.byAnchor.values());
    expect(anchors).toEqual([{ sheet: 0, r0: 5, c0: 1, r1: 6, c1: 2 }]);
    // byCell remap follows the new anchor.
    expect(store.getState().merges.byCell.get('0:6:2')).toBe('0:5:1');
  });

  it('round-trips a native column merge through insert undo and redo', async () => {
    const nativeStore = createSpreadsheetStore();
    const nativeWb = await WorkbookHandle.createDefault();
    const history = new History();
    const original = { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 };
    const shifted = { ...original, c1: 3 };
    mutators.mergeRange(nativeStore, original);
    expect(nativeWb.engineAddMerge(0, original)).toBe(true);

    insertCols(nativeStore, nativeWb, history, 2, 1);
    expect(nativeWb.getMerges(0)).toEqual([shifted]);
    expect(Array.from(nativeStore.getState().merges.byAnchor.values())).toEqual([shifted]);

    history.undo();
    expect(nativeWb.getMerges(0)).toEqual([original]);
    expect(Array.from(nativeStore.getState().merges.byAnchor.values())).toEqual([original]);

    history.redo();
    expect(nativeWb.getMerges(0)).toEqual([shifted]);
    expect(Array.from(nativeStore.getState().merges.byAnchor.values())).toEqual([shifted]);
  });

  it('drops a merge fully inside a deleted row band', () => {
    mutators.mergeRange(store, { sheet: 0, r0: 2, c0: 0, r1: 3, c1: 1 });
    deleteRows(store, wb, null, 2, 2); // removes rows 2..3
    expect(store.getState().merges.byAnchor.size).toBe(0);
  });

  it('shrinks a merge partially covered by a column delete', () => {
    mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 4 });
    deleteCols(store, wb, null, 3, 1); // removes col 3
    const anchors = Array.from(store.getState().merges.byAnchor.values());
    expect(anchors).toEqual([{ sheet: 0, r0: 0, c0: 1, r1: 0, c1: 3 }]);
  });

  it('shifts conditional-format ranges with a row insert', () => {
    mutators.addConditionalRule(store, {
      kind: 'cell-value',
      range: { sheet: 0, r0: 2, c0: 0, r1: 5, c1: 0 },
      op: '>',
      a: 10,
      apply: { bold: true },
    });
    insertRows(store, wb, null, 1, 3);
    expect(store.getState().conditional.rules[0]?.range).toEqual({
      sheet: 0,
      r0: 5,
      c0: 0,
      r1: 8,
      c1: 0,
    });
  });

  it('shifts the autofilter region and per-column criteria on column insert', () => {
    store.setState((s) => ({
      ...s,
      ui: {
        ...s.ui,
        filterRange: { sheet: 0, r0: 0, c0: 1, r1: 4, c1: 3 },
        filterCriteria: [
          { range: { sheet: 0, r0: 0, c0: 1, r1: 4, c1: 3 }, byCol: 2, hiddenValues: ['x'] },
        ],
      },
    }));
    insertCols(store, wb, null, 0, 1); // insert one col at 0 → everything shifts right
    const ui = store.getState().ui;
    expect(ui.filterRange).toEqual({ sheet: 0, r0: 0, c0: 2, r1: 4, c1: 4 });
    expect(ui.filterCriteria[0]?.byCol).toBe(3);
    expect(ui.filterCriteria[0]?.range).toEqual({ sheet: 0, r0: 0, c0: 2, r1: 4, c1: 4 });
  });

  it('shifts the copy marquee so it keeps outlining the copied band', () => {
    mutators.setCopyRange(store, { sheet: 0, r0: 0, c0: 1, r1: 1048575, c1: 1 });
    insertCols(store, wb, null, 0, 1);
    expect(store.getState().ui.copyRange).toEqual({
      sheet: 0,
      r0: 0,
      c0: 2,
      r1: 1048575,
      c1: 2,
    });
    expect(store.getState().ui.copyMode).toBe('copy');

    mutators.setCopyRanges(store, [{ sheet: 0, r0: 3, c0: 0, r1: 3, c1: 16383 }]);
    insertRows(store, wb, null, 0, 2);
    expect(store.getState().ui.copyRanges).toEqual([{ sheet: 0, r0: 5, c0: 0, r1: 5, c1: 16383 }]);
    expect(store.getState().ui.copyMode).toBe('copy');
  });

  it('cancels a cut marquee after successful row and column edits', async () => {
    const cases = [
      {
        label: 'row insert',
        source: { sheet: 0, row: 1, col: 0 },
        range: { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 0 },
        edit: (s: SpreadsheetStore, w: WorkbookHandle) => insertRows(s, w, null, 1, 1),
        expected: { sheet: 0, row: 2, col: 0 },
      },
      {
        label: 'row delete',
        source: { sheet: 0, row: 1, col: 0 },
        range: { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 0 },
        edit: (s: SpreadsheetStore, w: WorkbookHandle) => deleteRows(s, w, null, 0, 1),
        expected: { sheet: 0, row: 0, col: 0 },
      },
      {
        label: 'column insert',
        source: { sheet: 0, row: 0, col: 1 },
        range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        edit: (s: SpreadsheetStore, w: WorkbookHandle) => insertCols(s, w, null, 1, 1),
        expected: { sheet: 0, row: 0, col: 2 },
      },
      {
        label: 'column delete',
        source: { sheet: 0, row: 0, col: 1 },
        range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        edit: (s: SpreadsheetStore, w: WorkbookHandle) => deleteCols(s, w, null, 0, 1),
        expected: { sheet: 0, row: 0, col: 0 },
      },
    ] as const;

    for (const testCase of cases) {
      const caseStore = createSpreadsheetStore();
      const caseWb = await newWb();
      caseWb.setText(testCase.source, 'cut source');
      mutators.setCopyRange(caseStore, testCase.range, 'cut');
      const revision = caseStore.getState().ui.copyRevision ?? 0;

      testCase.edit(caseStore, caseWb);

      expect(
        cellText(caseWb, testCase.expected.sheet, testCase.expected.row, testCase.expected.col),
        testCase.label,
      ).toBe('cut source');
      expect(caseStore.getState().ui.copyRange, testCase.label).toBeNull();
      expect(caseStore.getState().ui.copyRanges, testCase.label).toBeNull();
      expect(caseStore.getState().ui.copyMode, testCase.label).toBeNull();
      expect(caseStore.getState().ui.copyRevision, testCase.label).toBe(revision + 1);
    }
  });

  it('drops the copy marquee when the copied band is deleted', () => {
    mutators.setCopyRange(store, { sheet: 0, r0: 0, c0: 2, r1: 1048575, c1: 2 });
    deleteCols(store, wb, null, 2, 1);
    expect(store.getState().ui.copyRange).toBeNull();
  });

  it('clears the autofilter when its whole region is deleted', () => {
    store.setState((s) => ({
      ...s,
      ui: { ...s.ui, filterRange: { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 1 }, filterCriteria: [] },
    }));
    deleteRows(store, wb, null, 0, 5); // removes rows 0..4
    expect(store.getState().ui.filterRange).toBeNull();
  });

  it('drops a filter criterion whose column is deleted and shrinks the region', () => {
    store.setState((s) => ({
      ...s,
      ui: {
        ...s.ui,
        filterRange: { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 3 },
        filterCriteria: [
          { range: { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 3 }, byCol: 2, hiddenValues: [] },
        ],
      },
    }));
    deleteCols(store, wb, null, 2, 1); // removes col 2 (the criterion's column)
    expect(store.getState().ui.filterCriteria).toHaveLength(0);
    expect(store.getState().ui.filterRange).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 4, c1: 2 });
  });

  it('round-trips merge + conditional shifts through history undo/redo', () => {
    const h = new History();
    wb.attachHistory(h);
    mutators.mergeRange(store, { sheet: 0, r0: 2, c0: 0, r1: 2, c1: 1 });
    insertRows(store, wb, h, 0, 1);
    expect(Array.from(store.getState().merges.byAnchor.values())).toEqual([
      { sheet: 0, r0: 3, c0: 0, r1: 3, c1: 1 },
    ]);
    h.undo();
    expect(Array.from(store.getState().merges.byAnchor.values())).toEqual([
      { sheet: 0, r0: 2, c0: 0, r1: 2, c1: 1 },
    ]);
    h.redo();
    expect(Array.from(store.getState().merges.byAnchor.values())).toEqual([
      { sheet: 0, r0: 3, c0: 0, r1: 3, c1: 1 },
    ]);
  });
});
