import { describe, expect, it } from 'vitest';
import { History } from '../../../../src/commands/history.js';
import {
  deleteCols,
  deleteRows,
  insertCols,
  insertRows,
} from '../../../../src/commands/structure.js';
import type { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';

describe('insertRows / deleteRows / insertCols / deleteCols engine path', () => {
  interface FakeCell {
    addr: { sheet: number; row: number; col: number };
    value: { kind: 'number'; value: number };
    formula: string | null;
  }
  interface FakeWb {
    insertR: { sheet: number; row: number; count: number }[];
    deleteR: { sheet: number; row: number; count: number }[];
    insertC: { sheet: number; col: number; count: number }[];
    deleteC: { sheet: number; col: number; count: number }[];
    recalcs: number;
    setNumberCalls: { sheet: number; row: number; col: number; value: number }[];
  }
  const makeWb = (
    cells: FakeCell[] = [],
    nativeOk = true,
  ): { wb: WorkbookHandle; calls: FakeWb } => {
    const calls: FakeWb = {
      insertR: [],
      deleteR: [],
      insertC: [],
      deleteC: [],
      recalcs: 0,
      setNumberCalls: [],
    };
    const wb = {
      capabilities: { insertDeleteRowsCols: true },
      cells: function* (sheet: number) {
        for (const c of cells) if (c.addr.sheet === sheet) yield c;
      },
      recalc: () => {
        calls.recalcs += 1;
      },
      recalcAuto: () => {
        calls.recalcs += 1;
      },
      withBatchedRecalc: <T>(fn: () => T): T => fn(),
      engineInsertRows: (sheet: number, row: number, count: number) => {
        calls.insertR.push({ sheet, row, count });
        return nativeOk;
      },
      engineDeleteRows: (sheet: number, row: number, count: number) => {
        calls.deleteR.push({ sheet, row, count });
        return nativeOk;
      },
      engineInsertCols: (sheet: number, col: number, count: number) => {
        calls.insertC.push({ sheet, col, count });
        return nativeOk;
      },
      engineDeleteCols: (sheet: number, col: number, count: number) => {
        calls.deleteC.push({ sheet, col, count });
        return nativeOk;
      },
      setNumber: (a: { sheet: number; row: number; col: number }, value: number) => {
        calls.setNumberCalls.push({ sheet: a.sheet, row: a.row, col: a.col, value });
      },
      setText: () => {},
      setBool: () => {},
      setBlank: () => {},
      setFormula: () => {},
    } as unknown as WorkbookHandle;
    return { wb, calls };
  };

  it('insertRows calls engineInsertRows when capability is on', () => {
    const store = createSpreadsheetStore();
    const { wb, calls } = makeWb();
    insertRows(store, wb, null, 2, 3);
    expect(calls.insertR).toEqual([{ sheet: 0, row: 2, count: 3 }]);
    expect(calls.deleteR).toEqual([]);
  });

  it('leaves store and history untouched when the native op refuses', () => {
    const store = createSpreadsheetStore();
    const { wb, calls } = makeWb([], false);
    mutators.setCellFormat(store, { sheet: 0, row: 2, col: 0 }, { bold: true });
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, rowHeights: new Map([[2, 42]]) },
    }));
    const cutRange = { sheet: 0, r0: 2, c0: 0, r1: 2, c1: 0 };
    mutators.setCopyRange(store, cutRange, 'cut');
    const copyRevision = store.getState().ui.copyRevision;
    const before = store.getState();
    const h = new History();
    insertRows(store, wb, h, 2, 1);
    expect(calls.insertR).toEqual([{ sheet: 0, row: 2, count: 1 }]);
    expect(store.getState().format.formats).toBe(before.format.formats);
    expect(store.getState().layout.rowHeights).toBe(before.layout.rowHeights);
    expect(store.getState().ui.copyRange).toEqual(cutRange);
    expect(store.getState().ui.copyMode).toBe('cut');
    expect(store.getState().ui.copyRevision).toBe(copyRevision);
    expect(h.canUndo()).toBe(false);
  });

  it('rejects native inserts that would push populated tail cells off-grid', () => {
    const store = createSpreadsheetStore();
    const { wb, calls } = makeWb([
      {
        addr: { sheet: 0, row: 1_048_575, col: 0 },
        value: { kind: 'number', value: 7 },
        formula: null,
      },
      {
        addr: { sheet: 0, row: 0, col: 16_383 },
        value: { kind: 'number', value: 8 },
        formula: null,
      },
    ]);
    const h = new History();
    insertRows(store, wb, h, 1_048_574, 1);
    insertCols(store, wb, h, 16_382, 1);
    expect(calls.insertR).toEqual([]);
    expect(calls.insertC).toEqual([]);
    expect(h.canUndo()).toBe(false);
  });

  it('insertRows undo replays engineDeleteRows on the same band', () => {
    const store = createSpreadsheetStore();
    const { wb, calls } = makeWb();
    const h = new History();
    insertRows(store, wb, h, 1, 2);
    h.undo();
    expect(calls.deleteR).toEqual([{ sheet: 0, row: 1, count: 2 }]);
    h.redo();
    expect(calls.insertR).toEqual([
      { sheet: 0, row: 1, count: 2 },
      { sheet: 0, row: 1, count: 2 },
    ]);
  });

  it('deleteRows captures cells in the band and restores via setNumber on undo', () => {
    const store = createSpreadsheetStore();
    const { wb, calls } = makeWb([
      { addr: { sheet: 0, row: 1, col: 0 }, value: { kind: 'number', value: 42 }, formula: null },
      { addr: { sheet: 0, row: 5, col: 0 }, value: { kind: 'number', value: 99 }, formula: null },
    ]);
    const h = new History();
    deleteRows(store, wb, h, 1, 1);
    expect(calls.deleteR).toEqual([{ sheet: 0, row: 1, count: 1 }]);
    h.undo();
    expect(calls.insertR).toEqual([{ sheet: 0, row: 1, count: 1 }]);
    expect(calls.setNumberCalls).toEqual([{ sheet: 0, row: 1, col: 0, value: 42 }]);
  });

  it('insertCols / deleteCols route through their engine ops', () => {
    const store = createSpreadsheetStore();
    const { wb, calls } = makeWb();
    insertCols(store, wb, null, 4, 2);
    deleteCols(store, wb, null, 7, 1);
    expect(calls.insertC).toEqual([{ sheet: 0, col: 4, count: 2 }]);
    expect(calls.deleteC).toEqual([{ sheet: 0, col: 7, count: 1 }]);
  });
});
