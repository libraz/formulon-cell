import { FormulonError, StatusCode } from '@libraz/formulon';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../../src/commands/history.js';
import { deleteRows, insertRows } from '../../../../src/commands/structure.js';
import type { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';
import { newWb } from './fixtures.js';

// A cell write refused partway through the in-memory row shift must leave the
// sheet and the history as they were.

afterEach(() => {
  vi.restoreAllMocks();
});

async function fixture() {
  const store = createSpreadsheetStore();
  const wb = await newWb();
  const history = new History();
  wb.attachHistory(history);
  wb.setFormula({ sheet: 0, row: 0, col: 0 }, '=A3+1');
  wb.setNumber({ sheet: 0, row: 2, col: 0 }, 5);
  wb.setFormula({ sheet: 0, row: 3, col: 0 }, '=A3*2');
  wb.setFormula({ sheet: 0, row: 4, col: 0 }, '=A4+A3');
  wb.setNumber({ sheet: 0, row: 5, col: 0 }, 7);
  wb.recalc();
  history.clear();
  return { store, wb, history };
}

const cells = (wb: WorkbookHandle) =>
  Array.from(wb.cells(0))
    .map((c) => ({ r: c.addr.row, c: c.addr.col, v: c.value, f: c.formula }))
    .sort((x, y) => x.r - y.r || x.c - y.c);

/** Make the `nth` formula write on sheet 0 refuse. */
function refuseNthFormula(wb: WorkbookHandle, nth: number): () => number {
  const sheet = wb.engineWorkbook().sheet(0) as unknown as {
    setFormula: (...args: unknown[]) => void;
  };
  const setFormula = sheet.setFormula.bind(sheet);
  let calls = 0;
  vi.spyOn(sheet, 'setFormula').mockImplementation((...args: unknown[]) => {
    calls += 1;
    if (calls === nth) throw new FormulonError(StatusCode.InvalidArgument, 'refused', '');
    setFormula(...args);
  });
  return () => calls;
}

describe('row shift under a refused cell write', () => {
  it.each([
    { what: 'insert', edit: insertRows },
    { what: 'delete', edit: deleteRows },
  ])('$what leaves the cells and the history unchanged', async ({ edit }) => {
    const { store, wb, history } = await fixture();
    expect(wb.isStub).toBe(true);
    const before = cells(wb);

    const calls = refuseNthFormula(wb, 2);
    expect(() => edit(store, wb, history, 1, 1)).toThrow();
    expect(calls()).toBeGreaterThanOrEqual(2);
    expect(cells(wb)).toEqual(before);
    expect(history.canUndo()).toBe(false);
  });
});
