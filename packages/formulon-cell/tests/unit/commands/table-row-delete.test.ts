import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { deleteRows } from '../../../src/commands/structure.js';
import { hydrateTableOverlaysFromEngine } from '../../../src/engine/table-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const CALC = '=tblA[[#This Row],[Qty]]*10';
/** `CALC` as the engine reads it back inside the table (formula-bar form). */
const CALC_READ = '=[@Qty]*10';

/** tblA on A2:C7: header row 2, data rows 3-6 (Qty 1..4), totals row 7; B8 below the table. */
async function book() {
  const wb = await WorkbookHandle.createDefault();
  const heads = ['Item', 'Qty', 'Total'];
  for (const [col, h] of heads.entries()) wb.setText({ sheet: 0, row: 1, col }, h);
  for (let i = 0; i < 4; i += 1) {
    wb.setText({ sheet: 0, row: 2 + i, col: 0 }, 'abcd'[i] as string);
    wb.setNumber({ sheet: 0, row: 2 + i, col: 1 }, i + 1);
    wb.setFormula({ sheet: 0, row: 2 + i, col: 2 }, CALC);
  }
  wb.createTable(0, { range: 'A2:C7', name: 'tblA', columns: heads, totalsRow: true });
  wb.setFormula({ sheet: 0, row: 6, col: 1 }, '=SUBTOTAL(109,tblA[Qty])');
  wb.setNumber({ sheet: 0, row: 7, col: 1 }, 99);
  wb.setFormula({ sheet: 0, row: 3, col: 7 }, '=B8');
  wb.recalc();
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  hydrateTableOverlaysFromEngine(wb, store);
  wb.attachStore(store);
  return { wb, store, history: new History() };
}

const table = (store: ReturnType<typeof createSpreadsheetStore>) =>
  store.getState().tables.tables[0];
const num = (wb: WorkbookHandle, row: number, col: number) => {
  const v = wb.getValue({ sheet: 0, row, col });
  return v.kind === 'number' ? v.value : null;
};
const blank = (wb: WorkbookHandle, row: number, col: number) =>
  wb.getValue({ sheet: 0, row, col }).kind === 'blank';

describe('deleting sheet rows against a table', () => {
  it('refuses a delete that includes the header row', async () => {
    const { wb, store } = await book();
    expect(deleteRows(store, wb, null, 1, 1)).toBe(false);
    expect(deleteRows(store, wb, null, 0, 3)).toBe(false);
    expect(table(store)?.range).toEqual({ sheet: 0, r0: 1, c0: 0, r1: 6, c1: 2 });
    expect(wb.getTables()[0]?.ref).toBe('A2:C7');
    expect(num(wb, 7, 1)).toBe(99);
  });

  it('keeps one blank data row when every data row is deleted', async () => {
    const { wb, store, history } = await book();
    expect(deleteRows(store, wb, history, 2, 4)).toBe(true);
    expect(table(store)?.range).toEqual({ sheet: 0, r0: 1, c0: 0, r1: 3, c1: 2 });
    expect(table(store)?.showTotal).toBe(true);
    expect(wb.getTables()[0]?.ref).toBe('A2:C4');
    for (const col of [0, 1, 2]) expect(blank(wb, 2, col)).toBe(true);
    expect(wb.cellFormula({ sheet: 0, row: 3, col: 1 })).toMatch(/SUBTOTAL\(109/);
    expect(num(wb, 3, 1)).toBe(0);
    expect(num(wb, 4, 1)).toBe(99);
    history.undo();
    expect(table(store)?.range).toEqual({ sheet: 0, r0: 1, c0: 0, r1: 6, c1: 2 });
    expect(num(wb, 2, 1)).toBe(1);
    expect(wb.cellFormula({ sheet: 0, row: 2, col: 2 })).toBe(CALC_READ);
  });

  it('drops the totals row when the last data row and the totals row are deleted', async () => {
    const { wb, store } = await book();
    expect(deleteRows(store, wb, null, 5, 2)).toBe(true);
    expect(table(store)?.range).toEqual({ sheet: 0, r0: 1, c0: 0, r1: 4, c1: 2 });
    expect(table(store)?.showTotal).toBe(false);
    expect(wb.getTables()[0]?.ref).toBe('A2:C5');
    expect(num(wb, 5, 1)).toBe(99);
  });

  it('drops the totals row when the totals row and the row below are deleted', async () => {
    const { wb, store } = await book();
    expect(deleteRows(store, wb, null, 6, 2)).toBe(true);
    expect(table(store)?.range).toEqual({ sheet: 0, r0: 1, c0: 0, r1: 5, c1: 2 });
    expect(table(store)?.showTotal).toBe(false);
    expect(wb.getTables()[0]?.ref).toBe('A2:C6');
    expect(wb.cellFormula({ sheet: 0, row: 3, col: 7 })).toMatch(/#REF!/);
  });
});
