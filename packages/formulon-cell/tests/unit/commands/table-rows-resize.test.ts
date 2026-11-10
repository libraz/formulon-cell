import { describe, expect, it } from 'vitest';
import { deleteCells, insertCells } from '../../../src/commands/cell-shift.js';
import { History } from '../../../src/commands/history.js';
import { deleteRows, insertRows } from '../../../src/commands/structure.js';
import { hydrateTableOverlaysFromEngine } from '../../../src/engine/table-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const CALC = '=tblSales[[#This Row],[Qty]]*tblSales[[#This Row],[Price]]';
/** `CALC` as the engine reads it back inside the table (formula-bar form). */
const CALC_READ = '=[@Qty]*[@Price]';

/** tblSales on A1:D6: header row, data rows 2-5 (Qty 1..4, Price 10), totals row 6. */
async function book() {
  const wb = await WorkbookHandle.createDefault();
  const heads = ['Name', 'Qty', 'Price', 'Total'];
  for (const [col, h] of heads.entries()) wb.setText({ sheet: 0, row: 0, col }, h);
  for (let i = 1; i <= 4; i += 1) {
    wb.setText({ sheet: 0, row: i, col: 0 }, `n${i}`);
    wb.setNumber({ sheet: 0, row: i, col: 1 }, i);
    wb.setNumber({ sheet: 0, row: i, col: 2 }, 10);
    wb.setFormula({ sheet: 0, row: i, col: 3 }, CALC);
  }
  expect(
    wb.createTable(0, {
      range: 'A1:D6',
      name: 'tblSales',
      columns: heads,
      totalsRow: true,
    }),
  ).toBe(0);
  wb.setFormula({ sheet: 0, row: 5, col: 1 }, '=SUBTOTAL(109,tblSales[Qty])');
  wb.recalc();
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  hydrateTableOverlaysFromEngine(wb, store);
  wb.attachStore(store);
  mutators.setActive(store, { sheet: 0, row: 1, col: 0 });
  return { wb, store, history: new History() };
}

const range = (store: ReturnType<typeof createSpreadsheetStore>) =>
  store.getState().tables.tables[0]?.range;
const num = (wb: WorkbookHandle, row: number, col: number) => {
  const v = wb.getValue({ sheet: 0, row, col });
  return v.kind === 'number' ? v.value : null;
};

describe('table rows follow row edits', () => {
  it('Delete > Table Rows shrinks the table and moves the totals row', async () => {
    const { wb, store, history } = await book();
    expect(deleteCells(store, wb, history, { sheet: 0, r0: 3, c0: 0, r1: 3, c1: 3 }, 'up')).toBe(
      true,
    );
    expect(range(store)).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 4, c1: 3 });
    expect(wb.getTables()[0]?.ref).toBe('A1:D5');
    expect(wb.cellFormula({ sheet: 0, row: 4, col: 1 })).toMatch(/SUBTOTAL\(109/);
    expect(num(wb, 4, 1)).toBe(7);
    history.undo();
    expect(range(store)).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 5, c1: 3 });
  });

  it('inserting a sheet row at the totals row adds a data row with the calculated column', async () => {
    const { wb, store, history } = await book();
    expect(insertRows(store, wb, history, 5, 1)).toBe(true);
    expect(range(store)).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 6, c1: 3 });
    expect(wb.getTables()[0]?.ref).toBe('A1:D7');
    expect(wb.cellFormula({ sheet: 0, row: 5, col: 3 })).toBe(CALC_READ);
    expect(num(wb, 5, 3)).toBe(0);
    expect(wb.cellFormula({ sheet: 0, row: 6, col: 1 })).toMatch(/SUBTOTAL\(109/);
    history.undo();
    expect(range(store)).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 5, c1: 3 });
  });

  it('inserting and deleting sheet rows inside the body grow and shrink the table', async () => {
    const { wb, store } = await book();
    insertRows(store, wb, null, 2, 2);
    expect(range(store)).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 7, c1: 3 });
    expect(wb.getTables()[0]?.ref).toBe('A1:D8');
    deleteRows(store, wb, null, 2, 2);
    expect(range(store)).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 5, c1: 3 });
    expect(wb.getTables()[0]?.ref).toBe('A1:D6');
  });

  it('Insert > Table Rows Above grows the table', async () => {
    const { wb, store } = await book();
    insertCells(store, wb, null, { sheet: 0, r0: 2, c0: 0, r1: 2, c1: 3 }, 'down');
    expect(range(store)).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 6, c1: 3 });
    expect(wb.cellFormula({ sheet: 0, row: 2, col: 3 })).toBe(CALC_READ);
  });
});
