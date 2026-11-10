import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { applyMerge } from '../../../src/commands/merge.js';
import { setFreezePanes } from '../../../src/commands/row-col-layout.js';
import { deleteRows } from '../../../src/commands/structure.js';
import { setTableColumnFunction, toggleTableOption } from '../../../src/commands/table-design.js';
import {
  hydrateTableOverlaysFromEngine,
  totalsFunctionFromEngine,
} from '../../../src/engine/table-sync.js';
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

  it('forgets the column totals functions when the totals row is deleted', async () => {
    const { wb, store } = await book();
    const id = table(store)?.id ?? '';
    expect(setTableColumnFunction(store, null, wb, id, 1, 'max')).toBe(true);
    expect(table(store)?.totals?.[1]?.function).toBe('max');
    expect(deleteRows(store, wb, null, 6, 1)).toBe(true);
    expect(table(store)?.showTotal).toBe(false);
    mutators.setActive(store, { sheet: 0, row: 2, col: 0 });
    expect(toggleTableOption(store, null, wb, 'showTotal')).toBe('toggled');
    expect(table(store)?.showTotal).toBe(true);
    expect(table(store)?.totals?.[1]?.function).not.toBe('max');
    const engineColumn = wb.getTableStyleDetails(0)?.columns[1];
    expect(engineColumn && totalsFunctionFromEngine(engineColumn.totalsRowFunction)).not.toBe(
      'max',
    );
  });
  it('keeps the first data row blank, shifts what lies below by one row fewer, and undoes exactly', async () => {
    const { wb, store, history } = await book();
    // Below the table: B8 holds 99; give it a format, merge E9:F9 and freeze the top 8 rows.
    mutators.setCellFormat(store, { sheet: 0, row: 7, col: 1 }, { bold: true });
    expect(applyMerge(store, wb, null, { sheet: 0, r0: 8, c0: 4, r1: 8, c1: 5 })).toBe(true);
    setFreezePanes(store, 8, 0, wb);
    const snapshot = () => ({
      cells: Array.from({ length: 10 }, (_, row) =>
        [0, 1, 2].map((col) => ({
          value: wb.getValue({ sheet: 0, row, col }),
          formula: wb.cellFormula({ sheet: 0, row, col }),
        })),
      ),
      table: table(store)?.range,
      ref: wb.getTables()[0]?.ref,
      bold: [...store.getState().format.formats]
        .filter(([, f]) => f.bold)
        .map(([key]) => key)
        .sort(),
      merges: [...store.getState().merges.byAnchor.values()].map((m) => ({ ...m })),
      engineMerges: wb.getMerges(0),
      freezeRows: store.getState().layout.freezeRows,
    });
    const before = snapshot();

    // Rows 3-6 (every data row): the engine keeps row 3 blank and removes the three below it.
    expect(deleteRows(store, wb, history, 2, 4)).toBe(true);
    const after = snapshot();
    expect(after.table).toEqual({ sheet: 0, r0: 1, c0: 0, r1: 3, c1: 2 });
    expect(after.ref).toBe('A2:C4');
    for (const col of [0, 1, 2]) expect(blank(wb, 2, col)).toBe(true);
    expect(num(wb, 4, 1)).toBe(99);
    expect(after.bold).toEqual(['0:4:1']);
    expect(after.merges).toEqual([{ sheet: 0, r0: 5, c0: 4, r1: 5, c1: 5 }]);
    expect(after.engineMerges).toEqual([{ sheet: 0, r0: 5, c0: 4, r1: 5, c1: 5 }]);
    expect(after.freezeRows).toBe(5);

    expect(history.undo()).toBe(true);
    expect(snapshot()).toEqual(before);
    expect(history.canUndo()).toBe(false);
    expect(history.redo()).toBe(true);
    expect(snapshot()).toEqual(after);
  });

  it('blanks the only data row without removing a row', async () => {
    const wb = await WorkbookHandle.createDefault();
    wb.setText({ sheet: 0, row: 0, col: 0 }, 'Item');
    wb.setText({ sheet: 0, row: 1, col: 0 }, 'a');
    wb.setText({ sheet: 0, row: 2, col: 0 }, 'below');
    wb.createTable(0, { range: 'A1:A2', name: 'tblB', columns: ['Item'] });
    const store = createSpreadsheetStore();
    mutators.replaceCells(store, wb.cells(0));
    hydrateTableOverlaysFromEngine(wb, store);
    wb.attachStore(store);
    const history = new History();
    const text = (row: number) => wb.getValue({ sheet: 0, row, col: 0 });
    expect(deleteRows(store, wb, history, 1, 1)).toBe(true);
    expect(text(1)).toEqual({ kind: 'blank' });
    expect(text(2)).toEqual({ kind: 'text', value: 'below' });
    expect(table(store)?.range).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 });
    expect(history.undo()).toBe(true);
    expect(text(1)).toEqual({ kind: 'text', value: 'a' });
    expect(text(2)).toEqual({ kind: 'text', value: 'below' });
    expect(history.redo()).toBe(true);
    expect(text(1)).toEqual({ kind: 'blank' });
    expect(text(2)).toEqual({ kind: 'text', value: 'below' });
  });
  it('grows the engine table back on undo when it has no totals row', async () => {
    const wb = await WorkbookHandle.createDefault();
    [
      ['Item', 'Qty'],
      ['Pen', 3],
      ['Ink', 5],
      ['Pad', 2],
    ].forEach(([item, qty], row) => {
      wb.setText({ sheet: 0, row, col: 0 }, item as string);
      if (typeof qty === 'number') wb.setNumber({ sheet: 0, row, col: 1 }, qty);
      else wb.setText({ sheet: 0, row, col: 1 }, qty as string);
    });
    wb.createTable(0, { range: 'A1:B4', name: 'Sales', columns: ['Item', 'Qty'] });
    const store = createSpreadsheetStore();
    mutators.replaceCells(store, wb.cells(0));
    hydrateTableOverlaysFromEngine(wb, store);
    wb.attachStore(store);
    const history = new History();
    // The engine's own record of the table, not the overlay.
    const engineRef = () => wb.getTables().find((t) => t.name === 'Sales')?.ref;
    const text = (row: number) => wb.getValue({ sheet: 0, row, col: 0 });

    expect(deleteRows(store, wb, history, 1, 3)).toBe(true);
    expect(engineRef()).toBe('A1:B2');
    expect(text(1)).toEqual({ kind: 'blank' });

    expect(history.undo()).toBe(true);
    expect(engineRef()).toBe('A1:B4');
    expect(table(store)?.range).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 });
    expect([1, 2, 3].map((row) => text(row))).toEqual([
      { kind: 'text', value: 'Pen' },
      { kind: 'text', value: 'Ink' },
      { kind: 'text', value: 'Pad' },
    ]);
    expect(num(wb, 3, 1)).toBe(2);

    expect(history.redo()).toBe(true);
    expect(engineRef()).toBe('A1:B2');
    expect(text(1)).toEqual({ kind: 'blank' });
    expect(table(store)?.range).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    wb.detachStore(store);
    wb.dispose();
  });
});
