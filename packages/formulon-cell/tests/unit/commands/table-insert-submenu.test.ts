import { describe, expect, it } from 'vitest';
import { insertTableColumns, insertTableRows } from '../../../src/commands/cell-shift.js';
import { History } from '../../../src/commands/history.js';
import { hydrateTableOverlaysFromEngine } from '../../../src/engine/table-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings/en.js';
import { ja } from '../../../src/i18n/strings/ja.js';
import {
  buildMacCellEntries,
  withTableInsertMenu,
} from '../../../src/interact/context-menu-spec.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const CALC = '=tblA[[#This Row],[Qty]]*10';

/** tblA on A2:C7: header row 2, data rows 3-6, totals row 7; E:F beside it; B8 below; H3/H4 outside. */
async function book() {
  const wb = await WorkbookHandle.createDefault();
  const heads = ['Item', 'Qty', 'Total'];
  for (const [col, h] of heads.entries()) wb.setText({ sheet: 0, row: 1, col }, h);
  for (let i = 0; i < 4; i += 1) {
    wb.setText({ sheet: 0, row: 2 + i, col: 0 }, 'abcd'[i] as string);
    wb.setNumber({ sheet: 0, row: 2 + i, col: 1 }, i + 1);
    wb.setFormula({ sheet: 0, row: 2 + i, col: 2 }, CALC);
  }
  wb.createTable({ sheetIndex: 0, ref: 'A2:C7', name: 'tblA', columns: heads, totalsRow: true });
  wb.setText({ sheet: 0, row: 6, col: 0 }, 'Sum');
  wb.setFormula({ sheet: 0, row: 6, col: 1 }, '=SUBTOTAL(109,tblA[Qty])');
  for (let r = 1; r <= 6; r += 1) {
    wb.setText({ sheet: 0, row: r, col: 4 }, `E${r + 1}`);
    wb.setText({ sheet: 0, row: r, col: 5 }, `F${r + 1}`);
  }
  wb.setNumber({ sheet: 0, row: 7, col: 1 }, 99);
  wb.setFormula({ sheet: 0, row: 2, col: 7 }, '=SUM(tblA[Qty])');
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
const text = (wb: WorkbookHandle, row: number, col: number) => {
  const v = wb.getValue({ sheet: 0, row, col });
  return v.kind === 'text' ? v.value : null;
};
const num = (wb: WorkbookHandle, row: number, col: number) => {
  const v = wb.getValue({ sheet: 0, row, col });
  return v.kind === 'number' ? v.value : null;
};

describe('Insert > Table Rows Above', () => {
  it('grows the table by one row, fills the calculated column and leaves the columns beside it', async () => {
    const { wb, store } = await book();
    const sel = { sheet: 0, r0: 3, c0: 1, r1: 3, c1: 1 };
    expect(insertTableRows(store, wb, null, sel)).toBe(true);
    expect(table(store)?.range).toEqual({ sheet: 0, r0: 1, c0: 0, r1: 7, c1: 2 });
    expect(wb.getTables()[0]?.ref).toBe('A2:C8');
    expect(wb.cellFormula({ sheet: 0, row: 3, col: 2 })).toBe(CALC);
    expect(text(wb, 4, 0)).toBe('b');
    expect(num(wb, 8, 1)).toBe(99);
    expect(wb.cellFormula({ sheet: 0, row: 3, col: 7 })).toBe('=B9');
    expect(num(wb, 2, 7)).toBe(10);
    expect(text(wb, 3, 4)).toBe('E4');
    expect(text(wb, 6, 5)).toBe('F7');
    expect(wb.getValue({ sheet: 0, row: 7, col: 4 }).kind).toBe('blank');
  });

  it('inserts as many rows as the selection spans', async () => {
    const { wb, store } = await book();
    expect(insertTableRows(store, wb, null, { sheet: 0, r0: 3, c0: 1, r1: 4, c1: 1 })).toBe(true);
    expect(table(store)?.range).toEqual({ sheet: 0, r0: 1, c0: 0, r1: 8, c1: 2 });
    expect(wb.cellFormula({ sheet: 0, row: 4, col: 2 })).toBe(CALC);
    expect(num(wb, 9, 1)).toBe(99);
  });

  it('does nothing outside a table or on the header row', async () => {
    const { wb, store } = await book();
    expect(insertTableRows(store, wb, null, { sheet: 0, r0: 12, c0: 1, r1: 12, c1: 1 })).toBe(
      false,
    );
    expect(insertTableRows(store, wb, null, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 })).toBe(false);
    expect(table(store)?.range).toEqual({ sheet: 0, r0: 1, c0: 0, r1: 6, c1: 2 });
  });
});

describe('Insert > Table Columns to the Left', () => {
  const HEADER = /^(Column|列)1$/;

  it('shifts the table rows right, widens the table and names the new header', async () => {
    const { wb, store, history } = await book();
    wb.setText({ sheet: 0, row: 0, col: 1 }, 'aboveB');
    wb.setText({ sheet: 0, row: 8, col: 0 }, 'below');
    wb.setFormula({ sheet: 0, row: 4, col: 7 }, '=A9');
    const sel = { sheet: 0, r0: 3, c0: 1, r1: 3, c1: 1 };
    expect(insertTableColumns(store, wb, history, sel)).toBe(true);
    expect(table(store)?.range).toEqual({ sheet: 0, r0: 1, c0: 0, r1: 6, c1: 3 });
    expect(wb.getTables()[0]?.ref).toBe('A2:D7');
    expect(text(wb, 1, 1)).toMatch(HEADER);
    expect(text(wb, 1, 2)).toBe('Qty');
    expect(text(wb, 1, 3)).toBe('Total');
    expect(wb.getValue({ sheet: 0, row: 2, col: 1 }).kind).toBe('blank');
    expect(num(wb, 3, 2)).toBe(2);
    expect(wb.cellFormula({ sheet: 0, row: 3, col: 3 })).toMatch(/\[Qty\]/);
    expect(wb.cellFormula({ sheet: 0, row: 6, col: 2 })).toMatch(/SUBTOTAL\(109/);
    // Rows above and below the table stay put.
    expect(text(wb, 0, 1)).toBe('aboveB');
    expect(num(wb, 7, 1)).toBe(99);
    expect(text(wb, 8, 0)).toBe('below');
    // Cells to the right on the table's rows shift right.
    expect(text(wb, 1, 5)).toBe('E2');
    expect(text(wb, 6, 6)).toBe('F7');
    expect(wb.getValue({ sheet: 0, row: 3, col: 4 }).kind).toBe('blank');
    expect(wb.cellFormula({ sheet: 0, row: 2, col: 8 })).toBe('=SUM(tblA[Qty])');
    expect(wb.cellFormula({ sheet: 0, row: 3, col: 8 })).toBe('=B8');
    expect(wb.cellFormula({ sheet: 0, row: 4, col: 8 })).toBe('=A9');
    history.undo();
    expect(table(store)?.range).toEqual({ sheet: 0, r0: 1, c0: 0, r1: 6, c1: 2 });
  });

  it('grows the table leftwards from the first column', async () => {
    const { wb, store } = await book();
    expect(insertTableColumns(store, wb, null, { sheet: 0, r0: 3, c0: 0, r1: 3, c1: 0 })).toBe(
      true,
    );
    expect(table(store)?.range).toEqual({ sheet: 0, r0: 1, c0: 0, r1: 6, c1: 3 });
    expect(wb.getTables()[0]?.ref).toBe('A2:D7');
    expect(text(wb, 1, 0)).toMatch(HEADER);
    expect(text(wb, 1, 1)).toBe('Item');
    expect(text(wb, 6, 1)).toBe('Sum');
    expect(wb.cellFormula({ sheet: 0, row: 6, col: 2 })).toMatch(/SUBTOTAL\(109/);
    expect(num(wb, 7, 1)).toBe(99);
  });
});

describe('table insert menu entries', () => {
  it('replaces the flat Insert entry with a submenu holding the two table entries', () => {
    const entries = withTableInsertMenu(buildMacCellEntries(ja, false), ja, { rows: true });
    const insert = entries.find((e) => e.kind === 'submenu' && e.id === 'insertTableMenu');
    expect(insert?.kind === 'submenu' && insert.label).toBe(ja.macContextMenu.insert);
    expect(entries.some((e) => e.kind === 'item' && e.id === 'insertCells')).toBe(false);
    const kids = insert?.kind === 'submenu' ? insert.children : [];
    expect(kids.map((k) => (k.kind === 'item' ? [k.id, k.label] : null))).toEqual([
      ['insertTableColumns', 'テーブルの列 (左)'],
      ['insertTableRows', 'テーブルの行 (上)'],
    ]);
  });

  it('draws the row entry disabled on the header row and has English wording', () => {
    const entries = withTableInsertMenu(buildMacCellEntries(en, false), en, { rows: false });
    const insert = entries.find((e) => e.kind === 'submenu' && e.id === 'insertTableMenu');
    const kids = insert?.kind === 'submenu' ? insert.children : [];
    expect(kids).toMatchObject([
      { id: 'insertTableColumns', label: 'Table Columns to the Left' },
      { id: 'insertTableRows', label: 'Table Rows Above', disabled: true },
    ]);
  });
});
