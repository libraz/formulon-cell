import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { insertRows } from '../../../src/commands/structure.js';
import { hydrateTableOverlaysFromEngine } from '../../../src/engine/table-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const STRUCT = '=tblSales[[#This Row],[Qty]]*tblSales[[#This Row],[Price]]';

type Variant = 'structured' | 'a1' | 'exception';

/** tblSales on A1:D6: header, data rows 2-5 (Qty 1..4, Price 10), totals row 6. */
async function book(variant: Variant) {
  const wb = await WorkbookHandle.createDefault();
  const heads = ['Name', 'Qty', 'Price', 'Total'];
  for (const [col, h] of heads.entries()) wb.setText({ sheet: 0, row: 0, col }, h);
  for (let i = 1; i <= 4; i += 1) {
    wb.setText({ sheet: 0, row: i, col: 0 }, `n${i}`);
    wb.setNumber({ sheet: 0, row: i, col: 1 }, i);
    wb.setNumber({ sheet: 0, row: i, col: 2 }, 10);
    wb.setFormula({ sheet: 0, row: i, col: 3 }, variant === 'a1' ? `=B${i + 1}*C${i + 1}` : STRUCT);
  }
  wb.createTable(0, { range: 'A1:D6', name: 'tblSales', columns: heads, totalsRow: true });
  wb.setFormula({ sheet: 0, row: 5, col: 1 }, '=SUBTOTAL(109,tblSales[Qty])');
  if (variant === 'exception') wb.setNumber({ sheet: 0, row: 3, col: 3 }, 99);
  wb.recalc();
  return wb;
}

/** The sheet-row insert command (native insert plus overlay follow-up). */
async function viaCommand(variant: Variant, at: number, count: number) {
  const wb = await book(variant);
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  hydrateTableOverlaysFromEngine(wb, store);
  wb.attachStore(store);
  mutators.setActive(store, { sheet: 0, row: 1, col: 0 });
  const history = new History();
  expect(insertRows(store, wb, history, at, count)).toBe(true);
  wb.recalc();
  return { wb, store, history };
}

/** Native sheet insert only. */
async function viaSheet(variant: Variant, at: number, count: number) {
  const wb = await book(variant);
  expect(wb.engineInsertRows(0, at, count)).toBe(true);
  wb.recalc();
  return wb;
}

/** Native Table.insertRows; `at` is the sheet row, the data row position is `at - 1`. */
async function viaTable(variant: Variant, at: number, count: number) {
  const wb = await book(variant);
  expect(wb.insertTableRows(0, at - 1, at - 1 + count - 1)).toBe(true);
  wb.recalc();
  return wb;
}

const num = (wb: WorkbookHandle, row: number, col: number) => {
  const v = wb.getValue({ sheet: 0, row, col });
  return v.kind === 'number' ? v.value : null;
};

/** Total-column formula a new data row `row` holds; a column with an exception is still filled. */
function newRowFormula(variant: Variant, row: number): string {
  return variant === 'a1' ? `=B${row + 1}*C${row + 1}` : '=[@Qty]*[@Price]';
}

/** Assert the table after `count` rows went in at sheet row `at`. */
function expectInserted(wb: WorkbookHandle, variant: Variant, at: number, count: number): void {
  expect(wb.getTables().map((t) => t.ref)).toEqual([`A1:D${6 + count}`]);
  for (let row = at; row < at + count; row += 1) {
    expect(wb.cellFormula({ sheet: 0, row, col: 3 })).toBe(newRowFormula(variant, row));
  }
  // Data rows below the insert moved down with their calculated cells.
  for (let orig = 1; orig <= 4; orig += 1) {
    if (orig < at) continue;
    const row = orig + count;
    expect(num(wb, row, 1)).toBe(orig);
    const overridden = variant === 'exception' && orig === 3;
    expect(wb.cellFormula({ sheet: 0, row, col: 3 })).toBe(
      overridden ? null : newRowFormula(variant, row),
    );
    expect(num(wb, row, 3)).toBe(overridden ? 99 : orig * 10);
  }
  if (variant === 'structured') {
    for (let row = at; row < at + count; row += 1) expect(num(wb, row, 3)).toBe(0);
  } else if (variant === 'exception') {
    expect(num(wb, 3 >= at ? 3 + count : 3, 3)).toBe(99);
  }
  expect(wb.cellFormula({ sheet: 0, row: 5 + count, col: 1 })).toMatch(/SUBTOTAL\(109/);
}

const CASES: { name: string; at: number; count: number }[] = [
  { name: 'top of the table', at: 1, count: 1 },
  { name: 'middle of the table', at: 3, count: 1 },
  { name: 'at the totals row (bottom)', at: 5, count: 1 },
  { name: 'several rows in the middle', at: 3, count: 3 },
  { name: 'several rows at the bottom', at: 5, count: 2 },
];

describe('the engine fills calculated columns of inserted table rows', () => {
  for (const variant of ['structured', 'a1', 'exception'] as Variant[]) {
    for (const c of CASES) {
      it(`${variant}: sheet insert, ${c.name}`, async () => {
        expectInserted(await viaSheet(variant, c.at, c.count), variant, c.at, c.count);
      });

      it(`${variant}: Table.insertRows, ${c.name}`, async () => {
        expectInserted(await viaTable(variant, c.at, c.count), variant, c.at, c.count);
      });

      it(`${variant}: insertRows command, ${c.name}`, async () => {
        const { wb } = await viaCommand(variant, c.at, c.count);
        expectInserted(wb, variant, c.at, c.count);
      });
    }
  }

  for (const variant of ['structured', 'a1'] as Variant[]) {
    it(`${variant}: undo removes the new rows and redo restores them`, async () => {
      const { wb, store, history } = await viaCommand(variant, 3, 2);
      history.undo();
      wb.recalc();
      expect(wb.getTables().map((t) => t.ref)).toEqual(['A1:D6']);
      expect(store.getState().tables.tables[0]?.range.r1).toBe(5);
      const fresh = await book(variant);
      for (let row = 0; row < 8; row += 1) {
        for (let col = 0; col < 4; col += 1) {
          const addr = { sheet: 0, row, col };
          expect(wb.cellFormula(addr)).toBe(fresh.cellFormula(addr));
          expect(wb.getValue(addr)).toEqual(fresh.getValue(addr));
        }
      }
      history.redo();
      wb.recalc();
      expectInserted(wb, variant, 3, 2);
      expect(store.getState().tables.tables[0]?.range.r1).toBe(7);
    });
  }
});
