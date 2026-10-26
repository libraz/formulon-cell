import { afterEach, describe, expect, it } from 'vitest';
import { commitCellInput } from '../../../src/commands/cell-input-commit.js';
import {
  formatAsTable,
  type TableOverlay,
  totalsCellAt,
  updateTableOverlay,
} from '../../../src/commands/format-as-table.js';
import { History } from '../../../src/commands/history.js';
import {
  renameActiveTable,
  setTableColumnFunction,
  toggleTableOption,
} from '../../../src/commands/table-design.js';
import { hydrateTableOverlaysFromEngine, totalsFromCell } from '../../../src/engine/table-sync.js';
import type { Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const GRID: string[][] = [
  ['Item', 'Region', 'Qty'],
  ['A', 'East', '1'],
  ['B', 'West', '2'],
  ['C', 'East', '3'],
];
const range: Range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 2 };

const open: WorkbookHandle[] = [];
afterEach(() => {
  for (const wb of open.splice(0)) wb.dispose();
});

async function setup(): Promise<{ wb: WorkbookHandle; store: SpreadsheetStore; history: History }> {
  const wb = await WorkbookHandle.createDefault();
  open.push(wb);
  GRID.forEach((cols, row) => {
    cols.forEach((text, col) => {
      if (col === 2 && row > 0) wb.setNumber({ sheet: 0, row, col }, Number(text));
      else wb.setText({ sheet: 0, row, col }, text);
    });
  });
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  hydrateTableOverlaysFromEngine(wb, store);
  wb.attachStore(store);
  const table = formatAsTable(store, range, { workbook: wb });
  expect(table?.name).toBe('Table1');
  mutators.setActive(store, { sheet: 0, row: 1, col: 1 });
  const history = new History();
  wb.attachHistory(history);
  return { wb, store, history };
}

const only = (store: SpreadsheetStore): TableOverlay => {
  const tables = store.getState().tables.tables;
  expect(tables).toHaveLength(1);
  return tables[0] as TableOverlay;
};

describe('table design commands', () => {
  it('renames through the name box command and reports refused names', async () => {
    const { wb, store, history } = await setup();
    expect(renameActiveTable(store, history, 'Sales')).toBe('renamed');
    expect(only(store).name).toBe('Sales');
    expect(wb.getTables().map((t) => t.name)).toEqual(['Sales']);
    expect(renameActiveTable(store, history, 'Sales')).toBe('unchanged');
    expect(renameActiveTable(store, history, 'bad name')).toBe('invalid');
    expect(renameActiveTable(store, history, '1abc')).toBe('invalid');
    expect(renameActiveTable(store, history, 'A1')).toBe('invalid');
    expect(renameActiveTable(store, history, '')).toBe('invalid');
    expect(only(store).name).toBe('Sales');
    history.undo();
    expect(only(store).name).toBe('Table1');
    expect(wb.getTables().map((t) => t.name)).toEqual(['Table1']);
  });

  it('refuses a name another table uses', async () => {
    const { wb, store, history } = await setup();
    formatAsTable(store, { sheet: 0, r0: 6, c0: 0, r1: 8, c1: 1 }, { workbook: wb });
    expect(store.getState().tables.tables).toHaveLength(2);
    expect(renameActiveTable(store, history, 'table2')).toBe('duplicate');
  });

  it('toggles the total row and fills the totals cell, with undo', async () => {
    const { wb, store, history } = await setup();
    expect(toggleTableOption(store, history, wb, 'showTotal')).toBe('toggled');
    const table = only(store);
    expect(table.showTotal).toBe(true);
    expect(wb.getTableStyleDetails(0)?.totalsRowCount).toBe(1);
    // The total row is added below the data; the default total sums the last numeric column.
    expect(table.totals?.[2]?.function).toBe('sum');
    expect(totalsCellAt(store.getState(), 0, 4, 2)?.column).toBe(2);
    expect(totalsCellAt(store.getState(), 0, 1, 2)).toBeNull();
    history.undo();
    expect(only(store).showTotal).toBe(false);
    expect(wb.getTableStyleDetails(0)?.totalsRowCount).toBe(0);
  });

  it('adds the total row below the data instead of overwriting the last row', async () => {
    const { wb, store, history } = await setup();
    toggleTableOption(store, history, wb, 'showTotal');
    expect(only(store).range).toMatchObject({ r0: 0, r1: 4 });
    expect(wb.getTables()[0]?.ref).toBe('A1:C5');
    expect(wb.getValue({ sheet: 0, row: 3, col: 0 })).toMatchObject({ kind: 'text', value: 'C' });
    expect(wb.getValue({ sheet: 0, row: 3, col: 2 })).toMatchObject({ kind: 'number', value: 3 });
    expect(wb.getValue({ sheet: 0, row: 4, col: 0 })).toMatchObject({ value: 'Total' });
  });

  it('labels a new total row in the given language', async () => {
    const { wb, store, history } = await setup();
    toggleTableOption(store, history, wb, 'showTotal', '集計');
    expect(wb.getValue({ sheet: 0, row: 4, col: 0 })).toMatchObject({ value: '集計' });
    expect(only(store).totals?.[0]).toMatchObject({ function: 'none', label: '集計' });
    const loaded = await WorkbookHandle.loadBytes(wb.save());
    open.push(loaded);
    expect(loaded.getTableStyleDetails(0)?.columns[0]?.totalsRowLabel).toBe('集計');
  });

  it('labels the total row of a table created with one in the given language', async () => {
    const { wb, store } = await setup();
    store.setState((s) => ({ ...s, tables: { ...s.tables, tables: [] } }));
    const created = formatAsTable(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 2 },
      { workbook: wb, showTotal: true, totalsLabel: '集計' },
    );
    expect(created?.totals?.[0]).toMatchObject({ function: 'none', label: '集計' });
    expect(wb.getValue({ sheet: 0, row: 4, col: 0 })).toMatchObject({ value: '集計' });
    expect(wb.getTableStyleDetails(0)?.columns[0]?.totalsRowLabel).toBe('集計');
  });

  it('keeps the existing total label when a column change re-creates the table', async () => {
    const { wb, store, history } = await setup();
    toggleTableOption(store, history, wb, 'showTotal', 'Summe');
    // Dropping the first column drops the column that carried the label.
    const id = only(store).id;
    expect(
      updateTableOverlay(store, id, { range: { sheet: 0, r0: 0, c0: 1, r1: 4, c1: 2 } }),
    ).not.toBeNull();
    expect(wb.getTables()[0]?.ref).toBe('B1:C5');
    expect(wb.getTableStyleDetails(0)?.columns[0]?.totalsRowLabel).toBe('Summe');
    expect(wb.getValue({ sheet: 0, row: 4, col: 1 })).toMatchObject({ value: 'Summe' });
  });

  it('shifts cells below the table down when the total row would land on them', async () => {
    const { wb, store, history } = await setup();
    wb.setText({ sheet: 0, row: 4, col: 0 }, 'below');
    wb.setText({ sheet: 0, row: 4, col: 5 }, 'outside');
    mutators.replaceCells(store, wb.cells(0));
    expect(toggleTableOption(store, history, wb, 'showTotal')).toBe('toggled');
    expect(wb.getValue({ sheet: 0, row: 5, col: 0 })).toMatchObject({ value: 'below' });
    expect(wb.getValue({ sheet: 0, row: 4, col: 5 })).toMatchObject({ value: 'outside' });
    expect(wb.getValue({ sheet: 0, row: 4, col: 0 })).toMatchObject({ value: 'Total' });
    // One undo step puts back both the cells and the table.
    history.undo();
    expect(only(store)).toMatchObject({ showTotal: false, range: { r1: 3 } });
    expect(wb.getTables()[0]?.ref).toBe('A1:C4');
    expect(wb.getValue({ sheet: 0, row: 4, col: 0 })).toMatchObject({ value: 'below' });
    expect(wb.getValue({ sheet: 0, row: 5, col: 0 }).kind).toBe('blank');
    history.redo();
    expect(wb.getValue({ sheet: 0, row: 5, col: 0 })).toMatchObject({ value: 'below' });
    expect(wb.getValue({ sheet: 0, row: 4, col: 0 })).toMatchObject({ value: 'Total' });
  });

  it('removes the total row and shrinks the table when it is turned off', async () => {
    const { wb, store, history } = await setup();
    toggleTableOption(store, history, wb, 'showTotal');
    expect(toggleTableOption(store, history, wb, 'showTotal')).toBe('toggled');
    expect(only(store).range).toMatchObject({ r1: 3 });
    expect(wb.getTables()[0]?.ref).toBe('A1:C4');
    expect(wb.getValue({ sheet: 0, row: 4, col: 0 }).kind).toBe('blank');
    expect(wb.getValue({ sheet: 0, row: 4, col: 2 }).kind).toBe('blank');
    history.undo();
    expect(wb.getTables()[0]?.ref).toBe('A1:C5');
    expect(wb.getValue({ sheet: 0, row: 4, col: 0 })).toMatchObject({ value: 'Total' });
    expect(wb.getValue({ sheet: 0, row: 4, col: 2 })).toMatchObject({ kind: 'number', value: 6 });
  });

  it('refuses a total row that would run into another table', async () => {
    const { wb, store, history } = await setup();
    formatAsTable(store, { sheet: 0, r0: 4, c0: 0, r1: 6, c1: 1 }, { workbook: wb });
    const first = store.getState().tables.tables[0];
    expect(toggleTableOption(store, history, wb, 'showTotal')).toBe('blocked');
    expect(store.getState().tables.tables[0]).toMatchObject({ id: first?.id, showTotal: false });
    expect(wb.getTables()[0]?.ref).toBe('A1:C4');
  });

  it('drops the header row and clears its cells', async () => {
    const { wb, store, history } = await setup();
    expect(toggleTableOption(store, history, wb, 'showHeader')).toBe('toggled');
    expect(only(store)).toMatchObject({ showHeader: false, range: { r0: 1, r1: 3 } });
    expect(wb.getTables()[0]?.ref).toBe('A2:C4');
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');
    history.undo();
    expect(wb.getTables()[0]?.ref).toBe('A1:C4');
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toMatchObject({ value: 'Region' });
  });

  describe('header row ON', () => {
    const headerless = async () => {
      const ctx = await setup();
      toggleTableOption(ctx.store, ctx.history, ctx.wb, 'showHeader');
      return ctx;
    };

    it('puts the header in an empty row above and extends the ref up, without shifting', async () => {
      const { wb, store, history } = await headerless();
      expect(toggleTableOption(store, history, wb, 'showHeader')).toBe('toggled');
      expect(only(store)).toMatchObject({ showHeader: true, range: { r0: 0, r1: 3 } });
      expect(wb.getTables()[0]?.ref).toBe('A1:C4');
      // The remembered column names come back, not generic ones.
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toMatchObject({ value: 'Item' });
      expect(wb.getValue({ sheet: 0, row: 0, col: 2 })).toMatchObject({ value: 'Qty' });
      expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toMatchObject({ value: 'A' });
      expect(wb.getValue({ sheet: 0, row: 3, col: 2 })).toMatchObject({ value: 3 });
      history.undo();
      expect(wb.getTables()[0]?.ref).toBe('A2:C4');
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');
      expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toMatchObject({ value: 'A' });
    });

    it('shifts the table down and uses its first row when the row above is occupied', async () => {
      const { wb, store, history } = await headerless();
      wb.setText({ sheet: 0, row: 0, col: 1 }, 'note');
      mutators.replaceCells(store, wb.cells(0));
      expect(toggleTableOption(store, history, wb, 'showHeader')).toBe('toggled');
      expect(only(store)).toMatchObject({ showHeader: true, range: { r0: 1, r1: 4 } });
      expect(wb.getTables()[0]?.ref).toBe('A2:C5');
      expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toMatchObject({ value: 'note' });
      expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toMatchObject({ value: 'Region' });
      expect(wb.getValue({ sheet: 0, row: 2, col: 1 })).toMatchObject({ value: 'East' });
      expect(wb.getValue({ sheet: 0, row: 4, col: 2 })).toMatchObject({ value: 3 });
      // One undo step restores the cells and the table.
      history.undo();
      expect(only(store)).toMatchObject({ showHeader: false, range: { r0: 1, r1: 3 } });
      expect(wb.getTables()[0]?.ref).toBe('A2:C4');
      expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toMatchObject({ value: 'East' });
      expect(wb.getValue({ sheet: 0, row: 4, col: 2 }).kind).toBe('blank');
      history.redo();
      expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toMatchObject({ value: 'Region' });
      expect(wb.getValue({ sheet: 0, row: 4, col: 2 })).toMatchObject({ value: 3 });
    });

    it('shifts the table down when it starts at the first sheet row', async () => {
      const ctx = await setup();
      const { wb, store, history } = ctx;
      // A table over the data rows only, starting at sheet row 1.
      store.setState((s) => ({ ...s, tables: { ...s.tables, tables: [] } }));
      expect(wb.getTables()).toEqual([]);
      formatAsTable(
        store,
        { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
        { workbook: wb, showHeader: false },
      );
      mutators.setActive(store, { sheet: 0, row: 1, col: 1 });
      expect(toggleTableOption(store, history, wb, 'showHeader')).toBe('toggled');
      expect(only(store)).toMatchObject({ showHeader: true, range: { r0: 0, r1: 3 } });
      expect(wb.getTables()[0]?.ref).toBe('A1:C4');
      expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toMatchObject({ value: 'Item' });
      expect(wb.getValue({ sheet: 0, row: 3, col: 0 })).toMatchObject({ value: 'B' });
    });
  });

  it('toggles banded columns and first/last column flags', async () => {
    const { wb, store, history } = await setup();
    expect(toggleTableOption(store, history, wb, 'colStripes')).toBe('toggled');
    expect(toggleTableOption(store, history, wb, 'firstCol')).toBe('toggled');
    expect(only(store)).toMatchObject({ colStripes: true, firstCol: true });
    expect(wb.getTableStyleDetails(0)).toMatchObject({
      showColumnStripes: true,
      showFirstColumn: true,
    });
  });

  it('keeps a picked totals function through save and load, with undo and redo', async () => {
    const { wb, store, history } = await setup();
    toggleTableOption(store, history, wb, 'showTotal');
    const id = only(store).id;
    expect(setTableColumnFunction(store, history, wb, id, 2, 'average')).toBe(true);
    expect(only(store).totals?.[2]?.function).toBe('average');
    expect(wb.getTableStyleDetails(0)?.columns[2]?.totalsRowFunction).toBe('average');
    expect(wb.getValue({ sheet: 0, row: 4, col: 2 })).toMatchObject({ kind: 'number', value: 2 });

    const loaded = await WorkbookHandle.loadBytes(wb.save());
    open.push(loaded);
    const store2 = createSpreadsheetStore();
    hydrateTableOverlaysFromEngine(loaded, store2);
    const reloaded = store2.getState().tables.tables[0];
    expect(reloaded?.showTotal).toBe(true);
    expect(reloaded?.totals?.map((t) => t.function)).toEqual(['none', 'none', 'average']);
    expect(reloaded?.totals?.[0]?.label).toBe('Total');

    history.undo();
    expect(only(store).totals?.[2]?.function).toBe('sum');
    expect(wb.getTableStyleDetails(0)?.columns[2]?.totalsRowFunction).toBe('sum');
    history.redo();
    expect(only(store).totals?.[2]?.function).toBe('average');
    expect(wb.getTableStyleDetails(0)?.columns[2]?.totalsRowFunction).toBe('average');
  });

  it('picks None for a column and clears its totals cell', async () => {
    const { wb, store, history } = await setup();
    toggleTableOption(store, history, wb, 'showTotal');
    expect(setTableColumnFunction(store, history, wb, only(store).id, 2, 'none')).toBe(true);
    expect(wb.getTableStyleDetails(0)?.columns[2]?.totalsRowFunction).toBe('none');
    expect(wb.getValue({ sheet: 0, row: 4, col: 2 }).kind).toBe('blank');
  });

  it('toggles the header row', async () => {
    const { wb, store, history } = await setup();
    expect(toggleTableOption(store, history, wb, 'showHeader')).toBe('toggled');
    expect(only(store).showHeader).toBe(false);
    expect(wb.getTableStyleDetails(0)?.headerRowCount).toBe(0);
    history.undo();
    expect(only(store).showHeader).toBe(true);
    expect(wb.getTableStyleDetails(0)?.headerRowCount).toBe(1);
  });

  describe('edits committed into a totals cell', () => {
    const commit = async (raw: string) => {
      const ctx = await setup();
      toggleTableOption(ctx.store, ctx.history, ctx.wb, 'showTotal');
      const addr = { sheet: 0, row: 4, col: 2 };
      commitCellInput({ store: ctx.store, wb: ctx.wb, addr, raw, origin: 'editor' });
      return ctx;
    };
    const reloaded = async (wb: WorkbookHandle) => {
      const loaded = await WorkbookHandle.loadBytes(wb.save());
      open.push(loaded);
      return loaded;
    };

    it('keeps a typed formula as a custom function through save and load, with undo', async () => {
      const { wb, store, history } = await commit('=SUM(C2:C4)*10');
      expect(only(store).totals?.[2]).toMatchObject({
        function: 'custom',
        formula: 'SUM(C2:C4)*10',
      });
      const loaded = await reloaded(wb);
      expect(loaded.getTableStyleDetails(0)?.columns[2]).toMatchObject({
        totalsRowFunction: 'custom',
        totalsRowFormula: 'SUM(C2:C4)*10',
      });
      expect(loaded.getValue({ sheet: 0, row: 4, col: 2 })).toMatchObject({ value: 60 });
      history.undo();
      expect(wb.getTableStyleDetails(0)?.columns[2]?.totalsRowFunction).toBe('sum');
    });

    it('keeps typed text as the column label', async () => {
      const { wb, store } = await commit('Subtotal');
      expect(only(store).totals?.[2]).toMatchObject({ function: 'none', label: 'Subtotal' });
      const loaded = await reloaded(wb);
      expect(loaded.getTableStyleDetails(0)?.columns[2]).toMatchObject({
        totalsRowFunction: 'none',
        totalsRowLabel: 'Subtotal',
      });
      expect(loaded.getValue({ sheet: 0, row: 4, col: 2 })).toMatchObject({ value: 'Subtotal' });
    });

    it('keeps a typed number as a constant custom formula', async () => {
      const { wb } = await commit('42');
      const loaded = await reloaded(wb);
      expect(loaded.getTableStyleDetails(0)?.columns[2]).toMatchObject({
        totalsRowFunction: 'custom',
        totalsRowFormula: '42',
      });
    });

    it('reads every built-in function back from the formula the engine writes', async () => {
      const { wb, store, history } = await setup();
      toggleTableOption(store, history, wb, 'showTotal');
      for (const fn of ['average', 'count', 'countNums', 'max', 'min', 'stdDev', 'sum', 'var']) {
        wb.setTableColumnTotals(0, 2, { function: fn as never });
        expect(totalsFromCell(wb, { sheet: 0, row: 4, col: 2 }).function).toBe(fn);
      }
    });

    it('follows an undone edit back to the built-in function', async () => {
      const { wb, store, history } = await commit('=SUM(C2:C4)*10');
      expect(only(store).totals?.[2]?.function).toBe('custom');
      history.undo();
      expect(only(store).totals?.[2]?.function).toBe('sum');
      expect(wb.getTableStyleDetails(0)?.columns[2]?.totalsRowFunction).toBe('sum');
      history.redo();
      expect(only(store).totals?.[2]?.function).toBe('custom');
      expect(wb.getTableStyleDetails(0)?.columns[2]?.totalsRowFormula).toBe('SUM(C2:C4)*10');
    });

    it('captures a cleared cell, a direct write and an atomic patch alike', async () => {
      const { wb, store, history } = await setup();
      toggleTableOption(store, history, wb, 'showTotal');
      wb.setBlank({ sheet: 0, row: 4, col: 2 });
      expect(only(store).totals?.[2]?.function).toBe('none');
      wb.applyCellPatchAtomic([
        { addr: { sheet: 0, row: 4, col: 2 }, value: { kind: 'blank' }, formula: '=MAX(C2:C4)' },
      ]);
      expect(only(store).totals?.[2]).toMatchObject({ function: 'custom', formula: 'MAX(C2:C4)' });
    });

    it('maps a cleared cell to none', async () => {
      const { wb, store } = await commit('');
      expect(only(store).totals?.[2]).toMatchObject({ function: 'none', label: '' });
      const loaded = await reloaded(wb);
      expect(loaded.getTableStyleDetails(0)?.columns[2]?.totalsRowFunction).toBe('none');
      expect(loaded.getValue({ sheet: 0, row: 4, col: 2 }).kind).toBe('blank');
    });
  });
});
