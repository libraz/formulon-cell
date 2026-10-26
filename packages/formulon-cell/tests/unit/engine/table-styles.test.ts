import { afterEach, describe, expect, it } from 'vitest';
import {
  createPivotTableStyleFromActivePivot,
  createTableStyleFromActiveTable,
  defaultTableOverlay,
  formatAsTable,
  overlayStyleName,
  parseTableStyleName,
  type TableOverlay,
  tableStyleName,
  tableStyleOptionsById,
  updateTableOverlay,
} from '../../../src/commands/format-as-table.js';
import { History } from '../../../src/commands/history.js';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { recordTablesChange } from '../../../src/commands/slice-history.js';
import { hydrateTableOverlaysFromEngine } from '../../../src/engine/table-sync.js';
import { PivotAggregation, type Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const range: Range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 2 };

const open: WorkbookHandle[] = [];

afterEach(() => {
  for (const wb of open.splice(0)) wb.dispose();
});

async function seeded(): Promise<WorkbookHandle> {
  const wb = await WorkbookHandle.createDefault();
  open.push(wb);
  const rows: (string | number)[][] = [
    ['Item', 'Region', 'Qty'],
    ['A', 'East', 1],
    ['B', 'West', 2],
    ['C', 'East', 3],
  ];
  rows.forEach((cols, row) => {
    cols.forEach((v, col) => {
      const addr = { sheet: 0, row, col };
      if (typeof v === 'number') wb.setNumber(addr, v);
      else wb.setText(addr, v);
    });
  });
  return wb;
}

function bound(wb: WorkbookHandle): SpreadsheetStore {
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  hydrateTableOverlaysFromEngine(wb, store);
  wb.attachStore(store);
  return store;
}

async function reloaded(
  wb: WorkbookHandle,
): Promise<{ wb: WorkbookHandle; store: SpreadsheetStore }> {
  const next = await WorkbookHandle.loadBytes(wb.save());
  open.push(next);
  return { wb: next, store: bound(next) };
}

const onlyTable = (store: SpreadsheetStore): TableOverlay => {
  const tables = store.getState().tables.tables;
  expect(tables).toHaveLength(1);
  return tables[0] as TableOverlay;
};

describe('built-in table style names', () => {
  it('maps the gallery to real names and parses every block back', () => {
    expect(tableStyleName('medium', '#4472c4')).toBe('TableStyleMedium2');
    expect(tableStyleName('light', '#808080')).toBe('TableStyleLight1');
    expect(tableStyleName('dark', '#70ad47')).toBe('TableStyleDark7');
    expect(parseTableStyleName('TableStyleMedium9')).toEqual({ style: 'medium', color: '#4472c4' });
    expect(parseTableStyleName('TableStyleLight21')).toEqual({ style: 'light', color: '#70ad47' });
    expect(parseTableStyleName('TableStyleMedium29')).toBeNull();
    expect(parseTableStyleName('Mine')).toBeNull();
  });

  it('maps Dark8-11 to the neutral and accent pairs', () => {
    const color = (name: string) => parseTableStyleName(name)?.color;
    expect(color('TableStyleDark8')).toBe('#808080');
    expect(color('TableStyleDark9')).toBe('#4472c4');
    expect(color('TableStyleDark10')).toBe('#a5a5a5');
    expect(color('TableStyleDark11')).toBe('#5b9bd5');
  });

  it('writes a new table as Medium2', () => {
    expect(
      overlayStyleName(defaultTableOverlay('t', { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 })),
    ).toBe('TableStyleMedium2');
  });
});

describe('table styles through the engine', () => {
  it('hydrates style, flags and rows of a table built in the engine', async () => {
    const wb = await seeded();
    expect(
      wb.createTable({
        sheetIndex: 0,
        ref: 'A1:C4',
        name: 'Sales',
        columns: ['Item', 'Region', 'Qty'],
        styleName: 'TableStyleMedium3',
        showRowStripes: false,
        showFirstColumn: true,
        showLastColumn: true,
        showColumnStripes: true,
        totalsRow: true,
      }),
    ).toBeGreaterThanOrEqual(0);

    const { store } = await reloaded(wb);

    expect(onlyTable(store)).toMatchObject({
      source: 'engine',
      name: 'Sales',
      style: 'medium',
      color: '#ed7d31',
      styleName: 'TableStyleMedium3',
      showHeader: true,
      showTotal: true,
      banded: false,
      firstCol: true,
      lastCol: true,
      colStripes: true,
    });
  });

  it('round-trips the show flags and the style chosen in the gallery', async () => {
    const wb = await seeded();
    const store = bound(wb);
    const overlay = formatAsTable(store, range, {
      workbook: wb,
      style: 'dark',
      color: '#ffc000',
      banded: false,
      firstCol: true,
      lastCol: true,
    });
    updateTableOverlay(store, overlay?.id ?? '', { colStripes: true });

    expect(wb.getTableStyleDetails(0)).toMatchObject({
      styleName: 'TableStyleDark5',
      showRowStripes: false,
      showFirstColumn: true,
      showLastColumn: true,
      showColumnStripes: true,
    });
    const back = await reloaded(wb);
    expect(onlyTable(back.store)).toMatchObject({
      style: 'dark',
      color: '#ffc000',
      banded: false,
      firstCol: true,
      lastCol: true,
      colStripes: true,
    });

    updateTableOverlay(back.store, onlyTable(back.store).id, { banded: true, firstCol: false });
    const again = await reloaded(back.wb);
    expect(onlyTable(again.store)).toMatchObject({ banded: true, firstCol: false, lastCol: true });
  });

  it('keeps an unrecognised built-in style name when only a flag changes', async () => {
    const wb = await seeded();
    wb.createTable({
      sheetIndex: 0,
      ref: 'A1:C4',
      name: 'Sales',
      columns: ['Item', 'Region', 'Qty'],
      styleName: 'TableStyleLight16',
    });
    const store = bound(wb);

    updateTableOverlay(store, onlyTable(store).id, { lastCol: true });

    expect(wb.getTableStyleDetails(0)).toMatchObject({
      styleName: 'TableStyleLight16',
      showLastColumn: true,
    });
  });

  it('writes a custom table style, assigns it by name, and loads it back', async () => {
    const wb = await seeded();
    const store = bound(wb);
    const history = new History();
    const overlay = formatAsTable(store, range, { workbook: wb });
    store.setState((s) => ({
      ...s,
      selection: { ...s.selection, active: { sheet: 0, row: 1, col: 1 } },
    }));

    createTableStyleFromActiveTable(store, history, 'Quarterly', {
      style: 'dark',
      color: '#70ad47',
      variant: 'bandedFirstCol',
    });
    recordTablesChange(history, store, () => {
      updateTableOverlay(
        store,
        overlay?.id ?? '',
        tableStyleOptionsById(store.getState(), 'custom-table:Quarterly'),
      );
    });

    expect(wb.getCustomTableStyles().map((s) => s.name)).toEqual(['Quarterly']);
    expect(wb.getTableStyleDetails(0)).toMatchObject({
      styleName: 'Quarterly',
      showRowStripes: true,
      showFirstColumn: true,
    });

    const back = await reloaded(wb);
    expect(back.store.getState().tables.customTableStyles).toEqual([
      {
        id: 'custom-table:Quarterly',
        label: 'Quarterly',
        style: 'dark',
        color: '#70ad47',
        variant: 'bandedFirstCol',
      },
    ]);
    expect(onlyTable(back.store)).toMatchObject({
      style: 'dark',
      color: '#70ad47',
      styleName: 'Quarterly',
      banded: true,
      firstCol: true,
    });
  });

  it('removes a custom style from the file when its creation is undone', async () => {
    const wb = await seeded();
    const store = bound(wb);
    const history = new History();
    createTableStyleFromActiveTable(store, history, 'Temp', { style: 'light' });
    expect(wb.getCustomTableStyles()).toHaveLength(1);

    history.undo();
    expect(wb.getCustomTableStyles()).toHaveLength(0);

    history.redo();
    expect(wb.getCustomTableStyles().map((s) => s.name)).toEqual(['Temp']);
  });

  it('round-trips a PivotTable style assignment', async () => {
    const wb = await seeded();
    if (!wb.capabilities.pivotTableMutate || !wb.capabilities.pivotStyleInfo) return;
    const store = bound(wb);
    const created = createPivotTableFromRange(wb, {
      source: range,
      destination: { sheet: 0, row: 6, col: 0 },
      name: 'Report',
      rowField: 'Region',
      valueField: 'Qty',
      aggregation: PivotAggregation.Sum,
    });
    expect(created).toMatchObject({ ok: true });
    if (!created.ok) return;

    createPivotTableStyleFromActivePivot(
      store,
      null,
      'Ledger',
      { sheetIndex: 0, pivotIndex: created.pivotIndex },
      { style: 'medium', color: '#ed7d31', variant: 'banded' },
    );

    expect(wb.getPivotStyleInfo(0, created.pivotIndex)).toMatchObject({ name: 'Ledger' });
    const back = await reloaded(wb);
    expect(back.store.getState().tables.customPivotTableStyles).toEqual([
      {
        id: 'custom-pivot-table:Ledger',
        label: 'Ledger',
        style: 'medium',
        color: '#ed7d31',
        variant: 'banded',
      },
    ]);
    expect(back.store.getState().tables.pivotTableStyles).toEqual([
      { sheetIndex: 0, pivotIndex: created.pivotIndex, styleId: 'custom-pivot-table:Ledger' },
    ]);
  });

  it('renames a table in place so structured references follow', async () => {
    const wb = await seeded();
    const store = bound(wb);
    const overlay = formatAsTable(store, range, { workbook: wb });
    wb.setFormula({ sheet: 0, row: 6, col: 0 }, '=SUM(Table1[Qty])');
    expect(wb.getValue({ sheet: 0, row: 6, col: 0 })).toEqual({ kind: 'number', value: 6 });

    expect(updateTableOverlay(store, overlay?.id ?? '', { name: 'Orders' })).not.toBeNull();

    expect(wb.getTables().map((t) => t.name)).toEqual(['Orders']);
    expect(wb.cellFormula({ sheet: 0, row: 6, col: 0 })).toContain('Orders[Qty]');
    expect(wb.getValue({ sheet: 0, row: 6, col: 0 })).toEqual({ kind: 'number', value: 6 });
    expect(onlyTable(store).id).toBe(overlay?.id);
    expect(wb.getTableAutoFilter(0)).not.toBeNull();
  });

  it('refuses a rename onto another table name', async () => {
    const wb = await seeded();
    const store = bound(wb);
    const first = formatAsTable(store, range, { workbook: wb });
    formatAsTable(store, { sheet: 0, r0: 10, c0: 0, r1: 12, c1: 1 }, { workbook: wb });

    expect(updateTableOverlay(store, first?.id ?? '', { name: 'table2' })).toBeNull();
    expect(wb.getTables().map((t) => t.name)).toEqual(['Table1', 'Table2']);
  });

  it('undoes a rename', async () => {
    const wb = await seeded();
    const store = bound(wb);
    const history = new History();
    const overlay = formatAsTable(store, range, { workbook: wb });
    recordTablesChange(history, store, () => {
      updateTableOverlay(store, overlay?.id ?? '', { name: 'Orders' });
    });
    expect(wb.getTables().map((t) => t.name)).toEqual(['Orders']);

    history.undo();

    expect(wb.getTables().map((t) => t.name)).toEqual(['Table1']);
  });

  it('gives a new total row a label and a sum, and keeps them through a re-create', async () => {
    const wb = await seeded();
    const store = bound(wb);
    const overlay = formatAsTable(store, range, { workbook: wb, showTotal: true });

    const columns = wb.getTableStyleDetails(0)?.columns ?? [];
    expect(columns.map((c) => [c.totalsRowFunction, c.totalsRowLabel])).toEqual([
      ['none', 'Total'],
      ['none', ''],
      ['sum', ''],
    ]);

    // A wider range changes the column count, which re-creates the table.
    expect(
      updateTableOverlay(store, overlay?.id ?? '', { range: { ...range, c1: 3 } }),
    ).not.toBeNull();
    const after = wb.getTableStyleDetails(0);
    expect(after?.totalsRowCount).toBe(1);
    expect(after?.columns.find((c) => c.name === 'Qty')?.totalsRowFunction).toBe('sum');
  });
});
