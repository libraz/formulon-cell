import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  createPivotTableStyleFromActivePivot,
  createTableStyleFromActiveTable,
  defaultTableOverlay,
  formatAsTable,
} from '../../../src/commands/format-as-table.js';
import { History } from '../../../src/commands/history.js';
import { renameActiveTable } from '../../../src/commands/table-design.js';
import { TableMirror } from '../../../src/engine/table-mirror.js';
import { hydrateTableOverlaysFromEngine } from '../../../src/engine/table-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const open: WorkbookHandle[] = [];
afterEach(() => {
  vi.restoreAllMocks();
  for (const wb of open.splice(0)) wb.dispose();
});

async function setup() {
  const wb = await WorkbookHandle.createDefault();
  open.push(wb);
  for (const [row, cols] of [
    ['A', 'B'],
    ['1', '2'],
    ['3', '4'],
  ].entries()) {
    for (const [col, text] of cols.entries()) {
      if (row === 0) wb.setText({ sheet: 0, row, col }, text);
      else wb.setNumber({ sheet: 0, row, col }, Number(text));
    }
  }
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  hydrateTableOverlaysFromEngine(wb, store);
  wb.attachStore(store);
  formatAsTable(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 }, { workbook: wb });
  mutators.setActive(store, { sheet: 0, row: 1, col: 0 });
  return { wb, store, history: new History() };
}

describe('table styles and refresh edge cases', () => {
  it('does not call a protected-sheet rename a duplicate name', async () => {
    const { store, history } = await setup();
    mutators.setSheetProtected(store, 0, true);
    expect(renameActiveTable(store, history, 'Fresh')).toBe('protected');
    expect(store.getState().tables.tables[0]?.name).toBe('Table1');
  });

  it('drops a custom table style the engine refuses', async () => {
    const { wb, store, history } = await setup();
    vi.spyOn(wb, 'setCustomTableStyle').mockReturnValue(false);
    createTableStyleFromActiveTable(store, history, 'Mine');
    expect(store.getState().tables.customTableStyles).toEqual([]);
  });

  it('keeps a custom table style the engine accepts', async () => {
    const { wb, store, history } = await setup();
    createTableStyleFromActiveTable(store, history, 'Mine');
    expect(store.getState().tables.customTableStyles?.map((s) => s.label)).toEqual(['Mine']);
    expect(wb.getCustomTableStyles().map((s) => s.name)).toEqual(['Mine']);
  });

  it('rejects a pivot style that reuses a table style name with another look', async () => {
    const { wb, store, history } = await setup();
    createTableStyleFromActiveTable(store, history, 'Same', { style: 'medium' });
    createPivotTableStyleFromActivePivot(store, history, 'Same', null, { style: 'dark' });
    expect(store.getState().tables.customPivotTableStyles).toEqual([]);
    expect(store.getState().tables.customTableStyles?.map((s) => s.label)).toEqual(['Same']);
    expect(wb.getCustomTableStyles()[0]).toMatchObject({ name: 'Same', table: true, pivot: false });
  });

  it('shares one engine style between same-named table and pivot styles that look alike', async () => {
    const { wb, store, history } = await setup();
    createTableStyleFromActiveTable(store, history, 'Same', { style: 'medium' });
    createPivotTableStyleFromActivePivot(store, history, 'Same', null, { style: 'medium' });
    expect(wb.getCustomTableStyles()[0]).toMatchObject({ table: true, pivot: true });
  });

  it('retries from the old baseline when a table write throws', async () => {
    const wb = await WorkbookHandle.createDefault();
    open.push(wb);
    wb.setText({ sheet: 0, row: 0, col: 0 }, 'A');
    wb.setText({ sheet: 0, row: 0, col: 1 }, 'B');
    wb.setNumber({ sheet: 0, row: 1, col: 0 }, 1);
    wb.setNumber({ sheet: 0, row: 1, col: 1 }, 2);
    const store = createSpreadsheetStore();
    const mirror = new TableMirror(wb);
    mirror.rebase(store.getState().tables);
    mutators.upsertTableOverlay(store, {
      ...defaultTableOverlay('t1', { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 }),
      source: 'engine',
      name: 'Retry',
    });
    const create = vi.spyOn(wb, 'createTable').mockImplementationOnce(() => {
      throw new Error('engine down');
    });
    expect(() => mirror.sync(store)).toThrow('engine down');
    expect(wb.getTables()).toEqual([]);
    mirror.sync(store);
    expect(create).toHaveBeenCalledTimes(2);
    expect(wb.getTables().map((t) => t.name)).toEqual(['Retry']);
  });

  it('leaves manual calculation mode alone when refreshing totals-row cells', async () => {
    const wb = await WorkbookHandle.createDefault();
    open.push(wb);
    wb.setCalcMode(1);
    for (const [row, values] of [
      ['A', 'B'],
      [1, 2],
      [3, 4],
      [0, 0],
    ].entries()) {
      for (const [col, value] of values.entries()) {
        if (typeof value === 'string') wb.setText({ sheet: 0, row, col }, value);
        else wb.setNumber({ sheet: 0, row, col }, value);
      }
    }
    const store = createSpreadsheetStore();
    const mirror = new TableMirror(wb);
    mirror.rebase(store.getState().tables);
    mutators.upsertTableOverlay(store, {
      ...defaultTableOverlay('t1', { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 }),
      source: 'engine',
      name: 'Totals',
      showTotal: true,
    });
    const full = vi.spyOn(wb, 'recalc');
    mirror.sync(store);
    expect(full).not.toHaveBeenCalled();
  });
});
