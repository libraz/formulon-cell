import { afterEach, describe, expect, it, vi } from 'vitest';
import type { SelectionStats } from '../../../src/commands/aggregate.js';
import { applyValueFilter } from '../../../src/commands/filter.js';
import {
  clearTable,
  clearTablesInRange,
  formatAsTable,
  type TableOverlay,
  updateTableOverlay,
} from '../../../src/commands/format-as-table.js';
import { History } from '../../../src/commands/history.js';
import { executeQuickAnalysisAction } from '../../../src/commands/quick-analysis.js';
import { recordTablesChange } from '../../../src/commands/slice-history.js';
import { hydrateAutoFilterFromEngine } from '../../../src/engine/auto-filter-sync.js';
import { hydrateTableOverlaysFromEngine } from '../../../src/engine/table-sync.js';
import type { Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';
import { zipEntries, zipText } from './zip-entries.js';

const salesRange: Range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 };

const GRID: string[][] = [
  ['Item', 'Region', 'Qty'],
  ['A', 'East', '1'],
  ['B', 'West', '2'],
  ['C', 'East', '3'],
  ['D', 'West', '4'],
];

const stats: SelectionStats = {
  cells: 1,
  numericCount: 0,
  nonBlankCount: 0,
  sum: 0,
  avg: 0,
  min: 0,
  max: 0,
};

const open: WorkbookHandle[] = [];

afterEach(() => {
  for (const wb of open.splice(0)) wb.dispose();
});

async function seededWorkbook(): Promise<WorkbookHandle> {
  const wb = await WorkbookHandle.createDefault();
  open.push(wb);
  GRID.forEach((cols, row) => {
    cols.forEach((text, col) => {
      wb.setText({ sheet: 0, row, col }, text);
    });
  });
  return wb;
}

function boundStore(wb: WorkbookHandle): SpreadsheetStore {
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  hydrateTableOverlaysFromEngine(wb, store);
  hydrateAutoFilterFromEngine(wb, store, 0);
  wb.attachStore(store);
  return store;
}

async function reload(wb: WorkbookHandle): Promise<WorkbookHandle> {
  const next = await WorkbookHandle.loadBytes(wb.save());
  open.push(next);
  return next;
}

const onlyTable = (store: SpreadsheetStore): TableOverlay => {
  const tables = store.getState().tables.tables;
  expect(tables).toHaveLength(1);
  return tables[0] as TableOverlay;
};

const tableXml = (wb: WorkbookHandle): string => {
  const entries = zipEntries(wb.save());
  const path = [...entries.keys()].find((key) => key.startsWith('xl/tables/'));
  return path ? zipText(entries, path) : '';
};

describe('table mirror', () => {
  it('creates the engine table and its filter buttons, and removes both on clear', async () => {
    const wb = await seededWorkbook();
    const store = boundStore(wb);

    const overlay = formatAsTable(store, salesRange, { workbook: wb });

    expect(overlay).toMatchObject({ source: 'engine', name: 'Table1' });
    expect(wb.getTables()).toMatchObject([{ name: 'Table1', ref: 'A1:B4' }]);
    expect(wb.getTableAutoFilter(0)).not.toBeNull();
    expect(store.getState().ui.filterRanges).toEqual([salesRange]);

    clearTable(store, overlay?.id ?? '');

    expect(wb.getTables()).toEqual([]);
    expect(store.getState().ui.filterRanges).toEqual([]);
    expect(wb.getAutoFilter(0)).toBeNull();
    const restored = await reload(wb);
    expect(restored.getTables()).toEqual([]);
    expect(restored.getAutoFilter(0)).toBeNull();
  });

  it('drops a removed table filter criteria with its range', async () => {
    const wb = await seededWorkbook();
    const store = boundStore(wb);
    formatAsTable(store, salesRange, { workbook: wb });
    applyValueFilter(store.getState(), store, salesRange, 1, ['West'], wb);
    expect(store.getState().ui.filterCriteria).toHaveLength(1);

    clearTablesInRange(store, salesRange);

    expect(store.getState().ui.filterRanges).toEqual([]);
    expect(store.getState().ui.filterCriteria).toEqual([]);
    expect(wb.getAutoFilter(0)).toBeNull();
  });

  it('undoes and redoes Format as Table in the engine, filter buttons included', async () => {
    const wb = await seededWorkbook();
    const store = boundStore(wb);
    const history = new History();
    recordTablesChange(history, store, () => {
      formatAsTable(store, salesRange, { workbook: wb });
    });
    expect(wb.getTables()).toHaveLength(1);

    history.undo();
    expect(store.getState().tables.tables).toEqual([]);
    expect(wb.getTables()).toEqual([]);
    expect(store.getState().ui.filterRanges).toEqual([]);

    history.redo();
    expect(wb.getTables()).toMatchObject([{ name: 'Table1', ref: 'A1:B4' }]);
    expect(wb.getTableAutoFilter(0)).not.toBeNull();
    expect(store.getState().ui.filterRanges).toEqual([salesRange]);
  });

  it('restyles an existing table in place', async () => {
    const wb = await seededWorkbook();
    const store = boundStore(wb);
    const overlay = formatAsTable(store, salesRange, { workbook: wb });
    const createTable = vi.spyOn(wb, 'createTable');
    const updateTable = vi.spyOn(wb, 'updateTable');

    expect(updateTableOverlay(store, overlay?.id ?? '', { style: 'dark' })).not.toBeNull();

    expect(updateTable).toHaveBeenCalledWith(0, { styleName: 'TableStyleDark1' });
    expect(createTable).not.toHaveBeenCalled();
    expect(wb.getTables()).toHaveLength(1);
    expect(tableXml(wb)).toContain('TableStyleDark1');
  });

  it('resizes a table, re-creating it when the column count changes', async () => {
    const wb = await seededWorkbook();
    const store = boundStore(wb);
    const overlay = formatAsTable(store, salesRange, { workbook: wb });
    const id = overlay?.id ?? '';

    const taller: Range = { ...salesRange, r1: 4 };
    expect(updateTableOverlay(store, id, { range: taller })).not.toBeNull();
    expect(wb.getTables()).toMatchObject([{ name: 'Table1', ref: 'A1:B5' }]);
    expect(store.getState().ui.filterRanges).toEqual([taller]);

    const wider: Range = { ...taller, c1: 2 };
    expect(updateTableOverlay(store, id, { range: wider })).not.toBeNull();
    expect(wb.getTables()).toMatchObject([{ name: 'Table1', ref: 'A1:C5' }]);
    expect(wb.getTables()[0]?.columns).toEqual(['Item', 'Region', 'Qty']);
    expect(store.getState().ui.filterRanges).toEqual([wider]);
    expect(onlyTable(store).range).toEqual(wider);
  });

  it('removes the engine table on Clear Formats', async () => {
    const wb = await seededWorkbook();
    const store = boundStore(wb);
    formatAsTable(store, salesRange, { workbook: wb });

    const result = executeQuickAnalysisAction({
      store,
      wb,
      actionId: 'format-clear',
      range: salesRange,
      stats,
    });

    expect(result).toMatchObject({ ok: true });
    expect(store.getState().tables.tables).toEqual([]);
    expect(wb.getTables()).toEqual([]);
    expect(store.getState().ui.filterRanges).toEqual([]);
  });

  it('removes a table loaded from a file', async () => {
    const authored = await seededWorkbook();
    formatAsTable(boundStore(authored), salesRange, { workbook: authored });
    const wb = await reload(authored);
    const store = boundStore(wb);
    expect(onlyTable(store)).toMatchObject({ source: 'engine', name: 'Table1' });
    expect(store.getState().ui.filterRanges).toEqual([salesRange]);

    clearTablesInRange(store, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 });

    expect(wb.getTables()).toEqual([]);
    expect(store.getState().ui.filterRanges).toEqual([]);
    expect((await reload(wb)).getTables()).toEqual([]);
  });

  it('keeps the chosen look through a rehydrate without writing it back', async () => {
    const wb = await seededWorkbook();
    const store = boundStore(wb);
    const overlay = formatAsTable(store, salesRange, { workbook: wb, style: 'light' });
    updateTableOverlay(store, overlay?.id ?? '', { color: '#70ad47', banded: false });
    const updateTable = vi.spyOn(wb, 'updateTable');

    hydrateTableOverlaysFromEngine(wb, store);

    expect(onlyTable(store)).toMatchObject({
      id: overlay?.id,
      style: 'light',
      color: '#70ad47',
      banded: false,
      range: salesRange,
    });
    expect(updateTable).not.toHaveBeenCalled();
  });

  it('refuses a table over another table and keeps the screen in step', async () => {
    const wb = await seededWorkbook();
    const store = boundStore(wb);
    formatAsTable(store, salesRange, { workbook: wb });

    const overlapping = formatAsTable(
      store,
      { sheet: 0, r0: 1, c0: 1, r1: 4, c1: 2 },
      { workbook: wb },
    );

    expect(overlapping).toBeNull();
    expect(onlyTable(store).name).toBe('Table1');
    expect(wb.getTables()).toHaveLength(1);
  });
});
