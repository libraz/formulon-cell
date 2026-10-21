import { describe, expect, it } from 'vitest';
import { applyValueFilter, clearFilter, setAutoFilter } from '../../../src/commands/filter.js';
import { formatAsTable } from '../../../src/commands/format-as-table.js';
import { History } from '../../../src/commands/history.js';
import { hydrateAutoFilterFromEngine } from '../../../src/engine/auto-filter-sync.js';
import { hydrateLayoutFromEngine } from '../../../src/engine/layout-sync.js';
import { hydrateTableOverlaysFromEngine } from '../../../src/engine/table-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { toggleAutoFilterFromSelection } from '../../../src/interact/ribbon-filter-data.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const salesRange = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 };
const stockRange = { sheet: 0, r0: 0, c0: 3, r1: 3, c1: 4 };
const sheetRange = { sheet: 0, r0: 0, c0: 6, r1: 3, c1: 6 };

const GRID: string[][] = [
  ['Item', 'Region', '', 'Part', 'Bin', '', 'Note'],
  ['A', 'East', '', 'P1', 'X', '', 'keep'],
  ['B', 'West', '', 'P2', 'Y', '', 'drop'],
  ['C', 'East', '', 'P3', 'X', '', 'keep'],
];

function seed(wb: WorkbookHandle): void {
  GRID.forEach((cols, row) => {
    cols.forEach((text, col) => {
      if (text) wb.setText({ sheet: 0, row, col }, text);
    });
  });
}

/** Workbook with tables Sales (A1:B4) and Stock (D1:E4) side by side. */
async function twoTableWorkbook(): Promise<WorkbookHandle> {
  const wb = await WorkbookHandle.createDefault();
  seed(wb);
  expect(
    wb.createTable({ sheetIndex: 0, ref: 'A1:B4', name: 'Sales', columns: ['Item', 'Region'] }),
  ).toBe(0);
  expect(
    wb.createTable({ sheetIndex: 0, ref: 'D1:E4', name: 'Stock', columns: ['Part', 'Bin'] }),
  ).toBe(1);
  return wb;
}

function hydratedStore(wb: WorkbookHandle) {
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  hydrateLayoutFromEngine(wb, store, 0);
  hydrateTableOverlaysFromEngine(wb, store);
  hydrateAutoFilterFromEngine(wb, store, 0);
  return store;
}

describe('several AutoFilters on one sheet', () => {
  it('filters one table without dropping the other table filter', async () => {
    const wb = await twoTableWorkbook();
    try {
      const store = hydratedStore(wb);
      expect(store.getState().ui.filterRanges).toEqual([salesRange, stockRange]);

      applyValueFilter(store.getState(), store, stockRange, 4, ['Y'], wb);
      expect(wb.getTableAutoFilter(0)).not.toBeNull();
      expect(wb.getTableAutoFilter(1)?.columns[0]).toMatchObject({ colId: 1, values: ['X'] });
      expect(store.getState().ui.filterRanges).toEqual([salesRange, stockRange]);
      expect(store.getState().layout.hiddenRows).toEqual(new Set([2]));

      const restored = await WorkbookHandle.loadBytes(wb.save());
      try {
        expect(restored.getTableAutoFilter(0)).not.toBeNull();
        const restoredStore = hydratedStore(restored);
        expect(restoredStore.getState().ui.filterRanges).toEqual([salesRange, stockRange]);
        expect(restoredStore.getState().ui.filterCriteria).toEqual([
          { range: stockRange, byCol: 4, hiddenValues: ['Y'] },
        ]);
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('toggles only the filter of the table holding the active cell', async () => {
    const wb = await twoTableWorkbook();
    try {
      const store = hydratedStore(wb);
      mutators.setActive(store, { sheet: 0, row: 2, col: 4 });
      toggleAutoFilterFromSelection(store, new History(), () => wb);
      expect(store.getState().ui.filterRanges).toEqual([salesRange]);
      expect(wb.getTableAutoFilter(0)).not.toBeNull();
      expect(wb.getTableAutoFilter(1)).toBeNull();
    } finally {
      wb.dispose();
    }
  });

  it('keeps a worksheet filter and a table filter through a round-trip', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      seed(wb);
      expect(
        wb.createTable({ sheetIndex: 0, ref: 'A1:B4', name: 'Sales', columns: ['Item', 'Region'] }),
      ).toBe(0);
      const store = hydratedStore(wb);
      setAutoFilter(store, sheetRange, wb);
      applyValueFilter(store.getState(), store, sheetRange, 6, ['drop'], wb);
      expect(store.getState().ui.filterRanges).toEqual([salesRange, sheetRange]);
      expect(wb.getAutoFilter(0)).not.toBeNull();
      expect(wb.getTableAutoFilter(0)).not.toBeNull();

      const restored = await WorkbookHandle.loadBytes(wb.save());
      try {
        const restoredStore = hydratedStore(restored);
        expect(restoredStore.getState().ui.filterRanges).toEqual([sheetRange, salesRange]);
        expect(restoredStore.getState().ui.filterCriteria).toEqual([
          { range: sheetRange, byCol: 6, hiddenValues: ['drop'] },
        ]);

        clearFilter(restoredStore.getState(), restoredStore, sheetRange, restored);
        expect(restoredStore.getState().ui.filterRanges).toEqual([salesRange]);
        expect(restored.getAutoFilter(0)).toBeNull();
        expect(restored.getTableAutoFilter(0)).not.toBeNull();
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('shows the filter buttons of a table made by Format as Table at once', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      seed(wb);
      const store = hydratedStore(wb);
      wb.attachStore(store);
      expect(store.getState().ui.filterRanges).toEqual([]);
      expect(formatAsTable(store, salesRange, { workbook: wb })).not.toBeNull();
      expect(wb.getTableAutoFilter(0)).not.toBeNull();
      expect(store.getState().ui.filterRanges).toEqual([salesRange]);
    } finally {
      wb.dispose();
    }
  });
});
