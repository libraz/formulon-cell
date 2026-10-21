import { describe, expect, it } from 'vitest';
import { applyValueFilter } from '../../../src/commands/filter.js';
import { History } from '../../../src/commands/history.js';
import {
  createSlicer,
  recomputeSlicerFilters,
  setSlicerSelected,
} from '../../../src/commands/slicers.js';
import {
  hydrateAutoFilterFromEngine,
  syncAutoFilterToEngine,
} from '../../../src/engine/auto-filter-sync.js';
import { hydrateLayoutFromEngine } from '../../../src/engine/layout-sync.js';
import { hydrateTableOverlaysFromEngine } from '../../../src/engine/table-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { toggleAutoFilterFromSelection } from '../../../src/interact/ribbon-filter-data.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const tableRange = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 };

const ROWS: [string, string][] = [
  ['Item', 'Region'],
  ['A', 'East'],
  ['B', 'West'],
  ['C', 'East'],
];

/** Workbook with one table over A1:B4; its sheet has no sheet-level filter. */
async function tableWorkbook(): Promise<WorkbookHandle> {
  const wb = await WorkbookHandle.createDefault();
  ROWS.forEach(([item, region], row) => {
    wb.setText({ sheet: 0, row, col: 0 }, item);
    wb.setText({ sheet: 0, row, col: 1 }, region);
  });
  expect(
    wb.createTable({ sheetIndex: 0, ref: 'A1:B4', name: 'Sales', columns: ['Item', 'Region'] }),
  ).toBe(0);
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

describe('table-level AutoFilter', () => {
  it('hydrates a table filter loaded from a workbook', async () => {
    const wb = await tableWorkbook();
    try {
      expect(wb.capabilities.typedTableAutoFilter).toBe(true);
      const store = hydratedStore(wb);
      expect(store.getState().ui.filterRange).toEqual(tableRange);
      applyValueFilter(store.getState(), store, tableRange, 1, ['West'], wb);
      expect(wb.getAutoFilter(0)).toBeNull();

      const restored = await WorkbookHandle.loadBytes(wb.save());
      try {
        const restoredStore = hydratedStore(restored);
        expect(restored.getAutoFilter(0)).toBeNull();
        expect(restoredStore.getState().ui.filterRange).toEqual(tableRange);
        expect(restoredStore.getState().ui.filterCriteria).toEqual([
          { range: tableRange, byCol: 1, hiddenValues: ['West'] },
        ]);
        expect(restoredStore.getState().layout.hiddenRows).toEqual(new Set([2]));
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('writes criteria edited through the command layer to the table, not the sheet', async () => {
    const wb = await tableWorkbook();
    try {
      const store = hydratedStore(wb);
      applyValueFilter(store.getState(), store, tableRange, 1, ['West'], wb);
      expect(wb.getAutoFilter(0)).toBeNull();
      expect(wb.getTableAutoFilter(0)?.columns[0]).toMatchObject({ colId: 1, values: ['East'] });
      expect(wb.evaluateTableAutoFilter(0)?.match).toEqual([true, false, true]);

      applyValueFilter(store.getState(), store, tableRange, 1, [], wb);
      expect(wb.getTableAutoFilter(0)?.columns).toEqual([]);
      expect(store.getState().layout.hiddenRows).toEqual(new Set());
    } finally {
      wb.dispose();
    }
  });

  it('toggles the table filter buttons from a cell inside the table', async () => {
    const wb = await tableWorkbook();
    try {
      const store = hydratedStore(wb);
      const history = new History();
      // A table starts with its filter buttons shown.
      expect(store.getState().ui.filterRange).toEqual(tableRange);
      expect(wb.getTableAutoFilter(0)).not.toBeNull();
      mutators.setActive(store, { sheet: 0, row: 2, col: 1 });

      toggleAutoFilterFromSelection(store, history, () => wb);
      expect(store.getState().ui.filterRange).toBeNull();
      expect(wb.getTableAutoFilter(0)).toBeNull();
      expect(wb.getAutoFilter(0)).toBeNull();

      toggleAutoFilterFromSelection(store, history, () => wb);
      expect(store.getState().ui.filterRange).toEqual(tableRange);
      expect(wb.getTableAutoFilter(0)).not.toBeNull();
      expect(wb.getAutoFilter(0)).toBeNull();
    } finally {
      wb.dispose();
    }
  });

  it('keeps a sheet filter alongside a table filter', async () => {
    const wb = await tableWorkbook();
    try {
      const sheetRange = { sheet: 0, r0: 0, c0: 4, r1: 2, c1: 4 };
      wb.setText({ sheet: 0, row: 0, col: 4 }, 'H');
      wb.setText({ sheet: 0, row: 1, col: 4 }, 'x');
      wb.setText({ sheet: 0, row: 2, col: 4 }, 'y');
      const store = hydratedStore(wb);
      store.setState((s) => ({ ...s, ui: { ...s.ui, filterRange: sheetRange } }));
      expect(syncAutoFilterToEngine(wb, store.getState(), 0, store)).toBe(true);
      expect(wb.getAutoFilter(0)).not.toBeNull();

      store.setState((s) => ({ ...s, ui: { ...s.ui, filterRange: tableRange } }));
      expect(syncAutoFilterToEngine(wb, store.getState(), 0, store)).toBe(true);
      expect(wb.getTableAutoFilter(0)).not.toBeNull();
      // Moving the UI filter to the table drops the sheet filter it owned.
      expect(wb.getAutoFilter(0)).toBeNull();
    } finally {
      wb.dispose();
    }
  });

  it('filters a table slicer through the table criteria', async () => {
    const wb = await tableWorkbook();
    try {
      const store = hydratedStore(wb);
      const made = createSlicer(store, wb, { tableName: 'Sales', column: 'Region' });
      expect(made.ok).toBe(true);
      if (!made.ok) return;
      setSlicerSelected(store, made.spec.id, ['East']);
      expect(recomputeSlicerFilters(store, wb)).toBe(1);
      expect(store.getState().ui.filterCriteria).toEqual([
        { range: tableRange, byCol: 1, hiddenValues: ['West'] },
      ]);
      syncAutoFilterToEngine(wb, store.getState(), 0, store);
      expect(wb.getAutoFilter(0)).toBeNull();
      expect(wb.getTableAutoFilter(0)?.columns[0]).toMatchObject({ colId: 1, values: ['East'] });
      expect(wb.evaluateTableAutoFilter(0)?.match).toEqual([true, false, true]);
    } finally {
      wb.dispose();
    }
  });
});
