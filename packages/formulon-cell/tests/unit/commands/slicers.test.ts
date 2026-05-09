import { describe, expect, it } from 'vitest';
import {
  clearSlicerSelection,
  createSlicer,
  findSlicerTable,
  listSlicers,
  listSlicerValues,
  recomputeSlicerFilters,
  removeSlicer,
  resolveSlicerSpec,
  setSlicerSelected,
  updateSlicer,
} from '../../../src/commands/slicers.js';
import type { CellValue } from '../../../src/engine/types.js';
import { addrKey, type WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const workbook = (): WorkbookHandle =>
  ({
    getTables: () => [
      {
        name: 'SalesTable',
        displayName: 'Sales',
        ref: 'A1:B4',
        sheetIndex: 0,
        columns: ['Item', 'Region'],
      },
    ],
  }) as unknown as WorkbookHandle;

const setCell = (store: SpreadsheetStore, row: number, col: number, value: CellValue): void => {
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    cells.set(addrKey({ sheet: 0, row, col }), { value, formula: null });
    return { ...s, data: { ...s.data, cells } };
  });
};

const seedRows = (store: SpreadsheetStore): void => {
  setCell(store, 0, 0, { kind: 'text', value: 'Item' });
  setCell(store, 0, 1, { kind: 'text', value: 'Region' });
  setCell(store, 1, 0, { kind: 'text', value: 'A' });
  setCell(store, 1, 1, { kind: 'text', value: 'East' });
  setCell(store, 2, 0, { kind: 'text', value: 'B' });
  setCell(store, 2, 1, { kind: 'text', value: 'West' });
  setCell(store, 3, 0, { kind: 'text', value: 'C' });
  setCell(store, 3, 1, { kind: 'text', value: 'East' });
};

describe('slicer commands', () => {
  it('creates a slicer from a table display name and resolves its range', () => {
    const store = createSpreadsheetStore();
    const wb = workbook();

    const result = createSlicer(store, wb, { tableName: 'Sales', column: 'Region' });

    expect(result).toEqual({
      ok: true,
      spec: {
        id: 'slicer-salestable-region',
        tableName: 'SalesTable',
        column: 'Region',
        selected: [],
        x: undefined,
        y: undefined,
      },
    });
    expect(findSlicerTable(wb, 'salestable')?.displayName).toBe('Sales');
    expect(result.ok && resolveSlicerSpec(wb, result.spec)).toEqual({
      range: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 },
      byCol: 1,
    });
    expect(listSlicers(store)).toEqual(result.ok ? [result.spec] : []);
  });

  it('returns explicit errors for missing table or column', () => {
    const store = createSpreadsheetStore();
    const wb = workbook();

    expect(createSlicer(store, wb, { tableName: 'Missing', column: 'Region' })).toEqual({
      ok: false,
      reason: 'table-not-found',
    });
    expect(createSlicer(store, wb, { tableName: 'Sales', column: 'Missing' })).toEqual({
      ok: false,
      reason: 'column-not-found',
    });
  });

  it('updates, clears, and removes slicer state', () => {
    const store = createSpreadsheetStore();
    const wb = workbook();
    const result = createSlicer(store, wb, { tableName: 'Sales', column: 'Region' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;

    setSlicerSelected(store, result.spec.id, ['East']);
    updateSlicer(store, result.spec.id, { x: 40, y: 60 });
    expect(listSlicers(store)[0]).toMatchObject({ selected: ['East'], x: 40, y: 60 });

    clearSlicerSelection(store, result.spec.id);
    expect(listSlicers(store)[0]?.selected).toEqual([]);

    removeSlicer(store, result.spec.id);
    expect(listSlicers(store)).toEqual([]);
  });

  it('lists distinct values and recomputes filter rows for selected chips', () => {
    const store = createSpreadsheetStore();
    const wb = workbook();
    seedRows(store);
    const result = createSlicer(store, wb, {
      tableName: 'Sales',
      column: 'Region',
      selected: ['East'],
    });
    expect(result.ok).toBe(true);
    if (!result.ok) return;

    expect(listSlicerValues(store, wb, result.spec)).toEqual(['East', 'West']);
    expect(recomputeSlicerFilters(store, wb)).toBe(1);

    expect(store.getState().layout.hiddenRows).toEqual(new Set([2]));
    expect(store.getState().ui.filterRange).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 });
  });
});
