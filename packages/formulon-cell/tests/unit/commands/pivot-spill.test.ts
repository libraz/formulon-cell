import { describe, expect, it } from 'vitest';
import {
  createPivotTableFromRange,
  refreshPivotReportsAfterEdit,
} from '../../../src/commands/pivot-table.js';
import { pivotStyleLayer } from '../../../src/engine/pivot-style.js';
import { PivotAggregation, PivotAxis } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const PRODUCT = 1;

/** Region / Product / Amt at K1:M5; report at A3 (rows Region): A3:B7; "X" at A10. */
const setup = async () => {
  const wb = await WorkbookHandle.createDefault();
  const rows: (string | number)[][] = [
    ['Region', 'Product', 'Amt'],
    ['East', 'A', 500],
    ['North', 'A', 100],
    ['North', 'B', 200],
    ['South', 'B', 300],
  ];
  for (const [r, row] of rows.entries()) {
    for (const [c, v] of row.entries()) {
      const addr = { sheet: 0, row: r, col: 10 + c };
      if (typeof v === 'number') wb.setNumber(addr, v);
      else wb.setText(addr, v);
    }
  }
  const created = createPivotTableFromRange(wb, {
    source: { sheet: 0, r0: 0, c0: 10, r1: 4, c1: 12 },
    destination: { sheet: 0, row: 2, col: 0 },
    rowField: 'Region',
    valueField: 'Amt',
    aggregation: PivotAggregation.Sum,
  });
  expect(created.ok).toBe(true);
  wb.setText({ sheet: 0, row: 9, col: 0 }, 'X');
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  pivotStyleLayer(wb, store.getState());
  return { wb, store };
};

const grow = (wb: WorkbookHandle) => {
  expect(wb.setPivotFieldAxis(0, 0, PRODUCT, PivotAxis.Row)).toBe(true);
};

describe('a report that would grow over occupied cells', () => {
  it('collapses to #SPILL! at its anchor and keeps the blocking cell', async () => {
    const { wb, store } = await setup();
    grow(wb);
    const result = refreshPivotReportsAfterEdit(store, wb);
    expect(result.spilled).toEqual([
      {
        pivotIndex: 0,
        anchor: { sheet: 0, row: 2, col: 0 },
        needed: { sheet: 0, r0: 2, c0: 0, r1: 10, c1: 1 },
        blockers: [{ sheet: 0, row: 9, col: 0 }],
      },
    ]);
    expect(wb.pivotBounds(0)).toEqual([
      { pivotIndex: 0, range: { sheet: 0, r0: 2, c0: 0, r1: 2, c1: 0 } },
    ]);
    const cells = [...wb.pivotCells(0)];
    expect(cells.map((c) => [c.addr.row, c.addr.col, c.value])).toEqual([
      [2, 0, { kind: 'error', code: 8, text: '#SPILL!' }],
    ]);
    const shown = store.getState().data.cells;
    expect(shown.get('0:9:0')?.value).toEqual({ kind: 'text', value: 'X' });
    expect(shown.get('0:3:0')?.value.kind ?? 'blank').toBe('blank');
    // The blocking cell can still be edited: it is not part of a report.
    expect(wb.pivotOverlapsRange({ sheet: 0, r0: 9, c0: 0, r1: 9, c1: 0 })).toBe(false);
    wb.dispose();
  });

  it('comes back in full once the blocking cell is cleared', async () => {
    const { wb, store } = await setup();
    grow(wb);
    refreshPivotReportsAfterEdit(store, wb);
    wb.setBlank({ sheet: 0, row: 9, col: 0 });
    expect(wb.pivotSpills(0)).toEqual([]);
    expect(wb.pivotBounds(0)[0]?.range).toEqual({ sheet: 0, r0: 2, c0: 0, r1: 10, c1: 1 });
    wb.dispose();
  });

  it('grows freely over empty cells', async () => {
    const { wb, store } = await setup();
    wb.setBlank({ sheet: 0, row: 9, col: 0 });
    grow(wb);
    expect(refreshPivotReportsAfterEdit(store, wb).spilled).toEqual([]);
    expect(wb.pivotBounds(0)[0]?.range.r1).toBe(10);
    wb.dispose();
  });
});

describe('a report growing back over the cells it cached', () => {
  it('is not blocked by its own values after shrinking', async () => {
    const { wb } = await setup();
    wb.setBlank({ sheet: 0, row: 9, col: 0 });
    // The report's values are also in the sheet's own cells, as in a saved file.
    for (const c of [...wb.pivotCells(0)]) {
      if (c.value.kind === 'text') wb.setText(c.addr, c.value.value);
      else if (c.value.kind === 'number') wb.setNumber(c.addr, c.value.value);
    }
    expect(wb.setPivotTableGrandTotals(0, 0, true, false)).toBe(true);
    expect(wb.pivotBounds(0)[0]?.range.r1).toBe(5);
    expect(wb.setPivotTableGrandTotals(0, 0, true, true)).toBe(true);
    expect(wb.pivotSpills(0)).toEqual([]);
    expect(wb.pivotBounds(0)[0]?.range.r1).toBe(6);
    wb.dispose();
  });

  it('is blocked by a value typed into the rows it gave up, once the edit settled', async () => {
    const { wb, store } = await setup();
    wb.setBlank({ sheet: 0, row: 9, col: 0 });
    wb.setPivotTableGrandTotals(0, 0, true, false);
    refreshPivotReportsAfterEdit(store, wb);
    wb.setText({ sheet: 0, row: 6, col: 0 }, 'Y');
    wb.setPivotTableGrandTotals(0, 0, true, true);
    expect(refreshPivotReportsAfterEdit(store, wb).spilled[0]?.blockers).toEqual([
      { sheet: 0, row: 6, col: 0 },
    ]);
    wb.dispose();
  });
});

describe('a report moved by a row deletion', () => {
  it('moves with its cached cells and takes a report filter afterwards', async () => {
    const { wb, store } = await setup();
    wb.setBlank({ sheet: 0, row: 9, col: 0 });
    // The report's values are also in the sheet's own cells, as in a saved file.
    for (const c of [...wb.pivotCells(0)]) {
      if (c.value.kind === 'text') wb.setText(c.addr, c.value.value);
      else if (c.value.kind === 'number') wb.setNumber(c.addr, c.value.value);
    }
    expect(wb.engineDeleteRows(0, 0, 2)).toBe(true);
    expect(wb.pivotSpills(0)).toEqual([]);
    expect(wb.pivotBounds(0)[0]?.range).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 4, c1: 1 });
    mutators.replaceCells(store, wb.cells(0));
    pivotStyleLayer(wb, store.getState());
    expect(wb.setPivotFieldAxis(0, 0, PRODUCT, PivotAxis.Page)).toBe(true);
    expect(() => refreshPivotReportsAfterEdit(store, wb)).not.toThrow();
    // Filter block A1:B1, blank row 2, report proper from A3.
    expect(wb.pivotBounds(0)[0]?.range.r0).toBe(0);
    expect(wb.pivotCellInfo({ sheet: 0, row: 2, col: 0 })?.kind).toBe(0);
    // The cells the report cached at A1:B2 before moving down are gone.
    for (const [row, col] of [
      [0, 2],
      [1, 0],
      [1, 1],
    ] as const) {
      expect(wb.getValue({ sheet: 0, row, col })).toEqual({ kind: 'blank' });
    }
    wb.dispose();
  });
});
