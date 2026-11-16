import { describe, expect, it } from 'vitest';
import {
  autofitPivotTableColumns,
  createPivotTableFromRange,
  refreshPivotReportsAfterEdit,
  settlePivotAnchor,
} from '../../../src/commands/pivot-table.js';
import { buildPivotStyleLayer, pivotStyleLayer } from '../../../src/engine/pivot-style.js';
import { PivotAggregation, PivotAxis } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

/** Region / Product / Amt at K1:M5 and a report (rows Region, Sum of Amt) at `at`. */
const setup = async (at: { row: number; col: number }) => {
  const wb = await WorkbookHandle.createDefault();
  wb.setSpreadsheetProfileId('windows-ja_JP');
  const rows: (string | number)[][] = [
    ['Region', 'Product', 'Amt'],
    ['East', 'A', 500],
    ['North', 'A', 100],
    ['North', 'B', 200],
    ['South', 'C', 300],
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
    destination: { sheet: 0, ...at },
    rowField: 'Region',
    valueField: 'Amt',
    aggregation: PivotAggregation.Sum,
  });
  expect(created.ok).toBe(true);
  return wb;
};

const PRODUCT = 1;
const top = (wb: WorkbookHandle) => wb.pivotBounds(0)[0]?.range.r0;

describe('report filter area', () => {
  it('moves a report at A1 down to A3 when a filter arrives, and keeps it there', async () => {
    const wb = await setup({ row: 0, col: 0 });
    expect(wb.setPivotFieldAxis(0, 0, PRODUCT, PivotAxis.Page)).toBe(true);
    settlePivotAnchor(wb, 0, 0);
    // Filter block A1:B1, blank row 2, report from A3.
    expect(top(wb)).toBe(0);
    expect(wb.pivotCellInfo({ sheet: 0, row: 2, col: 0 })?.kind).toBe(0);
    expect(wb.setPivotFieldAxis(0, 0, PRODUCT, null)).toBe(true);
    settlePivotAnchor(wb, 0, 0);
    expect(top(wb)).toBe(2);
    wb.dispose();
  });

  it('leaves a report at A3 where it is', async () => {
    const wb = await setup({ row: 2, col: 0 });
    wb.setPivotFieldAxis(0, 0, PRODUCT, PivotAxis.Page);
    settlePivotAnchor(wb, 0, 0);
    expect(top(wb)).toBe(0);
    wb.setPivotFieldAxis(0, 0, PRODUCT, null);
    settlePivotAnchor(wb, 0, 0);
    expect(top(wb)).toBe(2);
    wb.dispose();
  });

  it('puts the drop-down on the selection cell and keeps the header style below the block', async () => {
    const wb = await setup({ row: 2, col: 0 });
    wb.setPivotFieldAxis(0, 0, PRODUCT, PivotAxis.Page);
    const layer = buildPivotStyleLayer(wb, 0, {});
    expect(layer.hasButton(0, 1)).toBe(true);
    expect(layer.hasButton(2, 0)).toBe(true);
    expect(layer.formatAt(0, 0)).toBeUndefined();
    expect(layer.formatAt(0, 1)).toBeUndefined();
    expect(layer.formatAt(2, 0)?.bold).toBe(true);
    // Gridlines drop inside the report proper only.
    expect(layer.ranges[0]).toMatchObject({ r0: 2 });
    wb.dispose();
  });

  it('shows the selection text and fits its column to it', async () => {
    const wb = await setup({ row: 2, col: 0 });
    wb.setPivotFieldAxis(0, 0, PRODUCT, PivotAxis.Page);
    const text = () =>
      [...wb.pivotCells(0)].find((c) => c.addr.row === 0 && c.addr.col === 1)?.value;
    expect(text()).toEqual({ kind: 'text', value: '(すべて)' });
    wb.clearPivotFieldItems(0, 0, PRODUCT);
    for (const [label, visible] of [
      ['A', true],
      ['B', true],
      ['C', false],
    ] as const) {
      wb.addPivotFieldItem(0, 0, PRODUCT, label, visible);
    }
    expect(text()).toEqual({ kind: 'text', value: '(複数のアイテム)' });
    const store = createSpreadsheetStore();
    mutators.replaceCells(store, wb.cells(0));
    mutators.setColWidth(store, 1, 20);
    autofitPivotTableColumns(store, wb);
    const fitted = store.getState().layout.colWidths.get(1) ?? 0;
    expect(fitted).toBeGreaterThan(20);
    expect(pivotStyleLayer(wb, store.getState()).hasButton(0, 1)).toBe(true);
    wb.dispose();
  });
});

describe('cells a report leaves behind', () => {
  /** A report whose old layout is also in the sheet's own cells, as in a saved file. */
  const withCachedCells = async (at: { row: number; col: number }) => {
    const wb = await setup(at);
    for (const c of [...wb.pivotCells(0)]) {
      if (c.value.kind === 'text') wb.setText(c.addr, c.value.value);
      else if (c.value.kind === 'number') wb.setNumber(c.addr, c.value.value);
    }
    const store = createSpreadsheetStore();
    mutators.replaceCells(store, wb.cells(0));
    pivotStyleLayer(wb, store.getState());
    return { wb, store };
  };
  const shown = (store: ReturnType<typeof createSpreadsheetStore>, row: number, col: number) =>
    store.getState().data.cells.get(`0:${row}:${col}`)?.value;

  it('clears the old extent when a column field widens the report', async () => {
    const { wb, store } = await withCachedCells({ row: 2, col: 0 });
    expect(shown(store, 3, 0)).toEqual({ kind: 'text', value: 'East' });
    wb.setPivotFieldAxis(0, 0, PRODUCT, PivotAxis.Col);
    wb.setPivotColFieldOrder(0, 0, [PRODUCT]);
    refreshPivotReportsAfterEdit(store, wb);
    const projected = new Set([...wb.pivotCells(0)].map((c) => `${c.addr.row}:${c.addr.col}`));
    for (const [key, cell] of store.getState().data.cells) {
      const [, row, col] = key.split(':').map(Number);
      if ((col ?? 0) >= 10 || cell.value.kind === 'blank') continue;
      expect(projected.has(`${row}:${col}`), `${key} ${JSON.stringify(cell.value)}`).toBe(true);
    }
    wb.dispose();
  });

  it('clears the rows a filter push-down vacates', async () => {
    const { wb, store } = await withCachedCells({ row: 0, col: 0 });
    wb.setPivotFieldAxis(0, 0, PRODUCT, PivotAxis.Page);
    refreshPivotReportsAfterEdit(store, wb);
    // Row 2 (index 1) is the blank separator now; the old East row is gone.
    expect(shown(store, 1, 0)?.kind ?? 'blank').toBe('blank');
    expect(shown(store, 1, 1)?.kind ?? 'blank').toBe('blank');
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'blank' });
    wb.dispose();
  });
});
