import { describe, expect, it } from 'vitest';
import { refreshPivotReportsAfterEdit } from '../../../src/commands/pivot-report-update.js';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { addrKey } from '../../../src/engine/address.js';
import { PIVOT_CELL_KIND } from '../../../src/engine/pivot-cell-kind.js';
import { pivotStyleLayer } from '../../../src/engine/pivot-style.js';
import { PivotAggregation, PivotAxis } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import type { CellFormat } from '../../../src/store/store.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const YELLOW = '#ffff00';
const FIXED2 = { kind: 'fixed', decimals: 2 } as const;
const PRODUCT = 1;

/** Region / Product / Amt / Qty at K1:N5; report at A3: rows Region, values Amt then Qty. */
const setup = async () => {
  const wb = await WorkbookHandle.createDefault();
  const rows: (string | number)[][] = [
    ['Region', 'Product', 'Amt', 'Qty'],
    ['East', 'A', 500, 5],
    ['North', 'A', 100, 1],
    ['North', 'B', 200, 2],
    ['South', 'A', 300, 3],
  ];
  for (const [r, row] of rows.entries()) {
    for (const [c, v] of row.entries()) {
      const addr = { sheet: 0, row: r, col: 10 + c };
      if (typeof v === 'number') wb.setNumber(addr, v);
      else wb.setText(addr, v);
    }
  }
  const created = createPivotTableFromRange(wb, {
    source: { sheet: 0, r0: 0, c0: 10, r1: 4, c1: 13 },
    destination: { sheet: 0, row: 2, col: 0 },
    rowField: 'Region',
    valueField: 'Amt',
    valueFields: ['Amt', 'Qty'],
    aggregation: PivotAggregation.Sum,
  });
  expect(created.ok).toBe(true);
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  return { wb, store };
};

const setFormat = (
  store: ReturnType<typeof createSpreadsheetStore>,
  row: number,
  col: number,
  patch: Partial<CellFormat>,
) =>
  store.setState((s) => {
    const formats = new Map(s.format.formats);
    const key = addrKey({ sheet: 0, row, col });
    formats.set(key, { ...formats.get(key), ...patch });
    return { ...s, format: { ...s.format, formats } };
  });

const fmt = (store: ReturnType<typeof createSpreadsheetStore>, row: number, col: number) =>
  store.getState().format.formats.get(addrKey({ sheet: 0, row, col }));

/** Where the report's cells are: label rows by item, data columns by field. */
const locate = (wb: WorkbookHandle) => {
  const cells = [...wb.pivotCells(0)];
  const labelRow = (item: string) =>
    cells.find(
      (c) =>
        c.kind === PIVOT_CELL_KIND.RowLabel && c.value.kind === 'text' && c.value.value === item,
    )?.addr.row ?? -1;
  const dataCols = (field: string) => [
    ...new Set(
      cells
        .filter((c) => c.kind === PIVOT_CELL_KIND.Data && c.fieldName.includes(field))
        .map((c) => c.addr.col),
    ),
  ];
  const label = (item: string) =>
    cells.find(
      (c) =>
        c.kind === PIVOT_CELL_KIND.RowLabel && c.value.kind === 'text' && c.value.value === item,
    )?.addr;
  return { labelRow, dataCols, label };
};

describe('direct formats on a report follow its items across a layout change', () => {
  it('moves and spreads them when a column field arrives, and puts them back when it goes', async () => {
    const { wb, store } = await setup();
    // A5 = North label, B5 = North x Amt, B4:B7 = the Amt column incl. the grand total.
    expect(locate(wb).labelRow('North')).toBe(4);
    setFormat(store, 4, 0, { fill: YELLOW });
    setFormat(store, 4, 1, { bold: true });
    for (const row of [3, 4, 5, 6]) setFormat(store, row, 1, { numFmt: FIXED2 });
    pivotStyleLayer(wb, store.getState());

    wb.setPivotFieldAxis(0, 0, PRODUCT, PivotAxis.Col);
    wb.setPivotColFieldOrder(0, 0, [PRODUCT]);
    refreshPivotReportsAfterEdit(store, wb);

    const grown = locate(wb);
    const north = grown.labelRow('North');
    const amtCols = grown.dataCols('Amt');
    const qtyCols = grown.dataCols('Qty');
    expect(amtCols.length).toBeGreaterThan(1);
    const label = grown.label('North');
    expect(label && fmt(store, label.row, label.col)?.fill).toBe(YELLOW);
    for (const col of amtCols) {
      expect(fmt(store, north, col)?.bold).toBe(true);
      expect(fmt(store, north, col)?.numFmt).toEqual(FIXED2);
    }
    for (const col of qtyCols) {
      expect(fmt(store, north, col)?.bold).toBeFalsy();
      expect(fmt(store, north, col)?.numFmt).toBeUndefined();
    }

    wb.setPivotFieldAxis(0, 0, PRODUCT, null);
    wb.setPivotColFieldOrder(0, 0, []);
    refreshPivotReportsAfterEdit(store, wb);

    expect(fmt(store, 4, 0)?.fill).toBe(YELLOW);
    expect(fmt(store, 4, 1)?.bold).toBe(true);
    expect(fmt(store, 3, 1)?.bold).toBeFalsy();
    for (const row of [3, 4, 5, 6]) expect(fmt(store, row, 1)?.numFmt).toEqual(FIXED2);
    expect(fmt(store, 4, 2)?.bold).toBeFalsy();
    // The columns the wider report vacated keep nothing.
    for (let col = 3; col <= 7; col += 1) {
      for (let row = 2; row <= 9; row += 1) expect(fmt(store, row, col)).toBeUndefined();
    }
    wb.dispose();
  });
});
