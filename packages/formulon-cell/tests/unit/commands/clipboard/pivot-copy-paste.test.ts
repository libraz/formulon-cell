import { describe, expect, it } from 'vitest';
import { createCopyFormatResolver } from '../../../../src/commands/clipboard/format-resolver.js';
import { pasteSpecial } from '../../../../src/commands/clipboard/paste-special.js';
import { captureSnapshot } from '../../../../src/commands/clipboard/snapshot.js';
import { createPivotTableFromRange } from '../../../../src/commands/pivot-table.js';
import { builtinPivotStyleElements } from '../../../../src/engine/pivot-style.js';
import { PivotAggregation } from '../../../../src/engine/types.js';
import { addrKey, WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';
import { setActive } from './paste-special/fixtures.js';

const ALL = { what: 'all', operation: 'none', skipBlanks: false, transpose: false } as const;

/** Source Region/Amt/Qty at K1:M4 and a report at A3: rows Region, values Qty then Amt (A3:C7). */
const setup = async () => {
  const wb = await WorkbookHandle.createDefault();
  const rows: (string | number)[][] = [
    ['Region', 'Amt', 'Qty'],
    ['East', 500, 5],
    ['North', 300, 3],
    ['South', 700, 7],
  ];
  for (const [r, row] of rows.entries()) {
    for (const [c, v] of row.entries()) {
      const addr = { sheet: 0, row: r, col: 10 + c };
      if (typeof v === 'number') wb.setNumber(addr, v);
      else wb.setText(addr, v);
    }
  }
  const created = createPivotTableFromRange(wb, {
    source: { sheet: 0, r0: 0, c0: 10, r1: 3, c1: 12 },
    destination: { sheet: 0, row: 2, col: 0 },
    rowField: 'Region',
    valueField: 'Qty',
    valueFields: ['Qty', 'Amt'],
    aggregation: PivotAggregation.Sum,
  });
  expect(created.ok).toBe(true);
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  return { wb, store };
};

const copyAndPaste = (
  wb: WorkbookHandle,
  store: ReturnType<typeof createSpreadsheetStore>,
  range: { r0: number; c0: number; r1: number; c1: number },
  at: { row: number; col: number },
) => {
  const state = store.getState();
  const snap = captureSnapshot(state, { sheet: 0, ...range }, 'copy', {
    formatResolver: createCopyFormatResolver(wb, state),
  });
  if (!snap) throw new Error('no snapshot');
  setActive(store, at.row, at.col);
  const result = pasteSpecial(store.getState(), store, wb, snap, ALL);
  mutators.replaceCells(store, wb.cells(0));
  return result;
};

describe('copying part of a PivotTable', () => {
  it('pastes values with the style flattened into direct formats (A4:B7 -> E4)', async () => {
    const { wb, store } = await setup();
    const result = copyAndPaste(wb, store, { r0: 3, c0: 0, r1: 6, c1: 1 }, { row: 3, col: 4 });
    expect(result?.writtenRange).toMatchObject({ r0: 3, c0: 4, r1: 6, c1: 5 });
    expect(wb.getPivotTables()).toHaveLength(1);
    expect(wb.getValue({ sheet: 0, row: 6, col: 4 })).toEqual({
      kind: 'text',
      value: 'Grand Total',
    });
    expect(wb.getValue({ sheet: 0, row: 6, col: 5 })).toEqual({ kind: 'number', value: 15 });
    const style = builtinPivotStyleElements('PivotStyleLight16', wb.getTheme()?.colors);
    const formats = store.getState().format.formats;
    for (const col of [4, 5]) {
      expect(formats.get(addrKey({ sheet: 0, row: 6, col }))).toMatchObject({
        fill: style?.totalRow?.fill,
        bold: true,
        borders: { top: style?.totalRow?.borders?.top },
      });
      for (const row of [3, 4, 5]) {
        const f = formats.get(addrKey({ sheet: 0, row, col }));
        expect(f?.fill).toBeUndefined();
        expect(f?.bold).toBeFalsy();
        expect(f?.borders).toBeUndefined();
      }
    }
    wb.dispose();
  });
});

describe('copying a whole PivotTable', () => {
  it('creates a second report with the same layout and no direct formats (A3:C7 -> E10)', async () => {
    const { wb, store } = await setup();
    const widthBefore = store.getState().layout.colWidths.get(4);
    const result = copyAndPaste(wb, store, { r0: 2, c0: 0, r1: 6, c1: 2 }, { row: 9, col: 4 });
    expect(result?.writtenRange).toMatchObject({ r0: 9, c0: 4, r1: 13, c1: 6 });
    const pivots = wb.getPivotTables();
    expect(pivots).toHaveLength(2);
    const copy = pivots.find((p) => p.top === 9 && p.left === 4);
    expect(copy).toMatchObject({ rows: 5, cols: 3 });
    if (!copy) return;
    expect(wb.pivotTableName(0, copy.pivotIndex)).toBe('PivotTable2');
    const source = wb.getPivotDefinition(0, 0);
    const pasted = wb.getPivotDefinition(0, copy.pivotIndex);
    expect(pasted?.rowFieldOrder).toEqual(source?.rowFieldOrder);
    expect(pasted?.dataFields).toEqual(source?.dataFields);
    expect(pasted?.style).toEqual(source?.style);
    // Projected report cells, as the grid shows them.
    expect(store.getState().data.cells.get(addrKey({ sheet: 0, row: 13, col: 6 }))?.value).toEqual({
      kind: 'number',
      value: 1500,
    });
    const formats = store.getState().format.formats;
    for (const [row, col] of [
      [9, 4],
      [13, 4],
      [13, 5],
    ] as const) {
      expect(formats.get(addrKey({ sheet: 0, row, col }))?.fill).toBeUndefined();
    }
    expect(store.getState().layout.colWidths.get(4)).toBe(widthBefore);
    wb.dispose();
  });
});
