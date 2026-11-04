import { describe, expect, it } from 'vitest';
import {
  clearPivotFieldFilter,
  pivotFieldFilterModel,
  setPivotFieldItemsChecked,
} from '../../../src/commands/pivot-filter.js';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { buildPivotStyleLayer } from '../../../src/engine/pivot-style.js';
import { PivotAggregation } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const REGION = 0;

/** Region / Amt at K1:L4; report at A3: rows Region, Sum of Amt (A3:B7). */
const setup = async () => {
  const wb = await WorkbookHandle.createDefault();
  const rows: (string | number)[][] = [
    ['Region', 'Amt'],
    ['East', 500],
    ['North', 300],
    ['South', 700],
  ];
  for (const [r, row] of rows.entries()) {
    for (const [c, v] of row.entries()) {
      const addr = { sheet: 0, row: r, col: 10 + c };
      if (typeof v === 'number') wb.setNumber(addr, v);
      else wb.setText(addr, v);
    }
  }
  const created = createPivotTableFromRange(wb, {
    source: { sheet: 0, r0: 0, c0: 10, r1: 3, c1: 11 },
    destination: { sheet: 0, row: 2, col: 0 },
    rowField: 'Region',
    valueField: 'Amt',
    aggregation: PivotAggregation.Sum,
  });
  expect(created.ok).toBe(true);
  return wb;
};

const grandTotal = (wb: WorkbookHandle) =>
  [...wb.pivotCells(0)].find((c) => c.kind === 6 && c.value.kind === 'number')?.value;

describe('pivot field filter', () => {
  it('lists the field items, all checked, and reports no filter', async () => {
    const wb = await setup();
    const model = pivotFieldFilterModel(wb, 0, 0, REGION);
    expect(model?.label).toBe('Region');
    expect(model?.items.map((i) => [i.value, i.checked])).toEqual([
      ['East', true],
      ['North', true],
      ['South', true],
    ]);
    expect(model?.filtered).toBe(false);
    wb.dispose();
  });

  it('hides an item, recomputes the grand total, and clears again', async () => {
    const wb = await setup();
    const model = pivotFieldFilterModel(wb, 0, 0, REGION);
    if (!model) throw new Error('no model');
    expect(
      setPivotFieldItemsChecked(
        wb,
        0,
        0,
        REGION,
        model.items.map((i) => ({ ...i, checked: i.value !== 'North' })),
      ),
    ).toBe(true);
    expect(grandTotal(wb)).toEqual({ kind: 'number', value: 1200 });
    const after = pivotFieldFilterModel(wb, 0, 0, REGION);
    expect(after?.filtered).toBe(true);
    expect(after?.items.find((i) => i.value === 'North')?.checked).toBe(false);
    // The caption button shows the filtered glyph.
    expect(buildPivotStyleLayer(wb, 0, {}).buttonAt(2, 0)).toEqual({
      pivotIndex: 0,
      fieldIndex: REGION,
      filtered: true,
    });
    expect(clearPivotFieldFilter(wb, 0, 0, REGION)).toBe(true);
    expect(grandTotal(wb)).toEqual({ kind: 'number', value: 1500 });
    expect(pivotFieldFilterModel(wb, 0, 0, REGION)?.filtered).toBe(false);
    expect(buildPivotStyleLayer(wb, 0, {}).buttonAt(2, 0)?.filtered).toBe(false);
    wb.dispose();
  });
});
