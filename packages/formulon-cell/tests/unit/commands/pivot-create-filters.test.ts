import { describe, expect, it } from 'vitest';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import {
  PivotAggregation,
  PivotAxis,
  PivotFilterKind,
  type PivotFilterSpec,
} from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const SOURCE = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 2 };

const setup = async (): Promise<WorkbookHandle> => {
  const wb = await WorkbookHandle.createDefault();
  const rows: (string | number)[][] = [
    ['Region', 'Product', 'Amt'],
    ['East', 'A', 500],
    ['North', 'A', 100],
    ['North', 'B', 200],
    ['South', 'C', 300],
  ];
  for (const [r, row] of rows.entries()) {
    for (const [c, v] of row.entries()) {
      const addr = { sheet: 0, row: r, col: c };
      if (typeof v === 'number') wb.setNumber(addr, v);
      else wb.setText(addr, v);
    }
  }
  return wb;
};

const contains = (axis: PivotAxis, fieldName: string): PivotFilterSpec => ({
  axis,
  fieldName,
  kind: PivotFilterKind.LabelContains,
  value: 'o',
});

const create = (wb: WorkbookHandle, pivotFilters: readonly PivotFilterSpec[]) =>
  createPivotTableFromRange(wb, {
    source: SOURCE,
    destination: { sheet: 0, row: 8, col: 0 },
    rowField: 'Region',
    columnField: 'Product',
    filterField: 'Amt',
    pivotFilters,
    valueField: 'Amt',
    aggregation: PivotAggregation.Sum,
  });

describe('createPivotTableFromRange pivotFilters', () => {
  it('applies a spec to the row field', async () => {
    const wb = await setup();
    const result = create(wb, [contains(PivotAxis.Row, 'Region')]);
    expect(result.ok).toBe(true);
    const filters = wb.getPivotDefinition(0, 0)?.filters ?? [];
    expect(filters.map((f) => [f.fieldName, f.kind])).toEqual([
      ['Region', PivotFilterKind.LabelContains],
    ]);
    wb.dispose();
  });

  it('applies a spec to the column field', async () => {
    const wb = await setup();
    const result = create(wb, [contains(PivotAxis.Col, 'Product')]);
    expect(result.ok).toBe(true);
    const filters = wb.getPivotDefinition(0, 0)?.filters ?? [];
    expect(filters.map((f) => [f.fieldName, f.kind])).toEqual([
      ['Product', PivotFilterKind.LabelContains],
    ]);
    wb.dispose();
  });

  it('refuses a spec naming a report filter field', async () => {
    const wb = await setup();
    const result = create(wb, [contains(PivotAxis.Page, 'Amt')]);
    expect(result).toEqual({ ok: false, reason: 'invalid-field' });
    expect(wb.getPivotDefinition(0, 0)).toBeNull();
    wb.dispose();
  });

  it('refuses a spec naming a field outside the report', async () => {
    const wb = await setup();
    const result = createPivotTableFromRange(wb, {
      source: SOURCE,
      destination: { sheet: 0, row: 8, col: 0 },
      rowField: 'Region',
      pivotFilters: [contains(PivotAxis.Row, 'Product')],
      valueField: 'Amt',
    });
    expect(result).toEqual({ ok: false, reason: 'invalid-field' });
    expect(wb.getPivotDefinition(0, 0)).toBeNull();
    wb.dispose();
  });
});
