import { describe, expect, it } from 'vitest';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { PivotAxis, PivotFilterKind, type PivotFilterSpec } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  pivotFilterCategory,
  pivotFilterOperand,
} from '../../../src/interact/pivot-filter-kinds.js';

// 2026-05-01, 2026-05-19 and 2026-05-31 in the 1900 date system.
const MAY_1 = 46143;
const MAY_19 = 46161;
const MAY_31 = 46173;

const book = async (): Promise<WorkbookHandle> => {
  const wb = await WorkbookHandle.createDefault();
  const rows: (string | number)[][] = [
    ['Region', 'Date', 'Sales'],
    ['East', MAY_1, 12],
    ['West', MAY_19, 5],
    ['North', MAY_31, 8],
  ];
  for (const [r, row] of rows.entries()) {
    for (const [c, value] of row.entries()) {
      const addr = { sheet: 0, row: r, col: c };
      if (typeof value === 'number') wb.setNumber(addr, value);
      else wb.setText(addr, value);
    }
  }
  const created = createPivotTableFromRange(wb, {
    source: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 2 },
    destination: { sheet: 0, row: 6, col: 0 },
    rowField: 'Region',
    columnField: 'Date',
    valueField: 'Sales',
  });
  if (!created.ok) throw new Error(`pivot not created: ${created.reason}`);
  return wb;
};

/** A filter of `kind` with a representative operand, on the field its menu belongs to. */
const specFor = (kind: PivotFilterKind): PivotFilterSpec => {
  const date = pivotFilterCategory(kind) === 'date';
  const base = date
    ? { axis: PivotAxis.Col, fieldName: 'Date', kind }
    : { axis: PivotAxis.Row, fieldName: 'Region', kind };
  switch (pivotFilterOperand(kind)) {
    case 'text':
      return { ...base, value: 'E' };
    case 'textRange':
      return { ...base, value: 'A', valueHigh: 'M' };
    case 'number':
      return { ...base, value: 7 };
    case 'numberRange':
      return { ...base, value: 1, valueHigh: 10 };
    case 'count':
      return { ...base, value: 2 };
    case 'date':
      return { ...base, value: MAY_19 };
    case 'dateRange':
      return { ...base, value: MAY_1, valueHigh: MAY_31 };
    case 'none':
      return base;
  }
};

describe('PivotTable filters through a save', () => {
  it('keeps every filter kind with its operand', async () => {
    const kinds = Object.values(PivotFilterKind);
    expect(kinds).toHaveLength(68);
    for (const kind of kinds) {
      const wb = await book();
      const spec = specFor(kind);
      expect(wb.addPivotFilter(0, 0, spec), `add kind ${kind}`).toBe(true);
      const reopened = await WorkbookHandle.loadBytes(wb.save());
      expect(reopened.getPivotDefinition(0, 0)?.filters, `kind ${kind}`).toEqual([spec]);
    }
  });
});
