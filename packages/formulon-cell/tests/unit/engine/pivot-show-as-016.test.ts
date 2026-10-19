import { describe, expect, it } from 'vitest';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { PivotAggregation, PivotShowValuesAs } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

/** Value cell next to each row label of a one-field pivot, keyed by label. */
const valuesByLabel = (wb: WorkbookHandle, labels: readonly string[]): Record<string, unknown> => {
  const cells = Array.from(wb.pivotCells(0));
  const out: Record<string, unknown> = {};
  for (const label of labels) {
    const at = cells.find((c) => c.value.kind === 'text' && c.value.value === label);
    const value = cells.find(
      (c) => at && c.addr.row === at.addr.row && c.addr.col === at.addr.col + 1,
    );
    out[label] = value?.value.kind === 'number' ? value.value.value : value?.value;
  }
  return out;
};

const build = async (
  showValuesAs: PivotShowValuesAs,
  base: { showAsBaseField?: string; showAsBaseItem?: number },
): Promise<WorkbookHandle> => {
  const wb = await WorkbookHandle.createDefault();
  wb.setText({ sheet: 0, row: 0, col: 0 }, 'Region');
  wb.setText({ sheet: 0, row: 0, col: 1 }, 'Sales');
  const rows: [string, number][] = [
    ['East', 12],
    ['West', 8],
    ['North', 20],
  ];
  rows.forEach(([name, n], i) => {
    wb.setText({ sheet: 0, row: i + 1, col: 0 }, name);
    wb.setNumber({ sheet: 0, row: i + 1, col: 1 }, n);
  });
  const created = createPivotTableFromRange(wb, {
    source: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 },
    destination: { sheet: 0, row: 6, col: 0 },
    rowField: 'Region',
    valueField: 'Sales',
    aggregation: PivotAggregation.Sum,
    valueFieldSettings: [{ fieldName: 'Sales', showValuesAs, ...base }],
  });
  expect(created.ok).toBe(true);
  return wb;
};

describe('formulon 0.16 pivot show-values-as kinds', () => {
  it('ranks values largest to smallest along the base field', async () => {
    const wb = await build(PivotShowValuesAs.RankDescending, { showAsBaseField: 'Region' });
    try {
      expect(valuesByLabel(wb, ['East', 'West', 'North'])).toEqual({ East: 2, West: 3, North: 1 });
    } finally {
      wb.dispose();
    }
  });

  it('expresses each value as a fraction of the chosen base item', async () => {
    // Item 1 is West (8) in source order East, West, North.
    const wb = await build(PivotShowValuesAs.PercentOf, {
      showAsBaseField: 'Region',
      showAsBaseItem: 1,
    });
    try {
      expect(valuesByLabel(wb, ['East', 'West', 'North'])).toEqual({
        East: 1.5,
        West: 1,
        North: 2.5,
      });
    } finally {
      wb.dispose();
    }
  });
});
