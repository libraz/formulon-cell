import { describe, expect, it } from 'vitest';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

async function emptyReport(): Promise<WorkbookHandle> {
  const wb = await WorkbookHandle.createDefault();
  wb.setText({ sheet: 0, row: 0, col: 0 }, 'Region');
  wb.setText({ sheet: 0, row: 0, col: 1 }, 'Amt');
  wb.setText({ sheet: 0, row: 1, col: 0 }, 'East');
  wb.setNumber({ sheet: 0, row: 1, col: 1 }, 5);
  const created = createPivotTableFromRange(wb, {
    source: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
    destination: { sheet: 0, row: 2, col: 0 },
    rowField: '',
    valueField: '',
  });
  expect(created.ok).toBe(true);
  return wb;
}

describe('empty report extent is one definition', () => {
  it('pivotBounds, pivotOverlapsRange, pivotCellInfo and getPivotTables agree', async () => {
    const wb = await emptyReport();
    const t = wb.getPivotTables()[0];
    expect(t).toMatchObject({ top: 2, left: 0, rows: 18, cols: 3 });
    expect(wb.pivotBounds(0)).toEqual([
      { pivotIndex: 0, range: { sheet: 0, r0: 2, c0: 0, r1: 19, c1: 2 } },
    ]);
    const cell = (row: number, col: number) => ({ sheet: 0, r0: row, c0: col, r1: row, c1: col });
    expect(wb.pivotOverlapsRange(cell(9, 1))).toBe(true);
    expect(wb.pivotOverlapsRange(cell(19, 2))).toBe(true);
    expect(wb.pivotOverlapsRange(cell(20, 0))).toBe(false);
    expect(wb.pivotOverlapsRange(cell(2, 3))).toBe(false);
    expect(wb.pivotCellInfo({ sheet: 0, row: 9, col: 1 })).not.toBeNull();
    expect(wb.pivotCellInfo({ sheet: 0, row: 20, col: 0 })).toBeNull();
    wb.dispose();
  });
});
