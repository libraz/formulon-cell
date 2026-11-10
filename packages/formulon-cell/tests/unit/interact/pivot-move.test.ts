import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { movePivotTable } from '../../../src/interact/pivot-move.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const setup = async () => {
  const wb = await WorkbookHandle.createDefault();
  const rows: (string | number)[][] = [
    ['Region', 'Sales'],
    ['East', 12],
    ['West', 5],
  ];
  for (const [r, row] of rows.entries()) {
    for (const [c, v] of row.entries()) {
      const addr = { sheet: 0, row: r, col: c };
      if (typeof v === 'number') wb.setNumber(addr, v);
      else wb.setText(addr, v);
    }
  }
  const created = createPivotTableFromRange(wb, {
    source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
    destination: { sheet: 0, row: 2, col: 4 },
    name: 'Report',
    rowField: 'Region',
    valueField: 'Sales',
  });
  expect(created.ok).toBe(true);
  const store = createSpreadsheetStore();
  const history = new History();
  return {
    wb,
    store,
    history,
    deps: { store, workbook: wb, history, origin: 'ribbon' as const, commandId: 'pivotMove' },
  };
};

const boundsOf = (wb: WorkbookHandle, sheet: number) =>
  wb.pivotBounds(sheet).map((b) => [b.range.r0, b.range.c0, b.range.r1, b.range.c1]);

describe('movePivotTable', () => {
  it('moves a report on its sheet and selects the whole extent', async () => {
    const { wb, store, history, deps } = await setup();
    const before = boundsOf(wb, 0)[0] as number[];
    const range = movePivotTable(deps, { sheet: 0, pivotIndex: 0 }, { sheet: 0, row: 2, col: 7 });
    expect(range).toMatchObject({ r0: 2, c0: 7 });
    const [r0 = 0, c0 = 0, r1 = 0, c1 = 0] = before;
    expect(boundsOf(wb, 0)).toEqual([[2, 7, 2 + (r1 - r0), 7 + (c1 - c0)]]);
    expect(store.getState().selection.active).toMatchObject({ row: 2, col: 7 });
    history.undo();
    expect(boundsOf(wb, 0)).toEqual([before]);
    wb.dispose();
  });

  it('moves a report to another sheet as one undo step', async () => {
    const { wb, history, deps } = await setup();
    wb.addSheet('Other');
    const before = boundsOf(wb, 0)[0] as number[];
    const range = movePivotTable(deps, { sheet: 0, pivotIndex: 0 }, { sheet: 1, row: 2, col: 7 });
    expect(range).toMatchObject({ sheet: 1, r0: 2, c0: 7 });
    expect(boundsOf(wb, 0)).toEqual([]);
    history.undo();
    expect(boundsOf(wb, 1)).toEqual([]);
    expect(boundsOf(wb, 0)).toEqual([before]);
    wb.dispose();
  });

  it('moves a report onto a new sheet placed after its own, at A3', async () => {
    const { wb, store, deps } = await setup();
    wb.addSheet('Last');
    const range = movePivotTable(deps, { sheet: 0, pivotIndex: 0 }, 'new-sheet');
    expect(range).toMatchObject({ sheet: 1, r0: 2, c0: 0 });
    expect(wb.sheetName(2)).toBe('Last');
    expect(store.getState().data.sheetIndex).toBe(1);
    wb.dispose();
  });
});
