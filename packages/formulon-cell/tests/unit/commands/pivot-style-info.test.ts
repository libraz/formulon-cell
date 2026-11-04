import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { updatePivotStyleInfo } from '../../../src/commands/pivot-style-info.js';
import { createPivotTableAtPlacement } from '../../../src/commands/pivot-table.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const setup = async () => {
  const wb = await WorkbookHandle.createDefault();
  for (const [r, row] of [
    ['Region', 'Sales'],
    ['East', 12],
    ['West', 5],
  ].entries()) {
    for (const [c, value] of row.entries()) {
      const addr = { sheet: 0, row: r, col: c };
      if (typeof value === 'number') wb.setNumber(addr, value);
      else wb.setText(addr, value);
    }
  }
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  const history = new History();
  const result = createPivotTableAtPlacement(
    { store, workbook: wb, history },
    { source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 }, rowField: 'Region', valueField: 'Sales' },
    { row: 8, col: 4 },
  );
  if (!result.ok) throw new Error('pivot not created');
  const pivot = wb.getPivotTables()[0];
  if (!pivot) throw new Error('no pivot');
  return {
    wb,
    store,
    history,
    pivot: { sheetIndex: pivot.sheetIndex, pivotIndex: pivot.pivotIndex },
  };
};

describe('updatePivotStyleInfo', () => {
  it('applies a style name and toggles as one undo step', async () => {
    const { wb, store, history, pivot } = await setup();
    const before = wb.getPivotStyleInfo(pivot.sheetIndex, pivot.pivotIndex);
    expect(
      updatePivotStyleInfo({ store, workbook: wb, history }, pivot, {
        name: 'PivotStyleMedium9',
        showRowStripes: true,
      }),
    ).toBe(true);
    expect(wb.getPivotStyleInfo(pivot.sheetIndex, pivot.pivotIndex)).toMatchObject({
      name: 'PivotStyleMedium9',
      showRowStripes: true,
    });
    expect(history.undo()).toBe(true);
    expect(wb.getPivotStyleInfo(pivot.sheetIndex, pivot.pivotIndex)).toEqual(before);
    expect(history.redo()).toBe(true);
    expect(wb.getPivotStyleInfo(pivot.sheetIndex, pivot.pivotIndex)?.name).toBe(
      'PivotStyleMedium9',
    );
    wb.dispose();
  });

  it('leaves the other toggles as they were', async () => {
    const { wb, store, history, pivot } = await setup();
    updatePivotStyleInfo({ store, workbook: wb, history }, pivot, { showColStripes: true });
    expect(wb.getPivotStyleInfo(pivot.sheetIndex, pivot.pivotIndex)).toMatchObject({
      showRowHeaders: true,
      showColHeaders: true,
      showColStripes: true,
    });
    wb.dispose();
  });
});
