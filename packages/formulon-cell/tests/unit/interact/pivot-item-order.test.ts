import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { movePivotItem, pivotItemSlotsAt } from '../../../src/interact/pivot-item-order.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

// Report at A3:C7: header row 3, items East / North / South on rows 4-6, total row 7.
const setup = async () => {
  const wb = await WorkbookHandle.createDefault();
  const rows: (string | number)[][] = [
    ['Region', 'Amt', 'Qty'],
    ['East', 5, 5],
    ['North', 3, 3],
    ['South', 7, 7],
  ];
  for (const [r, row] of rows.entries()) {
    for (const [c, v] of row.entries()) {
      const addr = { sheet: 0, row: r, col: c + 10 };
      if (typeof v === 'number') wb.setNumber(addr, v);
      else wb.setText(addr, v);
    }
  }
  const created = createPivotTableFromRange(wb, {
    source: { sheet: 0, r0: 0, c0: 10, r1: 3, c1: 12 },
    destination: { sheet: 0, row: 2, col: 0 },
    name: 'Report',
    rowField: 'Region',
    valueField: 'Amt',
  });
  expect(created.ok).toBe(true);
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  const history = new History();
  const labels = (): string[] => {
    const cells = store.getState().data.cells;
    return [3, 4, 5].map((row) => {
      const v = cells.get(addrKey({ sheet: 0, row, col: 0 }))?.value;
      return v?.kind === 'text' ? v.value : '';
    });
  };
  return { wb, store, history, labels, deps: { store, workbook: wb, history } };
};

describe('pivot item reorder', () => {
  it('finds the item rows from a selected label cell', async () => {
    const { wb, store } = await setup();
    expect(pivotItemSlotsAt(store.getState(), wb, 0, 5, 0)).toMatchObject({
      firstRow: 3,
      lastRow: 5,
      col: 0,
      c1: wb.pivotBounds(0)[0]?.range.c1,
    });
    expect(pivotItemSlotsAt(store.getState(), wb, 0, 2, 0)).toBeNull();
    expect(pivotItemSlotsAt(store.getState(), wb, 0, 6, 0)).toBeNull();
    expect(pivotItemSlotsAt(store.getState(), wb, 0, 4, 1)).toBeNull();
    wb.dispose();
  });

  it('drops an item after the row it is released on', async () => {
    const { wb, store, labels, deps } = await setup();
    const slots = pivotItemSlotsAt(store.getState(), wb, 0, 5, 0);
    if (!slots) throw new Error('no slots');
    // South released on the header row lands first.
    expect(movePivotItem(deps, slots, 5, 2)).toBe(true);
    expect(labels()).toEqual(['South', 'East', 'North']);
    // North released on the first item's row lands right after it.
    expect(movePivotItem(deps, slots, 5, 3)).toBe(true);
    expect(labels()).toEqual(['South', 'North', 'East']);
    // Released on the row above its own: it stays where it is.
    expect(movePivotItem(deps, slots, 5, 4)).toBe(false);
    expect(labels()).toEqual(['South', 'North', 'East']);
    wb.dispose();
  });

  it('moves an item down and selects the item rows', async () => {
    const { wb, store, labels, deps } = await setup();
    const slots = pivotItemSlotsAt(store.getState(), wb, 0, 3, 0);
    if (!slots) throw new Error('no slots');
    expect(movePivotItem(deps, slots, 3, 4)).toBe(true);
    expect(labels()).toEqual(['North', 'East', 'South']);
    expect(store.getState().selection.range).toMatchObject({ r0: 3, r1: 5, c0: 0, c1: 0 });
    wb.dispose();
  });

  it('undoes in one step', async () => {
    const { wb, labels, store, history, deps } = await setup();
    const slots = pivotItemSlotsAt(store.getState(), wb, 0, 5, 0);
    if (!slots) throw new Error('no slots');
    movePivotItem(deps, slots, 5, 2);
    history.undo();
    expect(labels()).toEqual(['East', 'North', 'South']);
    history.redo();
    expect(labels()).toEqual(['South', 'East', 'North']);
    wb.dispose();
  });
});
