import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { sortActiveColumnAuto, sortRange } from '../../../src/commands/sort.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { executeRibbonFilterDataAction } from '../../../src/interact/ribbon-filter-data.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

// Qty column of a header + 10 row list; rows 3,5,8,11 (1-based) are filtered out.
const QTY = [12, 5, 18, 8, 22, 25, 3, 30, 40, 9];
const SCRAMBLED = [30, 5, 25, 8, 12, 40, 3, 22, 18, 9];

const setup = async (qty: number[]) => {
  const store = createSpreadsheetStore();
  const wb = await WorkbookHandle.createDefault({ preferStub: true });
  const cells = new Map(store.getState().data.cells);
  qty.forEach((value, i) => {
    wb.setNumber({ sheet: 0, row: i + 1, col: 2 }, value);
    cells.set(addrKey({ sheet: 0, row: i + 1, col: 2 }), {
      value: { kind: 'number', value },
      formula: null,
    });
  });
  const filter = { sheet: 0, r0: 0, c0: 0, r1: 10, c1: 4 };
  store.setState((s) => ({
    ...s,
    data: { ...s.data, cells },
    layout: { ...s.layout, hiddenRows: new Set([2, 4, 7, 10]) },
    ui: { ...s.ui, filterRanges: [filter] },
  }));
  return { store, wb, filter };
};

const column = (wb: WorkbookHandle) =>
  Array.from({ length: 10 }, (_, i) => {
    const v = wb.getValue({ sheet: 0, row: i + 1, col: 2 });
    return v.kind === 'number' ? v.value : null;
  });

describe('sort inside a filtered list', () => {
  it('reorders the visible rows only and leaves hidden rows in place', async () => {
    const { store, wb } = await setup(SCRAMBLED);
    const range = { sheet: 0, r0: 1, c0: 0, r1: 10, c1: 4 };
    const ok = sortRange(store.getState(), store, wb, range, {
      byCol: 2,
      direction: 'asc',
      hasHeader: false,
    });
    expect(ok).toBe(true);
    expect(column(wb)).toEqual(QTY);
    expect([...store.getState().layout.hiddenRows].sort((a, b) => a - b)).toEqual([2, 4, 7, 10]);
  });
});

describe('Data > Clear on a filtered list', () => {
  it('shows every row and keeps the AutoFilter range', async () => {
    const { store, wb, filter } = await setup(QTY);
    store.setState((s) => ({
      ...s,
      ui: { ...s.ui, filterCriteria: [{ range: filter, byCol: 2, hiddenValues: ['5'] }] },
      selection: {
        active: { sheet: 0, row: 1, col: 0 },
        anchor: { sheet: 0, row: 1, col: 0 },
        range: { sheet: 0, r0: 1, c0: 0, r1: 10, c1: 4 },
      },
    }));
    executeRibbonFilterDataAction({
      store,
      history: new History(),
      action: 'clear',
      getWb: () => wb,
    } as never);
    const s = store.getState();
    expect(s.layout.hiddenRows.size).toBe(0);
    expect(s.ui.filterRanges).toEqual([filter]);
    expect(s.ui.filterCriteria).toEqual([]);
  });
});

describe('Sort A to Z from one cell of a filtered list', () => {
  it('sorts the whole filter range, so a blank row inside it sorts last among the visible rows', async () => {
    const store = createSpreadsheetStore();
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    const qty: (number | null)[] = [12, 5, 30, 8, 25, null, 40, 3, 18, 22, 9];
    const cells = new Map(store.getState().data.cells);
    qty.forEach((value, i) => {
      if (value === null) return;
      wb.setNumber({ sheet: 0, row: i + 1, col: 2 }, value);
      cells.set(addrKey({ sheet: 0, row: i + 1, col: 2 }), {
        value: { kind: 'number', value },
        formula: null,
      });
    });
    wb.setText({ sheet: 0, row: 0, col: 2 }, 'Qty');
    cells.set(addrKey({ sheet: 0, row: 0, col: 2 }), {
      value: { kind: 'text', value: 'Qty' },
      formula: null,
    });
    const filter = { sheet: 0, r0: 0, c0: 0, r1: 11, c1: 4 };
    const active = { sheet: 0, row: 4, col: 2 };
    store.setState((s) => ({
      ...s,
      data: { ...s.data, cells },
      layout: { ...s.layout, hiddenRows: new Set([1, 3, 5, 8, 9, 11]) },
      ui: { ...s.ui, filterRanges: [filter] },
      selection: { active, anchor: active, range: { sheet: 0, r0: 4, c0: 2, r1: 4, c1: 2 } },
    }));
    sortActiveColumnAuto({
      store,
      workbook: wb,
      history: new History(),
      direction: 'asc',
      origin: 'ribbon',
      commandId: 'sortAsc',
    });
    const col = Array.from({ length: 12 }, (_, r) => wb.getValue({ sheet: 0, row: r, col: 2 }));
    const at = (r: number) =>
      col[r]?.kind === 'number' ? (col[r] as { value: number }).value : null;
    expect([2, 4, 6, 7, 10].map(at)).toEqual([5, 8, 22, 40, null]);
  });
});
