import { beforeEach, describe, expect, it } from 'vitest';
import {
  applyAdvancedFilter,
  applyConditionFilter,
  applyFilter,
  applyValueFilter,
  clearFilter,
  copyAdvancedFilterResult,
  filterBySelectedCellValue,
  inferAutoFilterRange,
  reapplyFilters,
  recordFilterChange,
  setAutoFilter,
} from '../../../src/commands/filter.js';
import { History } from '../../../src/commands/history.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const seedNumber = (store: SpreadsheetStore, row: number, col: number, value: number): void => {
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    cells.set(addrKey({ sheet: 0, row, col }), {
      value: { kind: 'number', value },
      formula: null,
    });
    return { ...s, data: { ...s.data, cells } };
  });
};

const seedText = (store: SpreadsheetStore, row: number, col: number, value: string): void => {
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    cells.set(addrKey({ sheet: 0, row, col }), {
      value: { kind: 'text', value },
      formula: null,
    });
    return { ...s, data: { ...s.data, cells } };
  });
};

describe('filter commands', () => {
  let store: SpreadsheetStore;

  beforeEach(() => {
    store = createSpreadsheetStore();
  });

  it('applyFilter hides rows that fail the predicate (header row preserved)', () => {
    // 4-row range: row 0 = header, rows 1..3 = data 10/20/30
    seedNumber(store, 0, 0, 0);
    seedNumber(store, 1, 0, 10);
    seedNumber(store, 2, 0, 20);
    seedNumber(store, 3, 0, 30);
    const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };
    const hidden = applyFilter(store.getState(), store, range, 0, (cell) => {
      const v = cell?.value as { kind: string; value: number } | undefined;
      return v?.kind === 'number' && v.value >= 20;
    });
    expect(hidden).toBe(1);
    const s = store.getState();
    expect(s.layout.hiddenRows.has(1)).toBe(true);
    expect(s.layout.hiddenRows.has(2)).toBe(false);
    expect(s.layout.hiddenRows.has(3)).toBe(false);
    // Header row never hides.
    expect(s.layout.hiddenRows.has(0)).toBe(false);
  });

  it('applyFilter stamps ui.filterRange so headers can paint chevrons', () => {
    seedNumber(store, 1, 0, 1);
    const range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 2 };
    applyFilter(store.getState(), store, range, 0, () => true);
    expect(store.getState().ui.filterRange).toEqual(range);
  });

  it('applies text and number condition filters', () => {
    seedText(store, 0, 0, 'Item');
    seedText(store, 0, 1, 'Qty');
    seedText(store, 1, 0, 'paper');
    seedNumber(store, 1, 1, 24);
    seedText(store, 2, 0, 'ink');
    seedNumber(store, 2, 1, 6);
    seedText(store, 3, 0, 'pencil');
    seedNumber(store, 3, 1, 12);
    const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 };

    expect(
      applyConditionFilter(store.getState(), store, range, 0, {
        op: 'contains',
        value: 'p',
      }),
    ).toBe(1);
    expect(store.getState().layout.hiddenRows.has(2)).toBe(true);

    expect(
      applyConditionFilter(store.getState(), store, range, 1, {
        op: 'greaterThanOrEqual',
        value: '12',
      }),
    ).toBe(1);
    expect(store.getState().layout.hiddenRows.has(1)).toBe(false);
    expect(store.getState().layout.hiddenRows.has(2)).toBe(true);
    expect(store.getState().layout.hiddenRows.has(3)).toBe(false);
  });

  it('setAutoFilter stamps the range without filtering any rows', () => {
    seedNumber(store, 1, 0, 1);
    seedNumber(store, 2, 0, 2);
    const range = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 };
    setAutoFilter(store, range);
    const s = store.getState();
    expect(s.ui.filterRange).toEqual(range);
    expect(s.layout.hiddenRows.size).toBe(0);
  });

  it('infers the surrounding current region for a single-cell Filter toggle', () => {
    seedText(store, 0, 0, 'Item');
    seedText(store, 0, 1, 'Qty');
    seedText(store, 1, 0, 'paper');
    seedNumber(store, 1, 1, 24);
    seedText(store, 2, 0, 'ink');
    seedNumber(store, 2, 1, 6);
    seedText(store, 4, 0, 'outside');
    mutators.setActive(store, { sheet: 0, row: 1, col: 1 });
    mutators.setRange(store, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 });

    expect(inferAutoFilterRange(store.getState())).toEqual({
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 2,
      c1: 1,
    });
  });

  it('keeps an explicit multi-cell Filter selection unchanged', () => {
    const range = { sheet: 0, r0: 2, c0: 2, r1: 4, c1: 3 };
    mutators.setRange(store, range);
    expect(inferAutoFilterRange(store.getState())).toEqual(range);
  });

  it('clearFilter() with no range clears all hidden rows AND filterRange', () => {
    seedNumber(store, 1, 0, 1);
    const range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 };
    applyFilter(store.getState(), store, range, 0, () => false);
    expect(store.getState().layout.hiddenRows.size).toBeGreaterThan(0);
    clearFilter(store.getState(), store);
    const s = store.getState();
    expect(s.layout.hiddenRows.size).toBe(0);
    expect(s.ui.filterRange).toBeNull();
  });

  it('clearFilter(range) clears filterRange only when the range matches', () => {
    seedNumber(store, 1, 0, 1);
    const range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 };
    applyFilter(store.getState(), store, range, 0, () => true);
    // Same range — clears.
    clearFilter(store.getState(), store, range);
    expect(store.getState().ui.filterRange).toBeNull();

    // Re-apply, then clear with a different range — filterRange survives.
    applyFilter(store.getState(), store, range, 0, () => true);
    clearFilter(store.getState(), store, { sheet: 0, r0: 5, c0: 5, r1: 5, c1: 5 });
    expect(store.getState().ui.filterRange).toEqual(range);
  });

  it('records filter range and hidden rows as one undoable action', () => {
    const history = new History();
    seedNumber(store, 0, 0, 0);
    seedNumber(store, 1, 0, 10);
    seedNumber(store, 2, 0, 20);
    const range = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 };

    recordFilterChange(history, store, () =>
      applyFilter(store.getState(), store, range, 0, (cell) => {
        const v = cell?.value as { kind: string; value: number } | undefined;
        return v?.kind === 'number' && v.value >= 20;
      }),
    );

    expect(store.getState().ui.filterRange).toEqual(range);
    expect(store.getState().layout.hiddenRows.has(1)).toBe(true);
    expect(history.canUndo()).toBe(true);

    history.undo();
    expect(store.getState().ui.filterRange).toBeNull();
    expect(store.getState().layout.hiddenRows.size).toBe(0);

    history.redo();
    expect(store.getState().ui.filterRange).toEqual(range);
    expect(store.getState().layout.hiddenRows.has(1)).toBe(true);
  });

  it('reapplies stored value-filter criteria after data changes', () => {
    seedNumber(store, 0, 0, 0);
    seedNumber(store, 1, 0, 10);
    seedNumber(store, 2, 0, 20);
    const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };

    applyValueFilter(store.getState(), store, range, 0, ['10']);
    expect(store.getState().layout.hiddenRows.has(1)).toBe(true);
    expect(store.getState().ui.filterCriteria).toEqual([{ range, byCol: 0, hiddenValues: ['10'] }]);

    seedNumber(store, 3, 0, 10);
    const hidden = reapplyFilters(store.getState(), store);
    expect(hidden).toBe(2);
    expect(store.getState().layout.hiddenRows.has(1)).toBe(true);
    expect(store.getState().layout.hiddenRows.has(3)).toBe(true);
  });

  it("filters by the selected cell's value and stores a reapplyable value criterion", () => {
    seedNumber(store, 0, 0, 0);
    seedNumber(store, 1, 0, 10);
    seedNumber(store, 2, 0, 20);
    seedNumber(store, 3, 0, 10);
    seedNumber(store, 4, 0, 20);
    const range = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 };
    mutators.setRange(store, range);
    mutators.setActive(store, { sheet: 0, row: 2, col: 0 });

    const hidden = filterBySelectedCellValue(store.getState(), store, range);

    expect(hidden).toBe(2);
    const s = store.getState();
    expect(s.layout.hiddenRows.has(1)).toBe(true);
    expect(s.layout.hiddenRows.has(2)).toBe(false);
    expect(s.layout.hiddenRows.has(3)).toBe(true);
    expect(s.layout.hiddenRows.has(4)).toBe(false);
    expect(s.ui.filterCriteria).toEqual([{ range, byCol: 0, hiddenValues: ['10'] }]);
  });

  it('filters by selected cell value using the inferred current region when no filter exists', () => {
    seedText(store, 0, 0, 'Item');
    seedText(store, 0, 1, 'Qty');
    seedText(store, 1, 0, 'paper');
    seedNumber(store, 1, 1, 24);
    seedText(store, 2, 0, 'ink');
    seedNumber(store, 2, 1, 6);
    seedText(store, 3, 0, 'paper');
    seedNumber(store, 3, 1, 2);
    mutators.setActive(store, { sheet: 0, row: 1, col: 0 });
    mutators.setRange(store, { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 0 });

    const hidden = filterBySelectedCellValue(store.getState(), store);

    expect(hidden).toBe(1);
    expect(store.getState().ui.filterRange).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 });
    expect(store.getState().layout.hiddenRows.has(2)).toBe(true);
    expect(store.getState().layout.hiddenRows.has(1)).toBe(false);
    expect(store.getState().layout.hiddenRows.has(3)).toBe(false);
  });

  it('applies advanced filter criteria rows as OR and columns as AND', () => {
    seedText(store, 0, 0, 'Item');
    seedText(store, 0, 1, 'Qty');
    seedText(store, 1, 0, 'paper');
    seedNumber(store, 1, 1, 24);
    seedText(store, 2, 0, 'ink');
    seedNumber(store, 2, 1, 6);
    seedText(store, 3, 0, 'paper');
    seedNumber(store, 3, 1, 2);
    seedText(store, 5, 0, 'Item');
    seedText(store, 5, 1, 'Qty');
    seedText(store, 6, 0, 'paper');
    seedText(store, 6, 1, '>10');
    seedText(store, 7, 0, 'ink');

    const hidden = applyAdvancedFilter(
      store.getState(),
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 },
      { sheet: 0, r0: 5, c0: 0, r1: 7, c1: 1 },
    );

    const s = store.getState();
    expect(hidden).toBe(1);
    expect(s.layout.hiddenRows.has(1)).toBe(false);
    expect(s.layout.hiddenRows.has(2)).toBe(false);
    expect(s.layout.hiddenRows.has(3)).toBe(true);
    expect(s.ui.filterRange).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 });
  });

  it('supports Excel-style wildcards in advanced filter text criteria', () => {
    seedText(store, 0, 0, 'Item');
    seedText(store, 1, 0, 'paper');
    seedText(store, 2, 0, 'pencil');
    seedText(store, 3, 0, 'ink');
    seedText(store, 5, 0, 'Item');
    seedText(store, 6, 0, 'p*');

    const hidden = applyAdvancedFilter(
      store.getState(),
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 },
      { sheet: 0, r0: 5, c0: 0, r1: 6, c1: 0 },
    );

    const s = store.getState();
    expect(hidden).toBe(1);
    expect(s.layout.hiddenRows.has(1)).toBe(false);
    expect(s.layout.hiddenRows.has(2)).toBe(false);
    expect(s.layout.hiddenRows.has(3)).toBe(true);
  });

  it('copies advanced filter results to another location with unique records', () => {
    seedText(store, 0, 0, 'Item');
    seedText(store, 0, 1, 'Qty');
    seedText(store, 1, 0, 'paper');
    seedNumber(store, 1, 1, 24);
    seedText(store, 2, 0, 'paper');
    seedNumber(store, 2, 1, 24);
    seedText(store, 3, 0, 'ink');
    seedNumber(store, 3, 1, 6);
    seedText(store, 5, 0, 'Item');
    seedText(store, 6, 0, 'p*');

    const copied = copyAdvancedFilterResult(
      store.getState(),
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 },
      { sheet: 0, r0: 5, c0: 0, r1: 6, c1: 0 },
      { sheet: 0, row: 8, col: 0 },
      { uniqueOnly: true },
    );

    expect(copied).toBe(2);
    const cells = store.getState().data.cells;
    expect(cells.get(addrKey({ sheet: 0, row: 8, col: 0 }))?.value).toEqual({
      kind: 'text',
      value: 'Item',
    });
    expect(cells.get(addrKey({ sheet: 0, row: 9, col: 0 }))?.value).toEqual({
      kind: 'text',
      value: 'paper',
    });
    expect(cells.get(addrKey({ sheet: 0, row: 9, col: 1 }))?.value).toEqual({
      kind: 'number',
      value: 24,
    });
    expect(cells.get(addrKey({ sheet: 0, row: 10, col: 0 }))).toBeUndefined();
  });
});
