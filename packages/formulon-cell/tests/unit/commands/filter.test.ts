import { beforeEach, describe, expect, it } from 'vitest';
import {
  applyAdvancedFilter,
  applyColorFilter,
  applyConditionFilter,
  applyFilter,
  applyFilterColumns,
  applyValueFilter,
  clearFilter,
  copyAdvancedFilterResult,
  distinctFilterItems,
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

  it('stores condition filter criteria and reapplies them after data changes', () => {
    seedText(store, 0, 0, 'Item');
    seedText(store, 1, 0, 'paper');
    seedText(store, 2, 0, 'ink');
    seedText(store, 3, 0, 'pencil');
    const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };

    expect(
      applyConditionFilter(store.getState(), store, range, 0, {
        op: 'contains',
        value: 'p',
      }),
    ).toBe(1);
    expect(store.getState().ui.filterCriteria).toEqual([
      {
        range,
        byCol: 0,
        hiddenValues: [],
        condition: { op: 'contains', value: 'p' },
      },
    ]);

    seedText(store, 3, 0, 'ink');
    expect(reapplyFilters(store.getState(), store)).toBe(2);
    expect(store.getState().layout.hiddenRows.has(2)).toBe(true);
    expect(store.getState().layout.hiddenRows.has(3)).toBe(true);
  });

  it('filters rows by selected cell fill color and reapplies the color criterion', () => {
    seedText(store, 0, 0, 'Status');
    seedText(store, 1, 0, 'Open');
    seedText(store, 2, 0, 'Closed');
    seedText(store, 3, 0, 'Pending');
    mutators.setCellFormat(store, { sheet: 0, row: 1, col: 0 }, { fill: '#ff0000' });
    mutators.setCellFormat(store, { sheet: 0, row: 2, col: 0 }, { fill: '#00ff00' });
    mutators.setCellFormat(store, { sheet: 0, row: 3, col: 0 }, { fill: '#ff0000' });
    const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };

    expect(
      applyColorFilter(store.getState(), store, range, 0, {
        kind: 'cellColor',
        color: '#FF0000',
      }),
    ).toBe(1);
    expect(store.getState().layout.hiddenRows.has(2)).toBe(true);
    expect(store.getState().ui.filterCriteria).toEqual([
      {
        range,
        byCol: 0,
        hiddenValues: [],
        color: { kind: 'cellColor', color: '#ff0000' },
      },
    ]);

    mutators.setCellFormat(store, { sheet: 0, row: 3, col: 0 }, { fill: '#00ff00' });
    expect(reapplyFilters(store.getState(), store)).toBe(2);
    expect(store.getState().layout.hiddenRows.has(3)).toBe(true);
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

  it("filters by the selected cell's displayed value instead of only the raw key", () => {
    seedText(store, 0, 0, 'Value');
    seedNumber(store, 1, 0, 1);
    seedText(store, 2, 0, '1.00');
    seedNumber(store, 3, 0, 2);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 1, col: 0 },
      {
        numFmt: { kind: 'fixed', decimals: 2 },
      },
    );
    const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };
    mutators.setRange(store, range);
    mutators.setActive(store, { sheet: 0, row: 1, col: 0 });

    const hidden = filterBySelectedCellValue(store.getState(), store, range);

    expect(hidden).toBe(1);
    const s = store.getState();
    expect(s.layout.hiddenRows.has(1)).toBe(false);
    expect(s.layout.hiddenRows.has(2)).toBe(false);
    expect(s.layout.hiddenRows.has(3)).toBe(true);
    expect(s.ui.filterCriteria).toEqual([{ range, byCol: 0, hiddenValues: ['2'] }]);
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

  it('distinctFilterItems orders numbers numerically, text next, blanks last', () => {
    seedText(store, 0, 0, 'Header');
    seedNumber(store, 1, 0, 10);
    seedText(store, 2, 0, 'apple');
    seedNumber(store, 3, 0, 2);
    // row 4 left blank
    seedNumber(store, 5, 0, 1);
    const range = { sheet: 0, r0: 0, c0: 0, r1: 5, c1: 0 };

    const items = distinctFilterItems(store.getState(), range, 0);
    expect(items.map((i) => i.key)).toEqual(['1', '2', '10', 'apple', '']);
  });

  it('distinctFilterItems labels numbers through the column number format', () => {
    seedText(store, 0, 0, 'Header');
    seedNumber(store, 1, 0, 1000);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 1, col: 0 },
      {
        numFmt: { kind: 'currency', decimals: 2, symbol: '$' },
      },
    );
    const range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 };

    const items = distinctFilterItems(store.getState(), range, 0);
    // Matching key stays the raw value; label mirrors the formatted grid text.
    expect(items).toEqual([{ key: '1000', label: '$1,000.00' }]);
  });

  it('advanced filter treats bare text as begins-with (case-insensitive)', () => {
    seedText(store, 0, 0, 'Name');
    seedText(store, 1, 0, 'Smith');
    seedText(store, 2, 0, 'Smart');
    seedText(store, 3, 0, 'Jones');
    seedText(store, 5, 0, 'Name');
    seedText(store, 6, 0, 'sm');

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

  it('advanced filter honors =exact, comparison, and blank/non-blank criteria', () => {
    // Exact match: `=Smith` must reject `Smart`.
    seedText(store, 0, 0, 'Name');
    seedText(store, 1, 0, 'Smith');
    seedText(store, 2, 0, 'Smart');
    seedText(store, 5, 0, 'Name');
    seedText(store, 6, 0, '=smith');
    const exactHidden = applyAdvancedFilter(
      store.getState(),
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 },
      { sheet: 0, r0: 5, c0: 0, r1: 6, c1: 0 },
    );
    expect(exactHidden).toBe(1);
    expect(store.getState().layout.hiddenRows.has(2)).toBe(true);

    // Text comparison operator: `>m` keeps only names lexically after "m".
    clearFilter(store.getState(), store);
    store.setState((st) => ({ ...st, data: { ...st.data, cells: new Map() } }));
    seedText(store, 0, 0, 'Name');
    seedText(store, 1, 0, 'apple');
    seedText(store, 2, 0, 'mango');
    seedText(store, 3, 0, 'zebra');
    seedText(store, 5, 0, 'Name');
    seedText(store, 6, 0, '>m');
    const cmpHidden = applyAdvancedFilter(
      store.getState(),
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 },
      { sheet: 0, r0: 5, c0: 0, r1: 6, c1: 0 },
    );
    expect(cmpHidden).toBe(1);
    expect(store.getState().layout.hiddenRows.has(1)).toBe(true);
    expect(store.getState().layout.hiddenRows.has(2)).toBe(false);

    // Blank criterion `=` keeps only empty cells; `<>` keeps only non-empty.
    clearFilter(store.getState(), store);
    store.setState((st) => ({ ...st, data: { ...st.data, cells: new Map() } }));
    seedText(store, 0, 0, 'Name');
    seedText(store, 1, 0, 'x');
    // row 2 blank
    seedText(store, 3, 0, 'y');
    seedText(store, 5, 0, 'Name');
    seedText(store, 6, 0, '=');
    const blankHidden = applyAdvancedFilter(
      store.getState(),
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 },
      { sheet: 0, r0: 5, c0: 0, r1: 6, c1: 0 },
    );
    expect(blankHidden).toBe(2);
    expect(store.getState().layout.hiddenRows.has(2)).toBe(false);
  });

  it('applyFilter replaces prior hides in the range on re-filter (no accumulation)', () => {
    seedNumber(store, 0, 0, 0);
    seedNumber(store, 1, 0, 10);
    seedNumber(store, 2, 0, 20);
    seedNumber(store, 3, 0, 30);
    const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };
    const num = (cell: unknown): number => {
      const v = (cell as { value?: { kind: string; value: number } } | undefined)?.value;
      return v?.kind === 'number' ? v.value : Number.NaN;
    };

    applyFilter(store.getState(), store, range, 0, (cell) => num(cell) >= 20);
    expect(store.getState().layout.hiddenRows.has(1)).toBe(true);

    // Re-filter the same range with the opposite predicate: the old hide on
    // row 1 must be revealed rather than left behind.
    applyFilter(store.getState(), store, range, 0, (cell) => num(cell) <= 20);
    const s = store.getState();
    expect(s.layout.hiddenRows.has(1)).toBe(false);
    expect(s.layout.hiddenRows.has(3)).toBe(true);
  });

  it('applyFilterColumns ANDs every column predicate in a single pass', () => {
    seedNumber(store, 0, 0, 0);
    seedNumber(store, 0, 1, 0);
    seedNumber(store, 1, 0, 10);
    seedNumber(store, 1, 1, 4);
    seedNumber(store, 2, 0, 20);
    seedNumber(store, 2, 1, 5);
    const range = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 };
    const num = (cell: unknown): number => {
      const v = (cell as { value?: { kind: string; value: number } } | undefined)?.value;
      return v?.kind === 'number' ? v.value : Number.NaN;
    };

    const hidden = applyFilterColumns(store.getState(), store, range, [
      { byCol: 0, predicate: (cell) => num(cell) >= 10 },
      { byCol: 1, predicate: (cell) => num(cell) % 2 === 0 },
    ]);
    // Row 1 passes both (10>=10, 4 even); row 2 fails the even test (5).
    expect(hidden).toBe(1);
    expect(store.getState().layout.hiddenRows.has(1)).toBe(false);
    expect(store.getState().layout.hiddenRows.has(2)).toBe(true);
  });

  it('applyFilterColumns refuses huge ranges without materializing hidden rows', () => {
    const range = { sheet: 0, r0: 0, c0: 0, r1: 1048575, c1: 0 };
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, hiddenRows: new Set([10, 200_000]) },
    }));

    const hidden = applyFilterColumns(store.getState(), store, range, [
      { byCol: 0, predicate: () => false },
    ]);

    const state = store.getState();
    expect(hidden).toBe(0);
    expect(state.ui.filterRange).toEqual(range);
    expect(state.layout.hiddenRows.size).toBe(0);
  });

  it('value filters on huge ranges stamp autofilter without storing massive criteria', () => {
    const range = { sheet: 0, r0: 0, c0: 0, r1: 1048575, c1: 0 };
    seedText(store, 1, 0, 'keep');

    const hidden = applyValueFilter(store.getState(), store, range, 0, ['keep']);

    const state = store.getState();
    expect(hidden).toBe(0);
    expect(state.ui.filterRange).toEqual(range);
    expect(state.ui.filterCriteria).toEqual([]);
    expect(state.layout.hiddenRows.size).toBe(0);
  });

  it('distinctFilterItems scans materialized cells for huge ranges', () => {
    const range = { sheet: 0, r0: 0, c0: 0, r1: 1048575, c1: 0 };
    seedText(store, 1, 0, 'b');
    seedText(store, 500_000, 0, 'a');

    const items = distinctFilterItems(store.getState(), range, 0);

    expect(items).toEqual([
      { key: 'a', label: 'a' },
      { key: 'b', label: 'b' },
    ]);
  });

  it('advanced filter refuses huge list ranges without materializing hidden rows', () => {
    const listRange = { sheet: 0, r0: 0, c0: 0, r1: 1048575, c1: 0 };
    const criteriaRange = { sheet: 0, r0: 5, c0: 0, r1: 6, c1: 0 };
    seedText(store, 5, 0, 'Name');
    seedText(store, 6, 0, 'x');

    const hidden = applyAdvancedFilter(store.getState(), store, listRange, criteriaRange);
    const copied = copyAdvancedFilterResult(store.getState(), store, listRange, criteriaRange, {
      sheet: 0,
      row: 10,
      col: 0,
    });

    expect(hidden).toBe(0);
    expect(copied).toBe(0);
    expect(store.getState().ui.filterRange).toEqual(listRange);
    expect(store.getState().layout.hiddenRows.size).toBe(0);
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
