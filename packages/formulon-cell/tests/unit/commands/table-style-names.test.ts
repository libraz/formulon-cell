import { describe, expect, it } from 'vitest';
import {
  createPivotTableStyleFromActivePivot,
  createTableStyleFromActiveTable,
  currentPivotStyleId,
  currentTableStyleId,
  customStyleNameError,
  defaultTableOverlay,
  defaultTableStyleId,
  formatAsTable,
  sameStyleName,
  tableStyleOptionsById,
} from '../../../src/commands/format-as-table.js';
import { History } from '../../../src/commands/history.js';
import {
  applyTableStyle,
  createCustomStyle,
  modifyCustomStyle,
} from '../../../src/commands/table-style-actions.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const elements = { headerRow: { fill: '#336699' } };
const ctxFor = (store: ReturnType<typeof createSpreadsheetStore>, history: History | null) => ({
  store,
  history,
});

describe('custom style names', () => {
  it('compares style names case-insensitively', () => {
    expect(sameStyleName('Brand', 'bRAND')).toBe(true);
    expect(sameStyleName('Brand', 'Brand 2')).toBe(false);
  });

  it('rejects a clash with a custom style of either kind or a built-in style, in any case', () => {
    const store = createSpreadsheetStore();
    createTableStyleFromActiveTable(store, null, 'Brand', { elements });
    createPivotTableStyleFromActivePivot(store, null, 'Ledger', { elements });
    const state = store.getState();
    expect(customStyleNameError(state, 'brand')).toBe('exists');
    expect(customStyleNameError(state, 'LEDGER')).toBe('exists');
    expect(customStyleNameError(state, 'tablestylemedium2')).toBe('exists');
    expect(customStyleNameError(state, 'PivotStyleLight16')).toBe('exists');
    expect(customStyleNameError(state, 'Fresh')).toBeNull();
    expect(customStyleNameError(state, 'BRAND', { exceptLabel: 'Brand' })).toBeNull();
  });

  it('refuses a new style whose name is taken instead of overwriting the old one', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const value = { name: 'Brand', elements, stripeSizes: {}, setAsDefault: false };
    expect(createCustomStyle(ctxFor(store, history), null, 'table', value)).toBe(true);
    const other = { ...value, name: 'brand', elements: { headerRow: { fill: '#ff0000' } } };
    expect(createCustomStyle(ctxFor(store, history), null, 'pivot', other)).toBe(false);
    expect(createCustomStyle(ctxFor(store, history), null, 'table', other)).toBe(false);
    expect(store.getState().tables.customTableStyles).toHaveLength(1);
    expect(store.getState().tables.customTableStyles?.[0]?.elements).toEqual(elements);
    expect(store.getState().tables.customPivotTableStyles ?? []).toEqual([]);
  });

  it('refuses Modify onto another style name and lets a style keep its own name', () => {
    const store = createSpreadsheetStore();
    createTableStyleFromActiveTable(store, null, 'A', { elements });
    createPivotTableStyleFromActivePivot(store, null, 'B', { elements });
    const next = { label: 'b', elements: {} };
    expect(modifyCustomStyle(ctxFor(store, null), 'A', next)).toBe(false);
    expect(store.getState().tables.customTableStyles?.map((s) => s.label)).toEqual(['A']);
    expect(modifyCustomStyle(ctxFor(store, null), 'A', { label: 'a', elements: {} })).toBe(true);
    expect(store.getState().tables.customTableStyles?.map((s) => s.label)).toEqual(['a']);
  });

  it('finds the workbook default table style by label in any case', () => {
    const store = createSpreadsheetStore();
    createTableStyleFromActiveTable(store, null, 'Brand', { elements });
    const workbook = { getTableStyleDefaults: () => ({ tableStyle: 'BRAND', pivotStyle: '' }) };
    expect(defaultTableStyleId(store.getState(), workbook)).toBe('custom-table:Brand');
  });
});

describe('current gallery style ids', () => {
  it('maps a style name to the custom style id in any case, else leaves the name', () => {
    const store = createSpreadsheetStore();
    createTableStyleFromActiveTable(store, null, 'Brand', { elements });
    createPivotTableStyleFromActivePivot(store, null, 'Ledger', { elements });
    const state = store.getState();
    expect(currentTableStyleId(state, 'brand')).toBe('custom-table:Brand');
    expect(currentTableStyleId(state, 'TableStyleLight1')).toBe('TableStyleLight1');
    expect(currentPivotStyleId(state, 'LEDGER')).toBe('custom-pivot-table:Ledger');
    expect(currentPivotStyleId(state, '')).toBe('');
  });
});

describe('table style ids', () => {
  it('has no options for an id that is neither custom, built-in nor a family', () => {
    const store = createSpreadsheetStore();
    expect(tableStyleOptionsById(store.getState(), 'custom-table:Missing')).toBeNull();
    expect(tableStyleOptionsById(store.getState(), 'Nonsense')).toBeNull();
    expect(tableStyleOptionsById(store.getState(), 'dark', '#ff0000')).toMatchObject({
      style: 'dark',
    });
  });

  it('refuses to restyle a table with an unknown style id', () => {
    const store = createSpreadsheetStore();
    const table = formatAsTable(store, { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 });
    if (!table) throw new Error('no table');
    const ctx = { ...ctxFor(store, null), origin: 'ribbon' as const, commandId: 'test' };
    expect(applyTableStyle(ctx, table, 'custom-table:Missing')).toBe(false);
    expect(store.getState().tables.tables[0]?.styleName).toBe(
      defaultTableOverlay('x', table.range).styleName,
    );
  });
});
