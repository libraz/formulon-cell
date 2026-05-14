import { describe, expect, it } from 'vitest';

import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

/**
 * Unit: state-shape mutators on the store. The existing suites already cover
 * selection / viewport / format paths; this file pins the surface that hadn't
 * been exercised — UI flags, theme, status aggregates, hover, editor mode,
 * and the multi-range extra-cell flow.
 */
describe('store mutators — UI flags', () => {
  it('setTheme updates ui.theme', () => {
    const store = createSpreadsheetStore();
    mutators.setTheme(store, 'ink');
    expect(store.getState().ui.theme).toBe('ink');
    mutators.setTheme(store, 'paper');
    expect(store.getState().ui.theme).toBe('paper');
  });

  it('toggle-style flags flip the UI slice', () => {
    const store = createSpreadsheetStore();
    const initial = store.getState().ui;
    mutators.setShowGridLines(store, !initial.showGridLines);
    expect(store.getState().ui.showGridLines).toBe(!initial.showGridLines);
    mutators.setShowHeaders(store, !initial.showHeaders);
    expect(store.getState().ui.showHeaders).toBe(!initial.showHeaders);
    mutators.setShowFormulas(store, true);
    expect(store.getState().ui.showFormulas).toBe(true);
    mutators.setR1C1(store, true);
    expect(store.getState().ui.r1c1).toBe(true);
  });

  it('setHover stores and clears the hover address', () => {
    const store = createSpreadsheetStore();
    const addr = { sheet: 0, row: 3, col: 1 };
    mutators.setHover(store, addr);
    expect(store.getState().ui.hover).toEqual(addr);
    mutators.setHover(store, null);
    expect(store.getState().ui.hover).toBeNull();
  });

  it('setEditor swaps the editor mode without touching other UI flags', () => {
    const store = createSpreadsheetStore();
    const before = store.getState().ui;
    mutators.setEditor(store, { kind: 'edit', raw: 'hello', caret: 5 });
    const after = store.getState().ui;
    expect(after.editor.kind).toBe('edit');
    // Other flags unchanged.
    expect(after.theme).toBe(before.theme);
    expect(after.showGridLines).toBe(before.showGridLines);
  });
});

describe('store mutators — status aggregates', () => {
  it('toggleStatusAgg adds/removes a key', () => {
    const store = createSpreadsheetStore();
    expect(store.getState().ui.statusAggs).toContain('sum');
    mutators.toggleStatusAgg(store, 'sum');
    expect(store.getState().ui.statusAggs).not.toContain('sum');
    mutators.toggleStatusAgg(store, 'sum');
    expect(store.getState().ui.statusAggs).toContain('sum');
  });
});

describe('store mutators — multi-range selection', () => {
  it('addExtraCell demotes the current range and promotes the new cell', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 1, col: 1 });
    mutators.extendRangeTo(store, { sheet: 0, row: 2, col: 3 });
    const before = store.getState().selection.range;

    mutators.addExtraCell(store, { sheet: 0, row: 5, col: 5 });
    const after = store.getState().selection;

    expect(after.active).toEqual({ sheet: 0, row: 5, col: 5 });
    expect(after.range).toEqual({ sheet: 0, r0: 5, c0: 5, r1: 5, c1: 5 });
    expect(after.extraRanges).toHaveLength(1);
    expect(after.extraRanges?.[0]).toEqual(before);
  });

  it('addExtraCell is a no-op when the new addr matches the active cell', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 2, col: 2 });
    const before = store.getState().selection;
    mutators.addExtraCell(store, { sheet: 0, row: 2, col: 2 });
    expect(store.getState().selection).toBe(before);
  });

  it('setActive clears extraRanges (single-cell selection wins)', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    mutators.addExtraCell(store, { sheet: 0, row: 4, col: 4 });
    expect(store.getState().selection.extraRanges?.length).toBe(1);

    mutators.setActive(store, { sheet: 0, row: 1, col: 1 });
    expect(store.getState().selection.extraRanges).toEqual([]);
  });
});
