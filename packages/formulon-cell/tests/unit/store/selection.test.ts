import { describe, expect, it } from 'vitest';

import { copy } from '../../../src/commands/clipboard/copy.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachNavigationPolicy } from '../../../src/interact/navigation-policy.js';
import { selectionContainsAddr } from '../../../src/store/selection-geometry.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const workbook = (): WorkbookHandle => ({ sheetCount: 1 }) as WorkbookHandle;

describe('store/selection — mutators', () => {
  it('setActive moves the active cell and collapses the range to it', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 4, col: 3 });
    const s = store.getState();
    expect(s.selection.active).toEqual({ sheet: 0, row: 4, col: 3 });
    expect(s.selection.range).toEqual({ sheet: 0, r0: 4, c0: 3, r1: 4, c1: 3 });
  });

  it('setActive selects the complete merge when the click lands on its body', () => {
    const store = createSpreadsheetStore();
    mutators.mergeRange(store, { sheet: 0, r0: 2, c0: 2, r1: 3, c1: 4 });

    mutators.setActive(store, { sheet: 0, row: 3, col: 4 });

    expect(store.getState().selection).toMatchObject({
      active: { sheet: 0, row: 2, col: 2 },
      anchor: { sheet: 0, row: 2, col: 2 },
      range: { sheet: 0, r0: 2, c0: 2, r1: 3, c1: 4 },
    });
  });

  it('copy uses the full merged selection dimensions after a body click', () => {
    const store = createSpreadsheetStore();
    mutators.setCell(store, { sheet: 0, row: 1, col: 1 }, { kind: 'text', value: 'title' });
    mutators.mergeRange(store, { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 });
    mutators.setActive(store, { sheet: 0, row: 2, col: 2 });

    const result = copy(store.getState());

    expect(result?.range).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 });
    expect(result?.tsv).toBe('title\t\r\n\t');
  });

  it('extendRangeTo grows the range from anchor toward the target', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 1, col: 1 });
    mutators.extendRangeTo(store, { sheet: 0, row: 4, col: 5 });
    const r = store.getState().selection.range;
    expect(r).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 4, c1: 5 });
  });

  it('extendRangeTo handles "shrink back" (target above-left of anchor)', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 5, col: 5 });
    mutators.extendRangeTo(store, { sheet: 0, row: 2, col: 2 });
    const r = store.getState().selection.range;
    expect(r.r0).toBe(2);
    expect(r.c0).toBe(2);
    expect(r.r1).toBe(5);
    expect(r.c1).toBe(5);
  });

  it('setRange overrides the selected range without touching active', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    mutators.setRange(store, { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 3 });
    const s = store.getState();
    expect(s.selection.range).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 3, c1: 3 });
    expect(s.selection.active).toEqual({ sheet: 0, row: 0, col: 0 });
  });

  it('range mutators reject off-sheet and inverted ranges without a navigation policy', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    const before = store.getState().selection;

    mutators.setRange(store, { sheet: 0, r0: 0, c0: 2, r1: 1048576, c1: 2 });
    mutators.setRange(store, { sheet: 0, r0: 4, c0: 0, r1: 4, c1: 16384 });
    mutators.setRange(store, { sheet: 0, r0: -1, c0: 0, r1: 2, c1: 2 });
    mutators.setRange(store, { sheet: 0, r0: 3, c0: 0, r1: 1, c1: 2 });
    mutators.addExtraRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1048576, c1: 0 });
    mutators.setActive(store, { sheet: 0, row: 1048576, col: 0 });
    mutators.extendRangeTo(store, { sheet: 0, row: 2, col: 16384 });

    expect(store.getState().selection).toEqual(before);
  });

  it('selectRow selects the entire row across the sheet width', () => {
    const store = createSpreadsheetStore();
    mutators.selectRow(store, 7);
    const r = store.getState().selection.range;
    expect(r.r0).toBe(7);
    expect(r.r1).toBe(7);
    // Whole-row selection spans many columns.
    expect(r.c1 - r.c0).toBeGreaterThan(50);
  });

  it('selectRows selects an entire row band and keeps the original anchor row', () => {
    const store = createSpreadsheetStore();
    mutators.selectRows(store, 7, 4);
    const s = store.getState().selection;
    expect(s.range).toEqual({ sheet: 0, r0: 4, c0: 0, r1: 7, c1: 16383 });
    expect(s.anchor).toEqual({ sheet: 0, row: 7, col: 0 });
    expect(s.active).toEqual({ sheet: 0, row: 4, col: 0 });
  });

  it('selectCol selects the entire column across the sheet height', () => {
    const store = createSpreadsheetStore();
    mutators.selectCol(store, 2);
    const r = store.getState().selection.range;
    expect(r.c0).toBe(2);
    expect(r.c1).toBe(2);
    expect(r.r1 - r.r0).toBeGreaterThan(50);
  });

  it('selectCols selects an entire column band and keeps the original anchor column', () => {
    const store = createSpreadsheetStore();
    mutators.selectCols(store, 5, 2);
    const s = store.getState().selection;
    expect(s.range).toEqual({ sheet: 0, r0: 0, c0: 2, r1: 1048575, c1: 5 });
    expect(s.anchor).toEqual({ sheet: 0, row: 0, col: 5 });
    expect(s.active).toEqual({ sheet: 0, row: 0, col: 2 });
  });

  it('selectAll selects every cell on the sheet', () => {
    const store = createSpreadsheetStore();
    mutators.selectAll(store);
    const r = store.getState().selection.range;
    expect(r.r0).toBe(0);
    expect(r.c0).toBe(0);
    expect(r.r1).toBeGreaterThan(50);
    expect(r.c1).toBeGreaterThan(50);
  });

  it('addExtraCell demotes the prior primary range into extraRanges and promotes the new cell', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    mutators.addExtraCell(store, { sheet: 0, row: 5, col: 5 });
    const s = store.getState();
    expect(s.selection.active).toEqual({ sheet: 0, row: 5, col: 5 });
    expect(s.selection.extraRanges?.length).toBe(1);
    expect(s.selection.extraRanges?.[0]).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
  });

  it('addExtraCell promotes the complete merge when Ctrl-click lands on its body', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    mutators.mergeRange(store, { sheet: 0, r0: 4, c0: 3, r1: 5, c1: 5 });

    mutators.addExtraCell(store, { sheet: 0, row: 5, col: 5 });

    expect(store.getState().selection.range).toEqual({
      sheet: 0,
      r0: 4,
      c0: 3,
      r1: 5,
      c1: 5,
    });
    expect(store.getState().selection.extraRanges).toEqual([
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
    ]);
  });

  it('addExtraCell is a no-op when called on the current active cell', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    mutators.addExtraCell(store, { sheet: 0, row: 0, col: 0 });
    const s = store.getState();
    expect(s.selection.extraRanges?.length ?? 0).toBe(0);
  });

  it('subtracts a rectangle atomically and promotes the fragment containing the active cell', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 });
    const base = store.getState().selection;

    expect(
      mutators.applySelectionRectangle(
        store,
        base,
        { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
        'subtract',
        { sheet: 0, row: 1, col: 1 },
        { sheet: 0, row: 1, col: 1 },
      ),
    ).toBe(true);

    const { selection } = store.getState();
    expect(selection.range).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });
    expect(selection.active).toEqual({ sheet: 0, row: 0, col: 0 });
    expect(selectionContainsAddr(selection, { sheet: 0, row: 1, col: 1 })).toBe(false);
    expect(selectionContainsAddr(selection, { sheet: 0, row: 2, col: 2 })).toBe(true);
    expect(selection.extraRanges).toHaveLength(3);
  });

  it('projects a removed active cell to the nearest surviving selected address', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 1, col: 1 });
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 });
    const base = store.getState().selection;

    mutators.applySelectionRectangle(
      store,
      base,
      { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
      'subtract',
      { sheet: 0, row: 1, col: 1 },
      { sheet: 0, row: 1, col: 1 },
    );

    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 0, col: 1 });
    expect(store.getState().selection.anchor).toEqual(store.getState().selection.active);
  });

  it('keeps a sole cell selected when a subtract gesture would empty the selection', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 2, col: 3 });
    const before = store.getState();

    expect(
      mutators.applySelectionRectangle(
        store,
        before.selection,
        { sheet: 0, r0: 2, c0: 3, r1: 2, c1: 3 },
        'subtract',
        { sheet: 0, row: 2, col: 3 },
        { sheet: 0, row: 2, col: 3 },
      ),
    ).toBe(false);
    expect(store.getState()).toBe(before);
  });

  it('clears pending format only when a marquee moves the active cell', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 });
    mutators.setPendingFormat(store, {
      addr: { sheet: 0, row: 0, col: 0 },
      format: { bold: true },
    });
    mutators.applySelectionRectangle(
      store,
      store.getState().selection,
      { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
      'subtract',
      { sheet: 0, row: 1, col: 1 },
      { sheet: 0, row: 1, col: 1 },
    );
    expect(store.getState().ui.pendingFormat).toEqual({
      addr: { sheet: 0, row: 0, col: 0 },
      format: { bold: true },
    });

    mutators.applySelectionRectangle(
      store,
      store.getState().selection,
      { sheet: 0, r0: 4, c0: 4, r1: 4, c1: 4 },
      'add',
      { sheet: 0, row: 4, col: 4 },
      { sheet: 0, row: 4, col: 4 },
    );
    expect(store.getState().ui.pendingFormat).toBeNull();
  });

  it('rejects clamped, list-denied, predicate-denied, and partial-merge gestures without mutation', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 1, col: 1 });
    const fixed = attachNavigationPolicy(store, workbook, {
      range: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 3 },
    });
    const base = store.getState().selection;
    const before = store.getState();
    expect(
      mutators.applySelectionRectangle(
        store,
        base,
        { sheet: 0, r0: 3, c0: 3, r1: 4, c1: 4 },
        'add',
        { sheet: 0, row: 3, col: 3 },
        { sheet: 0, row: 4, col: 4 },
      ),
    ).toBe(false);
    expect(store.getState()).toBe(before);
    fixed.setOptions({
      selectable: [{ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 }],
    });
    expect(
      mutators.applySelectionRectangle(
        store,
        store.getState().selection,
        { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 },
        'add',
        { sheet: 0, row: 1, col: 1 },
        { sheet: 0, row: 2, col: 2 },
      ),
    ).toBe(false);
    fixed.setOptions({ selectable: (addr) => addr.row === 0 || addr.col === 0 });
    expect(
      mutators.applySelectionRectangle(
        store,
        store.getState().selection,
        { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
        'add',
        { sheet: 0, row: 0, col: 0 },
        { sheet: 0, row: 1, col: 1 },
      ),
    ).toBe(false);
    fixed.setOptions(undefined);
    mutators.mergeRange(store, { sheet: 0, r0: 4, c0: 4, r1: 5, c1: 5 });
    const beforePartial = store.getState();
    expect(
      mutators.applySelectionRectangle(
        store,
        beforePartial.selection,
        { sheet: 0, r0: 4, c0: 4, r1: 4, c1: 4 },
        'add',
        { sheet: 0, row: 4, col: 4 },
        { sheet: 0, row: 4, col: 4 },
      ),
    ).toBe(false);
    expect(store.getState()).toBe(beforePartial);
    fixed.dispose();
  });

  it('normalizes selected active and anchor addresses to a merge anchor', () => {
    const store = createSpreadsheetStore();
    mutators.mergeRange(store, { sheet: 0, r0: 2, c0: 2, r1: 3, c1: 3 });
    store.setState((state) => ({
      ...state,
      selection: {
        active: { sheet: 0, row: 3, col: 3 },
        anchor: { sheet: 0, row: 3, col: 3 },
        range: { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 4 },
        extraRanges: [],
      },
    }));
    const base = store.getState().selection;

    mutators.applySelectionRectangle(
      store,
      base,
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      'subtract',
      { sheet: 0, row: 0, col: 0 },
      { sheet: 0, row: 0, col: 0 },
    );

    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 2, col: 2 });
    expect(store.getState().selection.anchor).toEqual({ sheet: 0, row: 2, col: 2 });
  });
});
