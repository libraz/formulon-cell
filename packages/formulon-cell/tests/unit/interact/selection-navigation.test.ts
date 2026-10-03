import { describe, expect, it } from 'vitest';
import { buildSelectionInputBatch } from '../../../src/interact/selection-input.js';
import {
  advanceAfterCommit,
  nextWithinSelection,
} from '../../../src/interact/selection-navigation.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const range = (r0: number, c0: number, r1: number, c1: number) => ({
  sheet: 0,
  r0,
  c0,
  r1,
  c1,
});

const addr = (row: number, col: number) => ({ sheet: 0, row, col });

describe('advanceAfterCommit', () => {
  it('traverses the selected rectangle only on Mac and otherwise steps over merges', () => {
    const store = createSpreadsheetStore();
    mutators.mergeRange(store, range(1, 0, 2, 0));
    mutators.setActive(store, addr(0, 0));
    mutators.setRange(store, range(0, 0, 0, 1));

    advanceAfterCommit(store, 'right', true);
    expect(store.getState().selection.active).toEqual(addr(0, 1));
    expect(store.getState().selection.range).toEqual(range(0, 0, 0, 1));

    mutators.setActive(store, addr(1, 0));
    advanceAfterCommit(store, 'down', false);
    expect(store.getState().selection.active).toEqual(addr(3, 0));
  });
});

describe('nextWithinSelection', () => {
  it('uses column-major Return order and wraps in both directions', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, addr(0, 0));
    mutators.setRange(store, range(0, 0, 1, 1));

    expect(nextWithinSelection(store.getState(), 'down')).toEqual(addr(1, 0));
    mutators.setActivePreservingSelection(store, addr(1, 0));
    expect(nextWithinSelection(store.getState(), 'down')).toEqual(addr(0, 1));
    mutators.setActivePreservingSelection(store, addr(0, 1));
    expect(nextWithinSelection(store.getState(), 'up')).toEqual(addr(1, 0));
  });

  it('uses row-major Tab order and wraps in both directions', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, addr(0, 0));
    mutators.setRange(store, range(0, 0, 1, 1));

    expect(nextWithinSelection(store.getState(), 'right')).toEqual(addr(0, 1));
    mutators.setActivePreservingSelection(store, addr(0, 1));
    expect(nextWithinSelection(store.getState(), 'right')).toEqual(addr(1, 0));
    mutators.setActivePreservingSelection(store, addr(1, 0));
    expect(nextWithinSelection(store.getState(), 'left')).toEqual(addr(0, 1));
  });

  it('skips merged interiors in traversal order and returns the anchor once', () => {
    const store = createSpreadsheetStore();
    mutators.mergeRange(store, range(0, 0, 1, 1));
    mutators.setActive(store, addr(0, 0));
    mutators.setRange(store, range(0, 0, 2, 2));

    expect(nextWithinSelection(store.getState(), 'down')).toEqual(addr(2, 0));
    expect(nextWithinSelection(store.getState(), 'right')).toEqual(addr(0, 2));

    store.setState((state) => ({
      ...state,
      selection: { ...state.selection, active: addr(1, 1) },
    }));
    expect(nextWithinSelection(store.getState(), 'down')).toEqual(addr(2, 0));
  });

  it('keeps unmerged cells ahead of partial-width and partial-height merges', () => {
    const store = createSpreadsheetStore();
    mutators.mergeRange(store, range(1, 0, 2, 1));
    mutators.setActive(store, addr(1, 0));
    mutators.setRange(store, range(0, 0, 2, 2));

    expect(nextWithinSelection(store.getState(), 'down')).toEqual(addr(0, 1));
    expect(nextWithinSelection(store.getState(), 'right')).toEqual(addr(1, 2));
  });

  it('returns the merge anchor at the correct point in reverse traversal', () => {
    const store = createSpreadsheetStore();
    mutators.mergeRange(store, range(0, 0, 1, 0));
    mutators.setActive(store, addr(0, 1));
    mutators.setRange(store, range(0, 0, 1, 1));

    expect(nextWithinSelection(store.getState(), 'up')).toEqual(addr(0, 0));

    mutators.mergeRange(store, range(0, 0, 0, 1));
    mutators.setActive(store, addr(0, 2));
    mutators.setRange(store, range(0, 0, 0, 2));
    expect(nextWithinSelection(store.getState(), 'left')).toEqual(addr(0, 0));

    mutators.mergeRange(store, range(0, 1, 0, 2));
    mutators.setActive(store, addr(1, 0));
    mutators.setRange(store, range(0, 0, 1, 2));
    expect(nextWithinSelection(store.getState(), 'left')).toEqual(addr(0, 1));
  });

  it('returns null when the selected rectangle contains only one merge stop', () => {
    const store = createSpreadsheetStore();
    mutators.mergeRange(store, range(4, 3, 5, 4));
    mutators.setActive(store, addr(4, 3));
    mutators.setRange(store, range(4, 3, 5, 4));

    expect(nextWithinSelection(store.getState(), 'down')).toBeNull();
    expect(nextWithinSelection(store.getState(), 'right')).toBeNull();
  });

  it('moves directly within a full-sheet selection without materializing its area', () => {
    const store = createSpreadsheetStore();
    const fullSheet = range(0, 0, 1_048_575, 16_383);
    mutators.setActive(store, addr(700_000, 10_000));
    mutators.setRange(store, fullSheet);

    expect(nextWithinSelection(store.getState(), 'down')).toEqual(addr(700_001, 10_000));
    expect(nextWithinSelection(store.getState(), 'right')).toEqual(addr(700_000, 10_001));
  });
});

describe('buildSelectionInputBatch', () => {
  it('deduplicates overlapping ranges, skips merge bodies, and shifts relative formulas', () => {
    const store = createSpreadsheetStore();
    const primary = range(0, 0, 1, 1);
    const overlapping = range(1, 1, 2, 2);
    mutators.setActive(store, addr(0, 0));
    mutators.setRange(store, primary);
    mutators.addExtraRange(store, overlapping, addr(1, 1));
    mutators.mergeRange(store, range(0, 1, 1, 1));
    mutators.setCellFormat(store, addr(0, 1), { numFmt: { kind: 'text' } });

    const batch = buildSelectionInputBatch(store.getState(), '=A1', addr(0, 0));

    expect(batch?.operation).toBe('formulaEdit');
    expect(batch?.changes).toHaveLength(6);
    expect(batch?.changes).toContainEqual({ addr: addr(0, 1), input: '=A1' });
    expect(batch?.changes).toContainEqual({ addr: addr(1, 0), input: '=A2' });
    expect(batch?.changes).toContainEqual({ addr: addr(1, 2), input: '=C2' });
    expect(batch?.changes).not.toContainEqual({ addr: addr(1, 1), input: '=B2' });
  });

  it('rejects selections above the 100,000-cell materialization bound', () => {
    const store = createSpreadsheetStore();
    store.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        range: range(0, 0, 1_048_575, 16_383),
      },
    }));

    expect(buildSelectionInputBatch(store.getState(), 'x', addr(0, 0))).toBeNull();
  });

  it('recognizes whitespace-prefixed formulas using the normalized input', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, addr(0, 0));
    mutators.setRange(store, range(0, 0, 0, 1));

    const batch = buildSelectionInputBatch(store.getState(), '  =A1 ', addr(0, 0));

    expect(batch?.operation).toBe('formulaEdit');
    expect(batch?.changes).toEqual([
      { addr: addr(0, 0), input: '=A1' },
      { addr: addr(0, 1), input: '=B1' },
    ]);
  });
});
