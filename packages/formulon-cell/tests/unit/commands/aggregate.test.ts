import { beforeEach, describe, expect, it } from 'vitest';
import {
  aggregateSelection,
  countUniqueRangeCells,
  statusAggregateValue,
  visibleStatusAggregates,
} from '../../../src/commands/aggregate.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const seedNumbers = (
  store: SpreadsheetStore,
  entries: Array<{ row: number; col: number; value: number }>,
): void => {
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    for (const e of entries) {
      cells.set(addrKey({ sheet: 0, row: e.row, col: e.col }), {
        value: { kind: 'number', value: e.value },
        formula: null,
      });
    }
    return { ...s, data: { ...s.data, cells } };
  });
};

const setRange = (
  store: SpreadsheetStore,
  r0: number,
  c0: number,
  r1: number,
  c1: number,
): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      ...s.selection,
      range: { sheet: 0, r0, c0, r1, c1 },
    },
  }));
};

describe('aggregateSelection', () => {
  let store: SpreadsheetStore;

  beforeEach(() => {
    store = createSpreadsheetStore();
  });

  it('returns empty stats when the range is degenerate', () => {
    store.setState((s) => ({
      ...s,
      selection: {
        ...s.selection,
        range: { sheet: 0, r0: 5, c0: 5, r1: 4, c1: 4 },
      },
    }));
    const stats = aggregateSelection(store.getState());
    expect(stats.cells).toBe(0);
    expect(stats.numericCount).toBe(0);
    expect(stats.sum).toBe(0);
  });

  it('counts cells from the range area, not the cell map size', () => {
    seedNumbers(store, [
      { row: 0, col: 0, value: 1 },
      { row: 0, col: 1, value: 2 },
    ]);
    setRange(store, 0, 0, 9, 9); // 10x10 = 100
    const stats = aggregateSelection(store.getState());
    expect(stats.cells).toBe(100);
    expect(stats.numericCount).toBe(2);
    expect(stats.sum).toBe(3);
  });

  it('computes sum / avg / min / max from numeric cells inside the range', () => {
    seedNumbers(store, [
      { row: 0, col: 0, value: 10 },
      { row: 1, col: 0, value: 20 },
      { row: 2, col: 0, value: 30 },
    ]);
    setRange(store, 0, 0, 2, 0);
    const stats = aggregateSelection(store.getState());
    expect(stats.numericCount).toBe(3);
    expect(stats.sum).toBe(60);
    expect(stats.avg).toBe(20);
    expect(stats.min).toBe(10);
    expect(stats.max).toBe(30);
  });

  it('ignores cells outside the range', () => {
    seedNumbers(store, [
      { row: 0, col: 0, value: 1 }, // in
      { row: 5, col: 5, value: 100 }, // out
    ]);
    setRange(store, 0, 0, 1, 1);
    const stats = aggregateSelection(store.getState());
    expect(stats.numericCount).toBe(1);
    expect(stats.sum).toBe(1);
  });

  it('ignores non-number kinds', () => {
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      cells.set(addrKey({ sheet: 0, row: 0, col: 0 }), {
        value: { kind: 'text', value: 'hello' },
        formula: null,
      });
      cells.set(addrKey({ sheet: 0, row: 1, col: 0 }), {
        value: { kind: 'number', value: 42 },
        formula: null,
      });
      return { ...s, data: { ...s.data, cells } };
    });
    setRange(store, 0, 0, 5, 5);
    const stats = aggregateSelection(store.getState());
    expect(stats.numericCount).toBe(1);
    expect(stats.sum).toBe(42);
    // Both text + number contribute to the non-blank count.
    expect(stats.nonBlankCount).toBe(2);
  });

  it('ignores cells from other sheets', () => {
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      cells.set(addrKey({ sheet: 0, row: 0, col: 0 }), {
        value: { kind: 'number', value: 10 },
        formula: null,
      });
      cells.set(addrKey({ sheet: 1, row: 0, col: 0 }), {
        value: { kind: 'number', value: 999 },
        formula: null,
      });
      return { ...s, data: { ...s.data, cells } };
    });
    setRange(store, 0, 0, 0, 0);
    const stats = aggregateSelection(store.getState());
    expect(stats.numericCount).toBe(1);
    expect(stats.sum).toBe(10);
  });

  it('returns zero numericCount + zeroed stats when range has no numbers', () => {
    setRange(store, 0, 0, 1, 1);
    const stats = aggregateSelection(store.getState());
    expect(stats.cells).toBe(4);
    expect(stats.numericCount).toBe(0);
    expect(stats.sum).toBe(0);
    expect(stats.avg).toBe(0);
    expect(stats.min).toBe(0);
    expect(stats.max).toBe(0);
  });

  it('multi-range: sums across primary + extraRanges', () => {
    seedNumbers(store, [
      { row: 0, col: 0, value: 1 },
      { row: 5, col: 5, value: 100 },
    ]);
    store.setState((s) => ({
      ...s,
      selection: {
        active: { sheet: 0, row: 0, col: 0 },
        anchor: { sheet: 0, row: 0, col: 0 },
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        extraRanges: [{ sheet: 0, r0: 5, c0: 5, r1: 5, c1: 5 }],
      },
    }));
    const stats = aggregateSelection(store.getState());
    expect(stats.numericCount).toBe(2);
    expect(stats.sum).toBe(101);
  });

  it('multi-range: cells overlapping primary + extra are counted only once', () => {
    seedNumbers(store, [{ row: 1, col: 1, value: 7 }]);
    store.setState((s) => ({
      ...s,
      selection: {
        active: { sheet: 0, row: 0, col: 0 },
        anchor: { sheet: 0, row: 0, col: 0 },
        range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
        extraRanges: [{ sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 }],
      },
    }));
    const stats = aggregateSelection(store.getState());
    expect(stats.cells).toBe(9);
    expect(stats.numericCount).toBe(1);
    expect(stats.sum).toBe(7);
  });

  it('counts unique cells across partially overlapping ranges', () => {
    expect(
      countUniqueRangeCells([
        { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
        { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 },
      ]),
    ).toBe(7);
  });

  it('builds visible status aggregate entries from the configured keys', () => {
    seedNumbers(store, [
      { row: 0, col: 0, value: 10 },
      { row: 1, col: 0, value: 20 },
    ]);
    setRange(store, 0, 0, 1, 0);
    store.setState((s) => ({ ...s, ui: { ...s.ui, statusAggs: ['sum', 'average', 'max'] } }));

    const stats = aggregateSelection(store.getState());
    expect(statusAggregateValue('count', stats)).toBe(2);
    expect(visibleStatusAggregates(store.getState())).toEqual([
      { key: 'sum', value: 30 },
      { key: 'average', value: 15 },
      { key: 'max', value: 20 },
    ]);
  });
});
