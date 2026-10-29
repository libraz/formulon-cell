import { beforeEach, describe, expect, it } from 'vitest';
import {
  aggregateSelection,
  aggregateSelectionFromEntries,
  countUniqueRangeCells,
  statusAggregateValue,
  visibleStatusAggregates,
} from '../../../src/commands/aggregate.js';
import type { CellValue } from '../../../src/engine/types.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

class AccessCountingMap<K, V> extends Map<K, V> {
  gets = 0;
  iterations = 0;

  override get(key: K): V | undefined {
    this.gets += 1;
    return super.get(key);
  }

  override [Symbol.iterator](): MapIterator<[K, V]> {
    this.iterations += 1;
    return super[Symbol.iterator]();
  }
}

type AggregateCell = { value: CellValue; formula: string | null };

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

  it('uses one direct lookup for a single-cell selection, including a missing selected cell', () => {
    const cells = new AccessCountingMap<string, AggregateCell>([
      [
        addrKey({ sheet: 1, row: 0, col: 0 }),
        { value: { kind: 'number', value: 999 }, formula: null },
      ],
      [
        addrKey({ sheet: 0, row: 1, col: 0 }),
        { value: { kind: 'number', value: 888 }, formula: null },
      ],
    ]);
    store.setState((s) => ({ ...s, data: { ...s.data, cells } }));
    setRange(store, 0, 0, 0, 0);

    const stats = aggregateSelection(store.getState());

    expect(stats).toEqual({
      cells: 1,
      numericCount: 0,
      nonBlankCount: 0,
      sum: 0,
      avg: 0,
      min: 0,
      max: 0,
    });
    expect(cells.gets).toBe(1);
    expect(cells.iterations).toBe(0);
  });

  it('keeps the public aggregate fresh after an in-place cell-map mutation', () => {
    const cells = new Map<string, AggregateCell>([
      [
        addrKey({ sheet: 0, row: 0, col: 0 }),
        { value: { kind: 'number', value: 1 }, formula: null },
      ],
    ]);
    store.setState((s) => ({ ...s, data: { ...s.data, cells } }));
    setRange(store, 0, 0, 0, 0);
    const state = store.getState();

    expect(aggregateSelection(state).sum).toBe(1);

    cells.set(addrKey({ sheet: 0, row: 0, col: 0 }), {
      value: { kind: 'number', value: 5 },
      formula: null,
    });

    expect(aggregateSelection(state).sum).toBe(5);
  });

  it('reduces supplied entries in insertion order and preserves non-finite results', () => {
    const cells = new Map<string, AggregateCell>([
      [
        addrKey({ sheet: 0, row: 0, col: 0 }),
        { value: { kind: 'number', value: 1e16 }, formula: null },
      ],
      [
        addrKey({ sheet: 0, row: 0, col: 2 }),
        { value: { kind: 'number', value: 1 }, formula: null },
      ],
      [
        addrKey({ sheet: 0, row: 0, col: 1 }),
        { value: { kind: 'number', value: -1e16 }, formula: null },
      ],
      [
        addrKey({ sheet: 0, row: 0, col: 3 }),
        { value: { kind: 'number', value: Number.POSITIVE_INFINITY }, formula: null },
      ],
      [
        addrKey({ sheet: 0, row: 0, col: 4 }),
        { value: { kind: 'number', value: Number.NEGATIVE_INFINITY }, formula: null },
      ],
      [
        addrKey({ sheet: 0, row: 0, col: 5 }),
        { value: { kind: 'number', value: Number.NaN }, formula: null },
      ],
    ]);
    store.setState((s) => ({ ...s, data: { ...s.data, cells } }));
    setRange(store, 0, 0, 0, 2);

    let state = store.getState();
    const roundedInInsertionOrder = aggregateSelectionFromEntries(state, cells);
    expect(roundedInInsertionOrder).toEqual(aggregateSelection(state));
    expect(roundedInInsertionOrder.sum).toBe(0);

    setRange(store, 0, 0, 0, 5);
    state = store.getState();
    const nonFinite = aggregateSelectionFromEntries(state, cells);
    expect(nonFinite).toEqual(aggregateSelection(state));
    expect(nonFinite.sum).toBeNaN();
    expect(nonFinite.avg).toBeNaN();
    expect(nonFinite.min).toBe(Number.NEGATIVE_INFINITY);
    expect(nonFinite.max).toBe(Number.POSITIVE_INFINITY);
  });

  it.each([
    {
      name: 'number',
      value: { kind: 'number', value: 12 } as CellValue,
      numericCount: 1,
      nonBlankCount: 1,
      sum: 12,
      avg: 12,
      min: 12,
      max: 12,
    },
    {
      name: 'text',
      value: { kind: 'text', value: 'hello' } as CellValue,
      numericCount: 0,
      nonBlankCount: 1,
      sum: 0,
      avg: 0,
      min: 0,
      max: 0,
    },
    {
      name: 'error',
      value: { kind: 'error', code: 1, text: '#DIV/0!' } as CellValue,
      numericCount: 0,
      nonBlankCount: 1,
      sum: 0,
      avg: 0,
      min: 0,
      max: 0,
    },
    {
      name: 'blank',
      value: { kind: 'blank' } as CellValue,
      numericCount: 0,
      nonBlankCount: 0,
      sum: 0,
      avg: 0,
      min: 0,
      max: 0,
    },
    {
      name: 'boolean',
      value: { kind: 'bool', value: true } as CellValue,
      numericCount: 0,
      nonBlankCount: 1,
      sum: 0,
      avg: 0,
      min: 0,
      max: 0,
    },
    {
      name: 'NaN',
      value: { kind: 'number', value: Number.NaN } as CellValue,
      numericCount: 1,
      nonBlankCount: 1,
      sum: Number.NaN,
      avg: Number.NaN,
      min: Number.POSITIVE_INFINITY,
      max: Number.NEGATIVE_INFINITY,
    },
    {
      name: 'positive infinity',
      value: { kind: 'number', value: Number.POSITIVE_INFINITY } as CellValue,
      numericCount: 1,
      nonBlankCount: 1,
      sum: Number.POSITIVE_INFINITY,
      avg: Number.POSITIVE_INFINITY,
      min: Number.POSITIVE_INFINITY,
      max: Number.POSITIVE_INFINITY,
    },
    {
      name: 'negative infinity',
      value: { kind: 'number', value: Number.NEGATIVE_INFINITY } as CellValue,
      numericCount: 1,
      nonBlankCount: 1,
      sum: Number.NEGATIVE_INFINITY,
      avg: Number.NEGATIVE_INFINITY,
      min: Number.NEGATIVE_INFINITY,
      max: Number.NEGATIVE_INFINITY,
    },
  ])(
    'preserves aggregate results for a single-cell $name cell',
    ({ value, numericCount, nonBlankCount, sum, avg, min, max }) => {
      const cells = new AccessCountingMap<string, AggregateCell>([
        [addrKey({ sheet: 0, row: 0, col: 0 }), { value, formula: null }],
      ]);
      store.setState((s) => ({ ...s, data: { ...s.data, cells } }));
      setRange(store, 0, 0, 0, 0);

      const stats = aggregateSelection(store.getState());

      expect(stats.numericCount).toBe(numericCount);
      expect(stats.nonBlankCount).toBe(nonBlankCount);
      expect(stats.sum).toBe(sum);
      expect(stats.avg).toBe(avg);
      expect(stats.min).toBe(min);
      expect(stats.max).toBe(max);
      expect(cells.gets).toBe(1);
      expect(cells.iterations).toBe(0);
    },
  );

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
