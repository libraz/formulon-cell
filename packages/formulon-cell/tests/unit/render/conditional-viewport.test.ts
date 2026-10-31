import { describe, expect, it } from 'vitest';
import { createConditionalCache, evaluateConditional } from '../../../src/render/conditional.js';
import type { ConditionalRule, State } from '../../../src/store/store.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import { seedCell, seedNumber, withConditionalRules } from './conditional-fixtures.js';

const range = { sheet: 0, r0: 2, c0: 3, r1: 4, c1: 5 };
const viewport = { r0: 3, c0: 2, r1: 4, c1: 3 };
const inViewKeys = ['0:3:3', '0:4:3'];
const textRule: ConditionalRule = {
  kind: 'text-contains',
  range,
  text: 'match',
  apply: { fill: '#text' },
};

const rules: ConditionalRule[] = [
  textRule,
  { kind: 'date-occurring', range, period: 'today', apply: { fill: '#date' } },
  { kind: 'blanks', range, apply: { fill: '#blank' } },
  { kind: 'non-blanks', range, apply: { color: '#nonblank' } },
  { kind: 'errors', range, apply: { bold: true } },
  { kind: 'no-errors', range, apply: { italic: true } },
];

function fixture(rule: ConditionalRule): State {
  const today = Math.floor(Date.now() / 86_400_000) + 25569;
  let state = createSpreadsheetStore().getState();
  state = seedCell(state, 2, 3, { kind: 'text', value: 'match outside viewport' });
  state = seedNumber(state, 3, 3, today);
  state = seedCell(state, 3, 4, { kind: 'error', code: 6, text: '#N/A' });
  state = seedCell(state, 4, 3, { kind: 'text', value: 'match' });
  state = withConditionalRules(state, [rule]);
  return state;
}

function restrictToViewport(
  overlay: Map<string, unknown>,
  keys = inViewKeys,
): Map<string, unknown> {
  return new Map([...overlay].filter(([key]) => keys.includes(key)));
}

function withReadCount(state: State): { state: State; readCount: () => number } {
  const cells = new Map(state.data.cells);
  let reads = 0;
  const read = cells.get.bind(cells);
  cells.get = (key: string) => {
    reads += 1;
    return read(key);
  };
  return {
    state: { ...state, data: { ...state.data, cells } },
    readCount: () => reads,
  };
}

describe('conditional formatting viewport bounds', () => {
  it.each(rules)('matches full evaluation within the viewport for %s', (rule) => {
    const state = fixture(rule);
    const full = evaluateConditional(state);
    const bounded = evaluateConditional(state, null, '', viewport);

    expect(bounded).toEqual(restrictToViewport(full));
    expect([...bounded.keys()].every((key) => inViewKeys.includes(key))).toBe(true);
  });

  it.each(rules)('does not evaluate cells for a disjoint viewport: %s', (rule) => {
    const tracked = withReadCount(fixture(rule));

    const overlay = evaluateConditional(tracked.state, null, '', {
      r0: 50,
      c0: 50,
      r1: 51,
      c1: 51,
    });

    expect(overlay.size).toBe(0);
    expect(tracked.readCount()).toBe(0);
  });

  it.each(rules)('checks only the 2x3 visible intersection for %s', (rule) => {
    const base = fixture(rule);
    const tracked = withReadCount(base);
    const state = {
      ...tracked.state,
      conditional: {
        ...tracked.state.conditional,
        rules: [{ ...rule, range: { ...range, r1: 101, c1: 102 } }],
      },
    };
    const visibleIntersection = { r0: 3, c0: 3, r1: 4, c1: 5 };

    evaluateConditional(state, null, '', visibleIntersection);

    expect(tracked.readCount()).toBe(6);
  });

  it.each(rules)('checks only explicit visible indices across a large gap for %s', (rule) => {
    const base = fixture(rule);
    const tracked = withReadCount(base);
    const state = {
      ...tracked.state,
      conditional: {
        ...tracked.state.conditional,
        rules: [{ ...rule, range: { sheet: 0, r0: 0, c0: 0, r1: 900_000, c1: 16_000 } }],
      },
    };

    evaluateConditional(state, null, '', {
      r0: 900_000,
      c0: 16_000,
      r1: 900_001,
      c1: 16_001,
      rowIndices: [3, 900_000],
      colIndices: [3, 16_000],
    });

    expect(tracked.readCount()).toBe(4);
  });

  it.each(rules)('treats an empty visible axis as no visible cells for %s', (rule) => {
    const tracked = withReadCount(fixture(rule));

    const overlay = evaluateConditional(tracked.state, null, '', {
      ...viewport,
      rowIndices: [],
      colIndices: [3, 4],
    });

    expect(overlay.size).toBe(0);
    expect(tracked.readCount()).toBe(0);
  });

  it('keeps unbounded output, including implicit blanks, for non-renderer callers', () => {
    const blankRule: ConditionalRule = {
      kind: 'blanks',
      range,
      apply: { fill: '#blank' },
    };
    const state = fixture(blankRule);

    const overlay = evaluateConditional(state);

    expect(overlay.get('0:2:4')?.fill).toBe('#blank');
    expect(overlay.get('0:3:3')?.fill).toBeUndefined();
    expect(overlay.size).toBe(5);
  });

  it('preserves rule priority inside the bounded region', () => {
    let state = fixture(textRule);
    state = withConditionalRules(state, [
      { ...textRule, apply: { fill: '#first' }, stopIfTrue: true },
      { ...textRule, apply: { color: '#blocked' } },
    ]);

    const overlay = evaluateConditional(state, null, '', viewport);

    expect(overlay.get('0:4:3')).toEqual({ fill: '#first' });
  });

  it('uses a new viewport and cell map after the state changes', () => {
    const rule = textRule;
    const state = fixture(rule);
    const cache = createConditionalCache();
    const first = evaluateConditional(state, null, '', viewport, cache);
    const moved = evaluateConditional(state, null, '', { r0: 2, c0: 3, r1: 2, c1: 3 }, cache);
    const changed = seedCell(state, 2, 3, { kind: 'text', value: 'other' });
    const afterCellChange = evaluateConditional(
      changed,
      null,
      '',
      { r0: 2, c0: 3, r1: 2, c1: 3 },
      cache,
    );

    expect(first.has('0:4:3')).toBe(true);
    expect(moved.has('0:2:3')).toBe(true);
    expect(afterCellChange.has('0:2:3')).toBe(false);
  });

  it('snapshots visible indices for cache equality and detects in-place array changes', () => {
    const state = seedCell(fixture(textRule), 2, 4, { kind: 'text', value: 'match' });
    const cache = createConditionalCache();
    const columns = [3];
    const tracked = withReadCount(state);
    const bounds = { r0: 2, c0: 3, r1: 2, c1: 4, rowIndices: [2], colIndices: columns };
    const first = evaluateConditional(tracked.state, null, '', bounds, cache);
    const readsAfterFirst = tracked.readCount();

    const equal = evaluateConditional(
      tracked.state,
      null,
      '',
      {
        ...bounds,
        rowIndices: [2],
        colIndices: [3],
      },
      cache,
    );
    expect(equal).toBe(first);
    expect(tracked.readCount()).toBe(readsAfterFirst);

    columns[0] = 4;
    const moved = evaluateConditional(tracked.state, null, '', bounds, cache);
    expect(moved.has('0:2:4')).toBe(true);
    expect(moved.has('0:2:3')).toBe(false);
    expect(tracked.readCount()).toBe(readsAfterFirst + 1);
  });

  it('distinguishes an empty visible-axis list from an omitted list in the cache', () => {
    const tracked = withReadCount(fixture(textRule));
    const cache = createConditionalCache();
    const bounds = { r0: 2, c0: 3, r1: 2, c1: 4 };
    const unbounded = evaluateConditional(tracked.state, null, '', bounds, cache);
    const readsAfterUnbounded = tracked.readCount();
    const empty = evaluateConditional(
      tracked.state,
      null,
      '',
      { ...bounds, colIndices: [] },
      cache,
    );

    expect(unbounded.has('0:2:3')).toBe(true);
    expect(empty.size).toBe(0);
    expect(tracked.readCount()).toBe(readsAfterUnbounded);
  });

  it('keeps aggregate statistics based on offscreen values', () => {
    const aggregateRule: ConditionalRule = {
      kind: 'color-scale',
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      stops: ['#000000', '#ffffff'],
      thresholds: [{ kind: 'min' }, { kind: 'max' }],
    };
    let state = createSpreadsheetStore().getState();
    state = seedNumber(state, 0, 0, 0);
    state = seedNumber(state, 0, 1, 100);
    state = withConditionalRules(state, [aggregateRule]);

    const full = evaluateConditional(state);
    const bounded = evaluateConditional(state, null, '', {
      r0: 0,
      c0: 0,
      r1: 0,
      c1: 0,
      rowIndices: [0],
      colIndices: [0],
    });

    expect(bounded.get('0:0:0')).toEqual(full.get('0:0:0'));
    expect(full.get('0:0:1')).toBeDefined();
  });

  it.each(rules)('skips rules from a different sheet: %s', (rule) => {
    const tracked = withReadCount(fixture({ ...rule, range: { ...range, sheet: 1 } }));

    const overlay = evaluateConditional(tracked.state, null, '', viewport);

    expect(overlay.size).toBe(0);
    expect(tracked.readCount()).toBe(0);
  });
});
