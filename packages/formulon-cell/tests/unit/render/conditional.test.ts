import { afterEach, describe, expect, it, vi } from 'vitest';
import type { CellValue } from '../../../src/engine/types.js';
import {
  _resetConditionalCache,
  evaluateConditional,
  iconSetSlotCount,
  iconSetSlotFor,
  parseFormulaPredicate,
  topBottomThreshold,
} from '../../../src/render/conditional.js';
import type { ConditionalRule, State } from '../../../src/store/store.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const seedCell = (state: State, row: number, col: number, value: CellValue): State => {
  const cells = new Map(state.data.cells);
  cells.set(`0:${row}:${col}`, { value, formula: null });
  return { ...state, data: { ...state.data, cells } };
};

const seedNumber = (state: State, row: number, col: number, value: number): State => {
  const cells = new Map(state.data.cells);
  cells.set(`0:${row}:${col}`, { value: { kind: 'number', value }, formula: null });
  return { ...state, data: { ...state.data, cells } };
};

const dateSerial = (year: number, month: number, day: number): number =>
  Date.UTC(year, month - 1, day) / 86_400_000 + 25569;

const cellValueRule = (range: ConditionalRule['range']): ConditionalRule => ({
  kind: 'cell-value',
  range,
  op: '>',
  a: 5,
  apply: { fill: '#ff0000' },
});

describe('evaluateConditional', () => {
  afterEach(() => {
    _resetConditionalCache();
    vi.useRealTimers();
  });

  it('returns empty overlay when no rules are configured', () => {
    const store = createSpreadsheetStore();
    const r = evaluateConditional(store.getState());
    expect(r.size).toBe(0);
  });

  it('marks cells whose value passes the cell-value predicate', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 10);
    s = seedNumber(s, 0, 1, 3);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [cellValueRule({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 })],
      },
    };
    const overlay = evaluateConditional(s);
    expect(overlay.get('0:0:0')?.fill).toBe('#ff0000');
    // Cells in the rule range that fail the predicate get no fill — the
    // renderer treats an empty overlay as a no-op.
    expect(overlay.get('0:0:1')?.fill).toBeUndefined();
  });

  it('cell-value rules compare text cells case-insensitively', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'Alpha' });
    s = seedCell(s, 0, 1, { kind: 'text', value: 'Beta' });
    s = seedCell(s, 0, 2, { kind: 'text', value: 'Delta' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'cell-value',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 },
            op: 'between',
            a: 'b',
            b: 'dzz',
            apply: { fill: '#text' },
          },
        ],
      },
    };
    const overlay = evaluateConditional(s);
    expect(overlay.get('0:0:0')?.fill).toBeUndefined();
    expect(overlay.get('0:0:1')?.fill).toBe('#text');
    expect(overlay.get('0:0:2')?.fill).toBe('#text');
  });

  it('preserves data-bar gradient versus solid rendering metadata', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 10);
    s = seedNumber(s, 0, 1, 20);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'data-bar',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            color: '#63a95c',
            gradient: true,
            showValue: false,
          },
          {
            kind: 'data-bar',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            color: '#70ad47',
            gradient: false,
          },
        ],
      },
    };
    const overlay = evaluateConditional(s);
    expect(overlay.get('0:0:0')).toMatchObject({
      barColor: '#63a95c',
      barGradient: true,
      showValue: false,
    });
    expect(overlay.get('0:0:1')).toMatchObject({ barColor: '#70ad47', barGradient: false });
  });

  it('data-bar overlays expose a zero axis and signed direction', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, -10);
    s = seedNumber(s, 0, 1, 0);
    s = seedNumber(s, 0, 2, 20);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'data-bar',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 },
            color: '#70ad47',
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.barAxis).toBeCloseTo(1 / 3);
    expect(overlay.get('0:0:0')?.barDirection).toBe('left');
    expect(overlay.get('0:0:0')?.bar).toBeCloseTo(1 / 3);
    expect(overlay.get('0:0:1')?.bar).toBe(0);
    expect(overlay.get('0:0:2')?.barDirection).toBe('right');
    expect(overlay.get('0:0:2')?.bar).toBeCloseTo(2 / 3);
  });

  it('returns the same Map reference when called twice with identical state', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 10);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [cellValueRule({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 })],
      },
    };
    const a = evaluateConditional(s);
    const b = evaluateConditional(s);
    expect(b).toBe(a);
  });

  it('returns the cached result when only an unrelated slice (selection) changed', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 10);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [cellValueRule({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 })],
      },
    };
    const a = evaluateConditional(s);
    // Mutate selection (and thus the top-level state object), but keep cells
    // and rules references untouched. Cache should still hit.
    const sNext: State = {
      ...s,
      selection: {
        ...s.selection,
        active: { sheet: 0, row: 5, col: 5 },
      },
    };
    const b = evaluateConditional(sNext);
    expect(b).toBe(a);
  });

  it('recomputes when the cells map reference changes', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 10);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [cellValueRule({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 })],
      },
    };
    const a = evaluateConditional(s);
    const sNext = seedNumber(s, 0, 0, 1); // value drops below 5 → no fill
    const b = evaluateConditional(sNext);
    expect(b).not.toBe(a);
    expect(a.get('0:0:0')?.fill).toBe('#ff0000');
    expect(b.get('0:0:0')?.fill).toBeUndefined();
  });

  it('recomputes when the rules array reference changes', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 10);
    const rule1 = cellValueRule({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    s = { ...s, conditional: { ...s.conditional, rules: [rule1] } };
    const a = evaluateConditional(s);
    // Same rule shape, new array reference — must invalidate.
    const sNext: State = { ...s, conditional: { ...s.conditional, rules: [{ ...rule1 }] } };
    const b = evaluateConditional(sNext);
    expect(b).not.toBe(a);
  });

  it('recomputes when the active sheet changes', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 10);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [cellValueRule({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 })],
      },
    };
    const a = evaluateConditional(s);
    const sNext: State = { ...s, data: { ...s.data, sheetIndex: 1 } };
    const b = evaluateConditional(sNext);
    expect(b).not.toBe(a);
    // Rule range targets sheet 0, so on sheet 1 nothing applies.
    expect(b.size).toBe(0);
  });

  it('keeps higher-priority rule attributes when lower-priority rules also match', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 10);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'cell-value',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            op: '>',
            a: 0,
            apply: { fill: '#high' },
          },
          {
            kind: 'cell-value',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            op: '>',
            a: 0,
            apply: { fill: '#low', color: '#low-text' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#high');
    expect(overlay.get('0:0:0')?.color).toBe('#low-text');
  });

  it('honors stopIfTrue by skipping lower-priority rules for matched cells', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 10);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'cell-value',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            op: '>',
            a: 0,
            apply: { fill: '#stop' },
            stopIfTrue: true,
          },
          {
            kind: 'cell-value',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            op: '>',
            a: 0,
            apply: { color: '#blocked' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#stop');
    expect(overlay.get('0:0:0')?.color).toBeUndefined();
  });

  it('evaluates date-occurring week periods with Monday week boundaries', () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date(Date.UTC(2026, 6, 8, 12))); // Wednesday, 2026-07-08.
    const store = createSpreadsheetStore();
    const rule: ConditionalRule = {
      kind: 'date-occurring',
      range: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 },
      period: 'this-week',
      apply: { fill: '#00ff00' },
    };
    let s = store.getState();
    s = seedNumber(s, 0, 0, dateSerial(2026, 7, 12)); // Sunday in the current Mon-Sun week.
    s = seedNumber(s, 1, 0, dateSerial(2026, 7, 5)); // Sunday in the previous Mon-Sun week.
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [rule],
      },
    };
    const thisWeek = evaluateConditional(s);
    expect(thisWeek.get('0:0:0')?.fill).toBe('#00ff00');
    expect(thisWeek.get('0:1:0')?.fill).toBeUndefined();

    _resetConditionalCache();
    const lastWeekState: State = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [{ ...rule, period: 'last-week' }],
      },
    };
    const lastWeek = evaluateConditional(lastWeekState);
    expect(lastWeek.get('0:0:0')?.fill).toBeUndefined();
    expect(lastWeek.get('0:1:0')?.fill).toBe('#00ff00');
  });

  it('honors color-scale threshold metadata for number and percentile stops', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 0);
    s = seedNumber(s, 0, 1, 10);
    s = seedNumber(s, 0, 2, 100);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'color-scale',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 },
            stops: ['#000000', '#ffffff'],
            thresholds: [{ kind: 'number', value: 10 }, { kind: 'max' }],
          },
        ],
      },
    };
    const overlay = evaluateConditional(s);
    expect(overlay.get('0:0:0')?.fill).toBe('rgb(0, 0, 0)');
    expect(overlay.get('0:0:1')?.fill).toBe('rgb(0, 0, 0)');
    expect(overlay.get('0:0:2')?.fill).toBe('rgb(255, 255, 255)');

    _resetConditionalCache();
    const s2: State = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'color-scale',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 },
            stops: ['#000000', '#808080', '#ffffff'],
            thresholds: [{ kind: 'min' }, { kind: 'percentile', value: 50 }, { kind: 'max' }],
          },
        ],
      },
    };
    expect(evaluateConditional(s2).get('0:0:1')?.fill).toBe('rgb(128, 128, 128)');
  });

  it('uses the midpoint color when a three-color scale has degenerate thresholds', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 5);
    s = seedNumber(s, 0, 1, 5);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'color-scale',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
            stops: ['#000000', '#808080', '#ffffff'],
            thresholds: [{ kind: 'min' }, { kind: 'percentile', value: 50 }, { kind: 'max' }],
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);
    expect(overlay.get('0:0:0')?.fill).toBe('rgb(128, 128, 128)');
    expect(overlay.get('0:0:1')?.fill).toBe('rgb(128, 128, 128)');
  });

  it('icon-set classifies cells by percentile and forwards reverseOrder', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 0);
    s = seedNumber(s, 0, 1, 50);
    s = seedNumber(s, 0, 2, 100);
    const rule: ConditionalRule = {
      kind: 'icon-set',
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 },
      icons: 'arrows3',
      showValue: false,
    };
    s = { ...s, conditional: { ...s.conditional, rules: [rule] } };
    const overlay = evaluateConditional(s);
    expect(overlay.get('0:0:0')?.iconSlot).toBe(0);
    expect(overlay.get('0:0:0')?.showValue).toBe(false);
    expect(overlay.get('0:0:1')?.iconSlot).toBe(1);
    expect(overlay.get('0:0:2')?.iconSlot).toBe(2);
    // Reverse order — slots 0/1/2 invert to 2/1/0.
    _resetConditionalCache();
    const reversed: ConditionalRule = { ...rule, reverseOrder: true };
    const s2 = { ...s, conditional: { ...s.conditional, rules: [reversed] } };
    const overlay2 = evaluateConditional(s2);
    expect(overlay2.get('0:0:0')?.iconSlot).toBe(2);
    expect(overlay2.get('0:0:2')?.iconSlot).toBe(0);
  });

  it('icon-set honors custom threshold metadata', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 10);
    s = seedNumber(s, 0, 1, 50);
    s = seedNumber(s, 0, 2, 90);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'icon-set',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 },
            icons: 'traffic3',
            thresholds: [
              { kind: 'number', value: 30 },
              { kind: 'number', value: 80 },
            ],
          },
        ],
      },
    };
    const overlay = evaluateConditional(s);
    expect(overlay.get('0:0:0')?.iconSlot).toBe(0);
    expect(overlay.get('0:0:1')?.iconSlot).toBe(1);
    expect(overlay.get('0:0:2')?.iconSlot).toBe(2);
  });

  it('classifies expanded Excel-style icon families with 3-slot and 5-slot thresholds', () => {
    expect(iconSetSlotCount('symbols3')).toBe(3);
    expect(iconSetSlotCount('trafficRim3')).toBe(3);
    expect(iconSetSlotCount('bars5')).toBe(5);
    expect(iconSetSlotCount('boxes5')).toBe(5);
    expect(iconSetSlotFor('symbols3', 0.9)).toBe(2);
    expect(iconSetSlotFor('bars5', 0.61)).toBe(3);
  });

  it('top-bottom selects the top-N values and ties at the threshold qualify', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    [10, 20, 30, 30, 40, 50].forEach((v, i) => {
      s = seedNumber(s, 0, i, v);
    });
    const rule: ConditionalRule = {
      kind: 'top-bottom',
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 5 },
      mode: 'top',
      n: 3,
      apply: { fill: '#abc' },
    };
    s = { ...s, conditional: { ...s.conditional, rules: [rule] } };
    const overlay = evaluateConditional(s);
    // Top 3 = 50, 40, 30 — both 30s tie at the cutoff so 4 cells qualify.
    expect(overlay.get('0:0:5')?.fill).toBe('#abc'); // 50
    expect(overlay.get('0:0:4')?.fill).toBe('#abc'); // 40
    expect(overlay.get('0:0:3')?.fill).toBe('#abc'); // 30
    expect(overlay.get('0:0:2')?.fill).toBe('#abc'); // 30 (tie)
    expect(overlay.get('0:0:1')?.fill).toBeUndefined(); // 20
    expect(overlay.get('0:0:0')?.fill).toBeUndefined(); // 10
  });

  it('top-bottom with percent picks ceil(count * n / 100) values', () => {
    expect(topBottomThreshold([1, 2, 3, 4, 5, 6, 7, 8, 9, 10], 'bottom', 30, true)).toBe(3);
    expect(topBottomThreshold([1, 2, 3, 4, 5, 6, 7, 8, 9, 10], 'top', 20, true)).toBe(9);
  });

  it('average rule compares numeric cells against the range average', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    [1, 2, 9].forEach((v, i) => {
      s = seedNumber(s, 0, i, v);
    });
    const rule: ConditionalRule = {
      kind: 'average',
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 },
      mode: 'above',
      apply: { fill: '#avg' },
    };
    s = { ...s, conditional: { ...s.conditional, rules: [rule] } };
    const overlay = evaluateConditional(s);
    expect(overlay.get('0:0:0')?.fill).toBeUndefined();
    expect(overlay.get('0:0:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBe('#avg');
  });

  it('average std-dev rules compare against average plus or minus the selected tier', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    [-10, 0, 10, 30].forEach((v, i) => {
      s = seedNumber(s, 0, i, v);
    });
    const rule: ConditionalRule = {
      kind: 'average',
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 3 },
      mode: 'above-std-dev',
      stdDev: 1,
      apply: { fill: '#std' },
    };
    s = { ...s, conditional: { ...s.conditional, rules: [rule] } };
    const overlay = evaluateConditional(s);
    expect(overlay.get('0:0:0')?.fill).toBeUndefined();
    expect(overlay.get('0:0:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBe('#std');
  });

  it('text-contains rule matches text case-insensitively by default', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'Alpha' });
    s = seedCell(s, 0, 1, { kind: 'text', value: 'beta' });
    const rule: ConditionalRule = {
      kind: 'text-contains',
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      text: 'alp',
      apply: { fill: '#txt' },
    };
    s = { ...s, conditional: { ...s.conditional, rules: [rule] } };
    const overlay = evaluateConditional(s);
    expect(overlay.get('0:0:0')?.fill).toBe('#txt');
    expect(overlay.get('0:0:1')?.fill).toBeUndefined();
  });

  it('text rules support begins-with, ends-with, and not-contains modes', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'Alpha' });
    s = seedCell(s, 0, 1, { kind: 'text', value: 'Beta' });
    s = seedCell(s, 0, 2, { kind: 'text', value: 'Gamma' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'text-contains',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 },
            text: 'a',
            mode: 'ends-with',
            apply: { fill: '#end' },
          },
          {
            kind: 'text-contains',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 },
            text: 'g',
            mode: 'begins-with',
            apply: { color: '#begin' },
          },
          {
            kind: 'text-contains',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 },
            text: 'm',
            mode: 'not-contains',
            apply: { bold: true },
          },
        ],
      },
    };
    const overlay = evaluateConditional(s);
    expect(overlay.get('0:0:0')?.fill).toBe('#end');
    expect(overlay.get('0:0:1')?.bold).toBe(true);
    expect(overlay.get('0:0:2')?.fill).toBe('#end');
    expect(overlay.get('0:0:2')?.color).toBe('#begin');
  });

  it('duplicates fires on values that appear more than once in range', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'a' });
    s = seedCell(s, 0, 1, { kind: 'text', value: 'b' });
    s = seedCell(s, 0, 2, { kind: 'text', value: 'a' });
    const rule: ConditionalRule = {
      kind: 'duplicates',
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 },
      apply: { fill: '#dup' },
    };
    s = { ...s, conditional: { ...s.conditional, rules: [rule] } };
    const overlay = evaluateConditional(s);
    expect(overlay.get('0:0:0')?.fill).toBe('#dup');
    expect(overlay.get('0:0:1')?.fill).toBeUndefined(); // unique 'b'
    expect(overlay.get('0:0:2')?.fill).toBe('#dup');
  });

  it('unique fires only on values that appear exactly once', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'a' });
    s = seedCell(s, 0, 1, { kind: 'text', value: 'b' });
    s = seedCell(s, 0, 2, { kind: 'text', value: 'a' });
    const rule: ConditionalRule = {
      kind: 'unique',
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 },
      apply: { fill: '#uni' },
    };
    s = { ...s, conditional: { ...s.conditional, rules: [rule] } };
    const overlay = evaluateConditional(s);
    expect(overlay.get('0:0:0')?.fill).toBeUndefined();
    expect(overlay.get('0:0:1')?.fill).toBe('#uni');
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
  });

  it('blanks / errors predicates classify cells by content kind', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    // (0,0) blank (no entry), (0,1) text, (0,2) error
    s = seedCell(s, 0, 1, { kind: 'text', value: 'x' });
    s = seedCell(s, 0, 2, { kind: 'error', code: 1, text: '#DIV/0!' });
    const range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 };
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          { kind: 'blanks', range, apply: { fill: '#bla' } },
          { kind: 'errors', range, apply: { fill: '#err' } },
        ],
      },
    };
    const overlay = evaluateConditional(s);
    expect(overlay.get('0:0:0')?.fill).toBe('#bla');
    expect(overlay.get('0:0:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBe('#err');
  });

  it('formula rule fires for comparator-prefix predicates and skips unparseable forms', () => {
    expect(parseFormulaPredicate('>10')?.test({ kind: 'number', value: 11 })).toBe(true);
    expect(parseFormulaPredicate('>10')?.test({ kind: 'number', value: 5 })).toBe(false);
    expect(parseFormulaPredicate('<>"foo"')?.test({ kind: 'text', value: 'bar' })).toBe(true);
    expect(parseFormulaPredicate('<>"foo"')?.test({ kind: 'text', value: 'foo' })).toBe(false);
    // `=A1>0` is reserved for engine evaluator (not implemented in v1) — null.
    expect(parseFormulaPredicate('=A1>0')).toBeNull();
    // Bare `=42` after stripping the leading `=` has no comparator — null.
    expect(parseFormulaPredicate('=42')).toBeNull();
    // `==42` after stripping the leading `=` becomes `=42` which matches.
    expect(parseFormulaPredicate('==42')?.test({ kind: 'number', value: 42 })).toBe(true);
  });

  it('formula rules evaluate A1 references relative to the rule range anchor', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 1);
    s = seedNumber(s, 1, 0, -1);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 1, c1: 1 },
            formula: '=A1>0',
            apply: { fill: '#ref' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#ref');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
  });

  it('formula rules honor absolute rows and columns in A1 references', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 1);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 1, c1: 1 },
            formula: '=$A$1>0',
            apply: { fill: '#abs' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#abs');
    expect(overlay.get('0:1:1')?.fill).toBe('#abs');
  });

  it('formula rules combine simple comparisons with AND/OR/NOT', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 7);
    s = seedNumber(s, 0, 1, 1);
    s = seedNumber(s, 1, 0, 3);
    s = seedNumber(s, 1, 1, 1);
    s = seedNumber(s, 2, 0, 7);
    s = seedNumber(s, 2, 1, 9);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 2, c1: 2 },
            formula: '=AND(A1>5,OR(B1=1,NOT(A1<5)))',
            apply: { fill: '#logic' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#logic');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
    expect(overlay.get('0:2:2')?.fill).toBe('#logic');
  });

  it('formula rules evaluate aggregate ranges relative to the rule range anchor', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 1);
    s = seedNumber(s, 1, 0, 2);
    s = seedNumber(s, 2, 0, 3);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 1, c1: 2 },
            formula: '=SUM(A1:A3)>5',
            apply: { fill: '#sum' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#sum');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
  });

  it('formula rules support absolute aggregate ranges and aggregate operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 1);
    s = seedNumber(s, 1, 0, 2);
    s = seedNumber(s, 2, 0, 3);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 1, c1: 2 },
            formula: '=AND(AVERAGE($A$1:$A$3)=2,MAX($A$1:$A$3)=COUNT($A$1:$A$3))',
            apply: { fill: '#agg' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#agg');
    expect(overlay.get('0:1:2')?.fill).toBe('#agg');
  });

  it('formula rules evaluate COUNTIF with relative criteria references', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 7);
    s = seedNumber(s, 1, 0, 3);
    s = seedNumber(s, 2, 0, 7);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 2, c1: 1 },
            formula: '=COUNTIF($A$1:$A$3,A1)>1',
            apply: { fill: '#countif-ref' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#countif-ref');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
    expect(overlay.get('0:2:1')?.fill).toBe('#countif-ref');
  });

  it('formula rules evaluate COUNTIF with comparator criteria', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 1);
    s = seedNumber(s, 1, 0, 2);
    s = seedNumber(s, 2, 0, 3);
    s = seedCell(s, 0, 1, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 1, { kind: 'text', value: 'north' });
    s = seedCell(s, 2, 1, { kind: 'text', value: 'South' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=AND(COUNTIF($A$1:$A$3,">1")=2,COUNTIF($B$1:$B$3,"north")=2)',
            apply: { fill: '#countif-criteria' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#countif-criteria');
  });

  it('formula rules evaluate COUNTIF wildcard criteria with tilde escapes', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'Northeast' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'N* literal' });
    s = seedCell(s, 3, 0, { kind: 'text', value: 'South' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(COUNTIF($A$1:$A$4,"Nor*")=2,COUNTIF($A$1:$A$4,"N~* literal")=1)',
            apply: { fill: '#countif-wild' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=COUNTIF($A$1:$A$4,"<>S*")=3',
            apply: { fill: '#countif-not-wild' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#countif-wild');
    expect(overlay.get('0:1:1')?.fill).toBe('#countif-not-wild');
  });

  it('formula rules evaluate COUNTIFS with multiple criteria ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'Northwest' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'South' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 8);
    s = seedNumber(s, 2, 1, 15);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=COUNTIFS($A$1:$A$3,"North*",$B$1:$B$3,">10")=1',
            apply: { fill: '#countifs' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#countifs');
  });

  it('formula rules leave mismatched COUNTIFS ranges unapplied', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedNumber(s, 0, 1, 12);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=COUNTIFS($A$1:$A$2,"North*",$B$1:$B$3,">10")=1',
            apply: { fill: '#countifs-mismatch' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate simple arithmetic expressions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 4);
    s = seedNumber(s, 0, 1, 3);
    s = seedNumber(s, 1, 0, 2);
    s = seedNumber(s, 1, 1, 3);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 1, c1: 2 },
            formula: '=A1+B1*2>8',
            apply: { fill: '#math' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#math');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate arithmetic over aggregate operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 1);
    s = seedNumber(s, 1, 0, 2);
    s = seedNumber(s, 2, 0, 3);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=SUM($A$1:$A$3)/COUNT($A$1:$A$3)=2',
            apply: { fill: '#math-agg' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#math-agg');
  });

  it('formula arithmetic keeps negative numeric literals attached to the operand', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 4);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=A1*-1=-4',
            apply: { fill: '#negative' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#negative');
  });

  it('formula rules evaluate exponent arithmetic', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 3);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 2 },
            formula: '=$A$1^2=9',
            apply: { fill: '#pow' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 2 },
            formula: '=2^3^2=512',
            apply: { fill: '#pow-right' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#pow');
    expect(overlay.get('0:0:2')?.fill).toBe('#pow');
    expect(overlay.get('0:1:1')?.fill).toBe('#pow-right');
    expect(overlay.get('0:1:2')?.fill).toBe('#pow-right');
  });

  it('formula rules accept explicit current-sheet A1 references', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 1);
    s = seedNumber(s, 1, 0, -1);
    s = seedNumber(s, 2, 0, 3);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 1, c1: 2 },
            formula: "=AND(Sheet1!A1>0,SUM('Sheet1'!$A$1:$A$3)>2)",
            apply: { fill: '#sheet' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#sheet');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate IF with boolean branches', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 7);
    s = seedNumber(s, 0, 1, 3);
    s = seedNumber(s, 1, 0, 7);
    s = seedNumber(s, 1, 1, 1);
    s = seedNumber(s, 2, 0, 2);
    s = seedNumber(s, 2, 1, 9);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 2, c1: 2 },
            formula: '=IF(A1>5,B1>2,FALSE)',
            apply: { fill: '#if' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#if');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
    expect(overlay.get('0:2:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate ISBLANK/ISERROR/ISNUMBER/ISTEXT predicates', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'label' });
    s = seedNumber(s, 1, 0, 42);
    s = seedCell(s, 2, 0, { kind: 'error', code: 1, text: '#DIV/0!' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 3, c1: 1 },
            formula: '=OR(ISTEXT(A1),ISNUMBER(A1),ISERROR(A1),ISBLANK(A1))',
            apply: { fill: '#is' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#is');
    expect(overlay.get('0:1:1')?.fill).toBe('#is');
    expect(overlay.get('0:2:1')?.fill).toBe('#is');
    expect(overlay.get('0:3:1')?.fill).toBe('#is');
  });

  it('formula rules leave unsupported sheet-qualified references unapplied', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 1);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=Sheet2!A1>0',
            apply: { fill: '#sheet2' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBeUndefined();
  });

  it('iconSetSlotFor honors the family threshold table', () => {
    expect(iconSetSlotFor('arrows3', 0)).toBe(0);
    expect(iconSetSlotFor('arrows3', 0.5)).toBe(1);
    expect(iconSetSlotFor('arrows3', 0.9)).toBe(2);
    expect(iconSetSlotFor('arrows5', 0.1)).toBe(0);
    expect(iconSetSlotFor('arrows5', 0.3)).toBe(1);
    expect(iconSetSlotFor('arrows5', 0.5)).toBe(2);
    expect(iconSetSlotFor('arrows5', 0.7)).toBe(3);
    expect(iconSetSlotFor('arrows5', 0.95)).toBe(4);
  });
});
