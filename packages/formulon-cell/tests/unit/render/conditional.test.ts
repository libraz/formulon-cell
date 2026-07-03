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

  it('formula rules combine simple comparisons with XOR', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 7);
    s = seedNumber(s, 0, 1, 1);
    s = seedNumber(s, 1, 0, 7);
    s = seedNumber(s, 1, 1, 9);
    s = seedNumber(s, 2, 0, 3);
    s = seedNumber(s, 2, 1, 9);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 2, c1: 2 },
            formula: '=XOR(A1>5,B1>5,FALSE)',
            apply: { fill: '#xor' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=AND(A1>5,B1>0)=TRUE()',
            apply: { fill: '#logical-comparison' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=N(XOR(A1>5,B1>5,FALSE))=1',
            apply: { fill: '#logical-coerce' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#xor');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
    expect(overlay.get('0:2:2')?.fill).toBe('#xor');
    expect(overlay.get('0:0:3')?.fill).toBe('#logical-comparison');
    expect(overlay.get('0:0:4')?.fill).toBe('#logical-coerce');
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

  it('formula rules evaluate COUNTA/COUNTBLANK/PRODUCT aggregate ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 2);
    s = seedCell(s, 1, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 2, 0, { kind: 'bool', value: true });
    s = seedCell(s, 4, 0, { kind: 'error', code: 1, text: '#DIV/0!' });
    s = seedNumber(s, 0, 1, 2);
    s = seedNumber(s, 1, 1, 3);
    s = seedNumber(s, 2, 1, 4);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=AND(COUNTA($A$1:$A$5)=4,COUNTBLANK($A$1:$A$5)=1)',
            apply: { fill: '#counta-countblank' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=PRODUCT($B$1:$B$3)=24',
            apply: { fill: '#product' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#counta-countblank');
    expect(overlay.get('0:0:3')?.fill).toBe('#product');
  });

  it('formula rules evaluate MEDIAN aggregate ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 12);
    s = seedNumber(s, 1, 0, 8);
    s = seedNumber(s, 2, 0, 15);
    s = seedCell(s, 3, 0, { kind: 'text', value: 'ignored' });
    s = seedNumber(s, 4, 0, 4);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=MEDIAN($A$1:$A$5)=10',
            apply: { fill: '#median-even' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=MEDIAN($A$1:$A$3)=12',
            apply: { fill: '#median-odd' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#median-even');
    expect(overlay.get('0:0:2')?.fill).toBe('#median-odd');
  });

  it('formula rules evaluate multi-argument aggregate operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 2);
    s = seedNumber(s, 0, 1, 4);
    s = seedNumber(s, 0, 2, 6);
    s = seedNumber(s, 1, 0, 8);
    s = seedCell(s, 1, 1, { kind: 'text', value: 'ignored' });
    s = seedCell(s, 1, 2, { kind: 'bool', value: true });
    s = seedNumber(s, 3, 0, 1);
    s = seedNumber(s, 3, 1, 2);
    s = seedNumber(s, 3, 2, 3);
    s = seedNumber(s, 3, 3, 2);
    s = seedNumber(s, 3, 4, 4);
    s = seedNumber(s, 3, 5, 7);
    s = seedNumber(s, 4, 0, 1);
    s = seedNumber(s, 4, 1, 2);
    s = seedNumber(s, 4, 2, 3);
    s = seedNumber(s, 4, 3, 4);
    s = seedNumber(s, 5, 0, 0.1);
    s = seedNumber(s, 5, 1, 0.2);
    s = seedNumber(s, 5, 2, 0.3);
    s = seedNumber(s, 5, 3, 0.4);
    s = seedNumber(s, 6, 0, 10);
    s = seedNumber(s, 6, 1, 20);
    s = seedNumber(s, 7, 0, 20);
    s = seedNumber(s, 7, 1, 40);
    s = seedNumber(s, 6, 3, 15);
    s = seedNumber(s, 6, 4, 15);
    s = seedNumber(s, 7, 3, 15);
    s = seedNumber(s, 7, 4, 45);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=AND(SUM(A1,B1,$C$1:$C$1)=12,PRODUCT(A1,B1,3)=24)',
            apply: { fill: '#multi-sum-product' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=AND(AVERAGE($A$1:$C$1,A2)=5,MEDIAN($A$1:$C$1,A2)=5)',
            apply: { fill: '#multi-average-median' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=AND(AVERAGEA($A$1:$C$2)=3.5,MINA($A$1:$C$2)=0,MAXA($A$1:$C$2)=8)',
            apply: { fill: '#aggregate-a' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=AND(COUNT(A1,B1,B2)=2,COUNTA(A1,B2,C3)=2,COUNTBLANK(B2,C3)=1)',
            apply: { fill: '#multi-counts' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 3, r1: 1, c1: 3 },
            formula: '=AND(ROUND(STDEV.S($A$1:$C$1),6)=2,ROUND(STDEV.P($A$1:$C$1),6)=1.632993)',
            apply: { fill: '#stdev-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 4, r1: 1, c1: 4 },
            formula: '=AND(VAR.S(A1,B1,C1)=4,ROUND(VAR.P(A1,B1,C1),6)=2.666667)',
            apply: { fill: '#var-args' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=AND(ROUND(STDEV($A$1:$C$1),6)=2,VAR(A1,B1,C1)=4)',
            apply: { fill: '#legacy-stdev-var' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=AND(ROUND(STDEVP($A$1:$C$1),6)=1.632993,ROUND(VARP(A1,B1,C1),6)=2.666667)',
            apply: { fill: '#legacy-stdevp-varp' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 5, r1: 1, c1: 5 },
            formula: '=STDEV.S(A1)=0',
            apply: { fill: '#stdev-s-single' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 6, r1: 1, c1: 6 },
            formula:
              '=AND(ROUND(GEOMEAN($A$1:$C$1),6)=3.634241,ROUND(HARMEAN(A1,B1,C1),6)=3.272727)',
            apply: { fill: '#geo-harmonic' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 7, r1: 1, c1: 7 },
            formula: '=GEOMEAN(A1,-1)=0',
            apply: { fill: '#geomean-negative' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 8, r1: 1, c1: 8 },
            formula: '=AND(DEVSQ(A1,B1,C1)=8,DEVSQ($A$1:$C$1,A2)=20)',
            apply: { fill: '#devsq' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 9, r1: 1, c1: 9 },
            formula: '=AND(ROUND(AVEDEV(A1,B1,C1),6)=1.333333,AVEDEV($A$1:$C$1,A2)=2)',
            apply: { fill: '#avedev' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 10, r1: 1, c1: 10 },
            formula: '=ROUND(SKEW(A1,B1,C1,20),6)=1.763633',
            apply: { fill: '#skew' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 11, r1: 1, c1: 11 },
            formula: '=SKEW(A1,B1)=0',
            apply: { fill: '#skew-too-few' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 32, r1: 1, c1: 32 },
            formula:
              '=AND(ROUND(SKEW.P(A1,B1,C1,20),6)=1.018234,ROUND(SKEW.P($A$1:$C$1,20),6)=1.018234)',
            apply: { fill: '#skew-p' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 33, r1: 1, c1: 33 },
            formula: '=SKEW.P(A1,B1)=0',
            apply: { fill: '#skew-p-too-few' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 34, r1: 1, c1: 34 },
            formula:
              '=AND(ROUND(Z.TEST($A$1:$C$1,3,2),6)=0.193238,ROUND(Z.TEST($A$1:$C$1,3),6)=0.193238,ROUND(ZTEST($A$1:$C$1,3,2),6)=0.193238)',
            apply: { fill: '#z-test' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 35, r1: 1, c1: 35 },
            formula: '=Z.TEST($A$1:$C$1,3,0)=0',
            apply: { fill: '#z-test-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 36, r1: 1, c1: 36 },
            formula:
              '=AND(ROUND(F.TEST($A$4:$C$4,$D$4:$F$4),6)=0.272727,ROUND(FTEST($A$4:$C$4,$D$4:$F$4),6)=0.272727)',
            apply: { fill: '#f-test' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 37, r1: 1, c1: 37 },
            formula: '=F.TEST($A$1:$A$1,$D$4:$F$4)=0',
            apply: { fill: '#f-test-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 38, r1: 1, c1: 38 },
            formula:
              '=AND(ROUND(T.TEST($A$4:$C$4,$D$4:$F$4,1,1),6)=0.059041,ROUND(T.TEST($A$4:$C$4,$D$4:$F$4,2,1),6)=0.118083)',
            apply: { fill: '#t-test-paired' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 39, r1: 1, c1: 39 },
            formula:
              '=AND(ROUND(T.TEST($A$4:$C$4,$D$4:$F$4,2,2),6)=0.209875,ROUND(TTEST($A$4:$C$4,$D$4:$F$4,2,3),6)=0.245113)',
            apply: { fill: '#t-test-independent' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 40, r1: 1, c1: 40 },
            formula: '=T.TEST($A$4:$C$4,$D$4:$F$4,3,2)=0',
            apply: { fill: '#t-test-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 41, r1: 1, c1: 41 },
            formula:
              '=AND(ROUND(CHISQ.TEST($A$7:$B$8,$D$7:$E$8),6)=ROUND(CHISQ.DIST.RT(5.555555555555556,1),6),ROUND(CHITEST($A$7:$B$8,$D$7:$E$8),6)=ROUND(CHISQ.DIST.RT(5.555555555555556,1),6))',
            apply: { fill: '#chisq-test' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 42, r1: 1, c1: 42 },
            formula: '=CHISQ.TEST($A$7:$B$8,$D$7:$F$8)=0',
            apply: { fill: '#chisq-test-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 12, r1: 1, c1: 12 },
            formula: '=ROUND(KURT(A1,B1,C1,20),6)=3.228',
            apply: { fill: '#kurt' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 13, r1: 1, c1: 13 },
            formula: '=KURT(A1,B1,C1)=0',
            apply: { fill: '#kurt-too-few' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 14, r1: 1, c1: 14 },
            formula: '=AND(MODE.SNGL(A1,B1,C1,B1)=4,MODE.SNGL($A$1:$C$1,4)=4)',
            apply: { fill: '#mode-single' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 14, r1: 0, c1: 14 },
            formula: '=AND(MODE(A1,B1,C1,B1)=4,MODE($A$1:$C$1,4)=4)',
            apply: { fill: '#legacy-mode' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 15, r1: 1, c1: 15 },
            formula: '=MODE.SNGL(A1,B1,C1)=2',
            apply: { fill: '#mode-no-duplicate' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 16, r1: 1, c1: 16 },
            formula:
              '=AND(ROUND(CORREL($A$4:$C$4,$D$4:$F$4),6)=0.993399,ROUND(COVARIANCE.P($A$4:$C$4,$D$4:$F$4),6)=1.666667,COVARIANCE.S($A$4:$C$4,$D$4:$F$4)=2.5)',
            apply: { fill: '#paired-stats' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 16, r1: 0, c1: 16 },
            formula:
              '=AND(ROUND(PEARSON($A$4:$C$4,$D$4:$F$4),6)=0.993399,ROUND(COVAR($A$4:$C$4,$D$4:$F$4),6)=1.666667)',
            apply: { fill: '#legacy-paired-stats' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 19, r1: 1, c1: 19 },
            formula:
              '=AND(SLOPE($D$4:$F$4,$A$4:$C$4)=2.5,ROUND(INTERCEPT($D$4:$F$4,$A$4:$C$4),6)=-0.666667,ROUND(RSQ($D$4:$F$4,$A$4:$C$4),6)=0.986842)',
            apply: { fill: '#regression-stats' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 21, r1: 1, c1: 21 },
            formula:
              '=AND(ROUND(STEYX($D$4:$F$4,$A$4:$C$4),6)=0.408248,ROUND(FORECAST.LINEAR(4,$D$4:$F$4,$A$4:$C$4),6)=9.333333,ROUND(FORECAST(4,$D$4:$F$4,$A$4:$C$4),6)=9.333333)',
            apply: { fill: '#forecast-stats' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 17, r1: 1, c1: 17 },
            formula: '=CORREL($A$4:$B$4,$D$4:$F$4)=1',
            apply: { fill: '#paired-stats-mismatch' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 18, r1: 1, c1: 18 },
            formula: '=CORREL($A$1:$C$1,$B$1:$B$3)=1',
            apply: { fill: '#paired-stats-zero-variance' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 20, r1: 1, c1: 20 },
            formula: '=SLOPE($A$1:$C$1,$B$1:$B$3)=1',
            apply: { fill: '#regression-zero-x-variance' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 22, r1: 1, c1: 22 },
            formula: '=FORECAST("x",$D$4:$F$4,$A$4:$C$4)=1',
            apply: { fill: '#forecast-nonnumeric-x' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 23, r1: 1, c1: 23 },
            formula: '=AND(PROB($A$5:$D$5,$A$6:$D$6,2,3)=0.5,PROB($A$5:$D$5,$A$6:$D$6,4)=0.4)',
            apply: { fill: '#probability-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 24, r1: 1, c1: 24 },
            formula: '=PROB($A$5:$D$5,$A$6:$C$6,2)=0.2',
            apply: { fill: '#probability-mismatch' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 25, r1: 1, c1: 25 },
            formula: '=PROB($A$5:$C$5,$A$6:$C$6,2)=0.2',
            apply: { fill: '#probability-total-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 26, r1: 1, c1: 26 },
            formula:
              '=AND(SUMX2MY2($A$4:$C$4,$D$4:$F$4)=-55,SUMX2PY2($A$4:$C$4,$D$4:$F$4)=83,SUMXMY2($A$4:$C$4,$D$4:$F$4)=21)',
            apply: { fill: '#sumx-paired' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 27, r1: 1, c1: 27 },
            formula: '=SUMXMY2($A$4:$B$4,$D$4:$F$4)=1',
            apply: { fill: '#sumx-mismatch' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 28, r1: 1, c1: 28 },
            formula:
              '=AND(SUBTOTAL(9,$A$1:$C$1)=12,SUBTOTAL(101,$A$1:$C$1)=4,SUBTOTAL(103,$A$1:$C$2)=6,ROUND(SUBTOTAL(107,$A$1:$C$1),6)=2)',
            apply: { fill: '#subtotal' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 29, r1: 1, c1: 29 },
            formula: '=SUBTOTAL(12,$A$1:$C$1)=0',
            apply: { fill: '#subtotal-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 30, r1: 1, c1: 30 },
            formula:
              '=AND(AGGREGATE(9,0,$A$1:$C$1)=12,AGGREGATE(1,0,$A$1:$C$1)=4,ROUND(AGGREGATE(7,0,$A$1:$C$1),6)=2)',
            apply: { fill: '#aggregate-function' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 31, r1: 1, c1: 31 },
            formula: '=AGGREGATE(20,0,$A$1:$C$1)=0',
            apply: { fill: '#aggregate-function-invalid' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#multi-sum-product');
    expect(overlay.get('0:0:4')?.fill).toBe('#multi-average-median');
    expect(overlay.get('0:0:5')?.fill).toBe('#aggregate-a');
    expect(overlay.get('0:1:2')?.fill).toBe('#multi-counts');
    expect(overlay.get('0:1:3')?.fill).toBe('#stdev-range');
    expect(overlay.get('0:1:4')?.fill).toBe('#var-args');
    expect(overlay.get('0:0:6')?.fill).toBe('#legacy-stdev-var');
    expect(overlay.get('0:0:7')?.fill).toBe('#legacy-stdevp-varp');
    expect(overlay.get('0:1:5')?.fill).toBeUndefined();
    expect(overlay.get('0:1:6')?.fill).toBe('#geo-harmonic');
    expect(overlay.get('0:1:7')?.fill).toBeUndefined();
    expect(overlay.get('0:1:8')?.fill).toBe('#devsq');
    expect(overlay.get('0:1:9')?.fill).toBe('#avedev');
    expect(overlay.get('0:1:10')?.fill).toBe('#skew');
    expect(overlay.get('0:1:11')?.fill).toBeUndefined();
    expect(overlay.get('0:1:12')?.fill).toBe('#kurt');
    expect(overlay.get('0:1:13')?.fill).toBeUndefined();
    expect(overlay.get('0:0:14')?.fill).toBe('#legacy-mode');
    expect(overlay.get('0:1:14')?.fill).toBe('#mode-single');
    expect(overlay.get('0:1:15')?.fill).toBeUndefined();
    expect(overlay.get('0:0:16')?.fill).toBe('#legacy-paired-stats');
    expect(overlay.get('0:1:16')?.fill).toBe('#paired-stats');
    expect(overlay.get('0:1:17')?.fill).toBeUndefined();
    expect(overlay.get('0:1:18')?.fill).toBeUndefined();
    expect(overlay.get('0:1:19')?.fill).toBe('#regression-stats');
    expect(overlay.get('0:1:20')?.fill).toBeUndefined();
    expect(overlay.get('0:1:21')?.fill).toBe('#forecast-stats');
    expect(overlay.get('0:1:22')?.fill).toBeUndefined();
    expect(overlay.get('0:1:23')?.fill).toBe('#probability-range');
    expect(overlay.get('0:1:24')?.fill).toBeUndefined();
    expect(overlay.get('0:1:25')?.fill).toBeUndefined();
    expect(overlay.get('0:1:26')?.fill).toBe('#sumx-paired');
    expect(overlay.get('0:1:27')?.fill).toBeUndefined();
    expect(overlay.get('0:1:28')?.fill).toBe('#subtotal');
    expect(overlay.get('0:1:29')?.fill).toBeUndefined();
    expect(overlay.get('0:1:30')?.fill).toBe('#aggregate-function');
    expect(overlay.get('0:1:31')?.fill).toBeUndefined();
    expect(overlay.get('0:1:32')?.fill).toBe('#skew-p');
    expect(overlay.get('0:1:33')?.fill).toBeUndefined();
    expect(overlay.get('0:1:34')?.fill).toBe('#z-test');
    expect(overlay.get('0:1:35')?.fill).toBeUndefined();
    expect(overlay.get('0:1:36')?.fill).toBe('#f-test');
    expect(overlay.get('0:1:37')?.fill).toBeUndefined();
    expect(overlay.get('0:1:38')?.fill).toBe('#t-test-paired');
    expect(overlay.get('0:1:39')?.fill).toBe('#t-test-independent');
    expect(overlay.get('0:1:40')?.fill).toBeUndefined();
    expect(overlay.get('0:1:41')?.fill).toBe('#chisq-test');
    expect(overlay.get('0:1:42')?.fill).toBeUndefined();
  });

  it('formula rules evaluate statistical functions over dynamic ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 2);
    s = seedNumber(s, 0, 1, 4);
    s = seedNumber(s, 0, 2, 6);
    s = seedNumber(s, 3, 0, 1);
    s = seedNumber(s, 3, 1, 2);
    s = seedNumber(s, 3, 2, 3);
    s = seedNumber(s, 3, 3, 2);
    s = seedNumber(s, 3, 4, 4);
    s = seedNumber(s, 3, 5, 7);
    s = seedNumber(s, 4, 0, 1);
    s = seedNumber(s, 4, 1, 2);
    s = seedNumber(s, 4, 2, 3);
    s = seedNumber(s, 4, 3, 4);
    s = seedNumber(s, 5, 0, 0.1);
    s = seedNumber(s, 5, 1, 0.2);
    s = seedNumber(s, 5, 2, 0.3);
    s = seedNumber(s, 5, 3, 0.4);
    s = seedNumber(s, 6, 0, 10);
    s = seedNumber(s, 6, 1, 20);
    s = seedNumber(s, 7, 0, 20);
    s = seedNumber(s, 7, 1, 40);
    s = seedNumber(s, 6, 3, 15);
    s = seedNumber(s, 6, 4, 15);
    s = seedNumber(s, 7, 3, 15);
    s = seedNumber(s, 7, 4, 45);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula:
              '=AND(ROUND(CORREL(OFFSET(A4,0,0,1,3),INDIRECT("D4:F4")),6)=0.993399,ROUND(F.TEST(OFFSET(A4,0,0,1,3),INDIRECT("D4:F4")),6)=0.272727,ROUND(FORECAST.LINEAR(4,OFFSET(D4,0,0,1,3),INDIRECT("A4:C4")),6)=9.333333)',
            apply: { fill: '#stats-dynamic-paired' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 6, r1: 1, c1: 6 },
            formula:
              '=AND(ROUND(Z.TEST(OFFSET(A1,0,0,1,3),3,2),6)=0.193238,ROUND(T.TEST(INDIRECT("A4:C4"),OFFSET(D4,0,0,1,3),2,2),6)=0.209875,PROB(OFFSET(A5,0,0,1,4),INDIRECT("A6:D6"),2,3)=0.5)',
            apply: { fill: '#stats-dynamic-tests' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 6, r1: 2, c1: 6 },
            formula:
              '=ROUND(CHISQ.TEST(OFFSET(A7,0,0,2,2),INDIRECT("D7:E8")),6)=ROUND(CHISQ.DIST.RT(5.555555555555556,1),6)',
            apply: { fill: '#stats-dynamic-chisq' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 3, c0: 6, r1: 3, c1: 6 },
            formula: '=CHISQ.TEST(OFFSET(A7,0,0,2,2),INDIRECT("D7:F8"))=0',
            apply: { fill: '#stats-dynamic-mismatch' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:6')?.fill).toBe('#stats-dynamic-paired');
    expect(overlay.get('0:1:6')?.fill).toBe('#stats-dynamic-tests');
    expect(overlay.get('0:2:6')?.fill).toBe('#stats-dynamic-chisq');
    expect(overlay.get('0:3:6')?.fill).toBeUndefined();
  });

  it('formula rules evaluate LARGE and SMALL ranked range operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 12);
    s = seedNumber(s, 1, 0, 8);
    s = seedNumber(s, 2, 0, 15);
    s = seedCell(s, 3, 0, { kind: 'text', value: 'ignored' });
    s = seedNumber(s, 4, 0, 4);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(LARGE($A$1:$A$5,2)=12,SMALL($A$1:$A$5,2)=8)',
            apply: { fill: '#ranked-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=LARGE($A$1:$A$5,5)=0',
            apply: { fill: '#ranked-out-of-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=AND(PERCENTILE.INC($A$1:$A$5,0.5)=10,QUARTILE.INC($A$1:$A$5,3)=12.75)',
            apply: { fill: '#percentile-quartile' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 3, r1: 1, c1: 3 },
            formula: '=AND(PERCENTILE($A$1:$A$5,0.5)=10,QUARTILE($A$1:$A$5,3)=12.75)',
            apply: { fill: '#legacy-percentile-quartile' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=PERCENTILE.INC($A$1:$A$5,1.2)=15',
            apply: { fill: '#percentile-out-of-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=QUARTILE.INC($A$1:$A$5,5)=15',
            apply: { fill: '#quartile-out-of-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=AND(PERCENTRANK.INC($A$1:$A$5,10)=0.5,PERCENTRANK.INC($A$1:$A$5,12,2)=0.66)',
            apply: { fill: '#percentrank' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 6, r1: 1, c1: 6 },
            formula: '=AND(PERCENTRANK($A$1:$A$5,10)=0.5,PERCENTRANK($A$1:$A$5,12,2)=0.66)',
            apply: { fill: '#legacy-percentrank' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=PERCENTRANK.INC($A$1:$A$5,10,)=0.5',
            apply: { fill: '#percentrank-omitted-significance' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=PERCENTRANK.INC($A$1:$A$5,20)=1',
            apply: { fill: '#percentrank-out-of-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=AND(PERCENTILE.EXC($A$1:$A$5,0.5)=10,QUARTILE.EXC($A$1:$A$5,1)=5)',
            apply: { fill: '#percentile-exc' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 10, r1: 0, c1: 10 },
            formula: '=AND(PERCENTRANK.EXC($A$1:$A$5,10)=0.5,PERCENTRANK.EXC($A$1:$A$5,12,2)=0.6)',
            apply: { fill: '#percentrank-exc' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 11, r1: 0, c1: 11 },
            formula: '=PERCENTILE.EXC($A$1:$A$5,0.1)=4',
            apply: { fill: '#percentile-exc-out-of-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 12, r1: 0, c1: 12 },
            formula: '=QUARTILE.EXC($A$1:$A$5,0)=4',
            apply: { fill: '#quartile-exc-out-of-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 13, r1: 0, c1: 13 },
            formula: '=AND(AGGREGATE(14,0,$A$1:$A$5,2)=12,AGGREGATE(15,0,$A$1:$A$5,2)=8)',
            apply: { fill: '#aggregate-ranked' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 14, r1: 0, c1: 14 },
            formula: '=AND(AGGREGATE(16,0,$A$1:$A$5,0.5)=10,AGGREGATE(17,0,$A$1:$A$5,3)=12.75)',
            apply: { fill: '#aggregate-percentile-inc' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 15, r1: 0, c1: 15 },
            formula: '=AND(AGGREGATE(18,0,$A$1:$A$5,0.5)=10,AGGREGATE(19,0,$A$1:$A$5,1)=5)',
            apply: { fill: '#aggregate-percentile-exc' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 16, r1: 0, c1: 16 },
            formula: '=AGGREGATE(14,8,$A$1:$A$5,2)=12',
            apply: { fill: '#aggregate-invalid-option' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#ranked-range');
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBe('#percentile-quartile');
    expect(overlay.get('0:1:3')?.fill).toBe('#legacy-percentile-quartile');
    expect(overlay.get('0:0:4')?.fill).toBeUndefined();
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
    expect(overlay.get('0:0:6')?.fill).toBe('#percentrank');
    expect(overlay.get('0:1:6')?.fill).toBe('#legacy-percentrank');
    expect(overlay.get('0:0:7')?.fill).toBe('#percentrank-omitted-significance');
    expect(overlay.get('0:0:8')?.fill).toBeUndefined();
    expect(overlay.get('0:0:9')?.fill).toBe('#percentile-exc');
    expect(overlay.get('0:0:10')?.fill).toBe('#percentrank-exc');
    expect(overlay.get('0:0:11')?.fill).toBeUndefined();
    expect(overlay.get('0:0:12')?.fill).toBeUndefined();
    expect(overlay.get('0:0:13')?.fill).toBe('#aggregate-ranked');
    expect(overlay.get('0:0:14')?.fill).toBe('#aggregate-percentile-inc');
    expect(overlay.get('0:0:15')?.fill).toBe('#aggregate-percentile-exc');
    expect(overlay.get('0:0:16')?.fill).toBeUndefined();
  });

  it('formula rules evaluate RANK variants over numeric ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 12);
    s = seedNumber(s, 1, 0, 8);
    s = seedNumber(s, 2, 0, 12);
    s = seedNumber(s, 3, 0, 15);
    s = seedCell(s, 4, 0, { kind: 'text', value: 'ignored' });
    s = seedNumber(s, 5, 0, 4);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(RANK.EQ(12,$A$1:$A$6)=2,RANK(12,$A$1:$A$6)=2)',
            apply: { fill: '#rank-eq' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=RANK.AVG(12,$A$1:$A$6)=2.5',
            apply: { fill: '#rank-avg' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=RANK.EQ(12,$A$1:$A$6,1)=3',
            apply: { fill: '#rank-ascending' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=RANK.EQ(12,$A$1:$A$6,)=2',
            apply: { fill: '#rank-omitted-order' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=RANK.EQ(10,$A$1:$A$6)=3',
            apply: { fill: '#rank-not-found' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#rank-eq');
    expect(overlay.get('0:0:2')?.fill).toBe('#rank-avg');
    expect(overlay.get('0:0:3')?.fill).toBe('#rank-ascending');
    expect(overlay.get('0:0:4')?.fill).toBe('#rank-omitted-order');
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
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

  it('formula rules evaluate SUMIF with optional sum ranges', () => {
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
            formula: '=SUMIF($A$1:$A$3,"North*",$B$1:$B$3)=20',
            apply: { fill: '#sumif' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=SUMIF($B$1:$B$3,">10")=27',
            apply: { fill: '#sumif-self' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#sumif');
    expect(overlay.get('0:1:2')?.fill).toBe('#sumif-self');
  });

  it('formula rules evaluate SUMIFS with multiple criteria ranges', () => {
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
            formula: '=SUMIFS($B$1:$B$3,$A$1:$A$3,"North*",$B$1:$B$3,">10")=12',
            apply: { fill: '#sumifs' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#sumifs');
  });

  it('formula rules evaluate SUMPRODUCT over aligned ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 2);
    s = seedNumber(s, 1, 0, 3);
    s = seedNumber(s, 2, 0, 4);
    s = seedNumber(s, 0, 1, 5);
    s = seedCell(s, 1, 1, { kind: 'text', value: 'skip' });
    s = seedNumber(s, 2, 1, 7);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=SUMPRODUCT($A$1:$A$3,$B$1:$B$3)=38',
            apply: { fill: '#sumproduct' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=SUMPRODUCT($A$1:$A$2,$B$1:$B$3)=10',
            apply: { fill: '#sumproduct-mismatch' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#sumproduct');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate SUMPRODUCT over dynamic ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 2);
    s = seedNumber(s, 1, 0, 3);
    s = seedNumber(s, 2, 0, 4);
    s = seedNumber(s, 0, 1, 5);
    s = seedCell(s, 1, 1, { kind: 'text', value: 'skip' });
    s = seedNumber(s, 2, 1, 7);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=SUMPRODUCT(OFFSET(A1,0,0,3,1),INDIRECT("B1:B3"))=38',
            apply: { fill: '#sumproduct-dynamic' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=SUMPRODUCT(OFFSET(A1,0,0,2,1),INDIRECT("B1:B3"))=10',
            apply: { fill: '#sumproduct-dynamic-mismatch' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#sumproduct-dynamic');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
  });

  it('formula rules leave mismatched SUMIFS ranges unapplied', () => {
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
            formula: '=SUMIFS($B$1:$B$2,$A$1:$A$3,"North*")=12',
            apply: { fill: '#sumifs-mismatch' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate AVERAGEIF with optional average ranges', () => {
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
            formula: '=AVERAGEIF($A$1:$A$3,"North*",$B$1:$B$3)=10',
            apply: { fill: '#averageif' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=AVERAGEIF($B$1:$B$3,">10")=13.5',
            apply: { fill: '#averageif-self' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#averageif');
    expect(overlay.get('0:1:2')?.fill).toBe('#averageif-self');
  });

  it('formula rules evaluate AVERAGEIFS and fail closed on empty numeric matches', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'Northwest' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'South' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 8);
    s = seedNumber(s, 2, 1, 15);
    s = seedCell(s, 0, 3, { kind: 'text', value: 'not numeric' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=AVERAGEIFS($B$1:$B$3,$A$1:$A$3,"North*",$B$1:$B$3,">10")=12',
            apply: { fill: '#averageifs' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=AVERAGEIF($A$1:$A$1,"North",$D$1:$D$1)=0',
            apply: { fill: '#averageif-empty' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#averageifs');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate MINIFS and MAXIFS with multiple criteria ranges', () => {
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
            formula:
              '=AND(MINIFS($B$1:$B$3,$A$1:$A$3,"North*")=8,MAXIFS($B$1:$B$3,$A$1:$A$3,"North*")=12)',
            apply: { fill: '#minmaxifs' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#minmaxifs');
  });

  it('formula rules leave MINIFS/MAXIFS without numeric matches unapplied', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 0, 1, { kind: 'text', value: 'not numeric' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=MINIFS($B$1:$B$1,$A$1:$A$1,"North")=0',
            apply: { fill: '#minifs-empty' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate criteria functions over dynamic ranges', () => {
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
            formula: '=COUNTIF(OFFSET(A1,0,0,3,1),"North*")=2',
            apply: { fill: '#countif-offset' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=COUNTIFS(INDIRECT("A1:A3"),"North*",OFFSET(B1,0,0,3,1),">10")=1',
            apply: { fill: '#countifs-dynamic' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=SUMIF(INDIRECT("A1:A3"),"North*",OFFSET(B1,0,0,3,1))=20',
            apply: { fill: '#sumif-dynamic' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=AVERAGEIF(OFFSET(A1,0,0,3,1),"North*",INDIRECT("B1:B3"))=10',
            apply: { fill: '#averageif-dynamic' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula:
              '=SUMIFS(OFFSET(B1,0,0,3,1),INDIRECT("A1:A3"),"North*",INDIRECT("B1:B3"),">10")=12',
            apply: { fill: '#sumifs-dynamic' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula:
              '=AND(MINIFS(OFFSET(B1,0,0,3,1),INDIRECT("A1:A3"),"North*")=8,MAXIFS(INDIRECT("B1:B3"),OFFSET(A1,0,0,3,1),"North*")=12)',
            apply: { fill: '#minmaxifs-dynamic' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=COUNTIF(OFFSET(A1,0,0,10001,1),"North*")=2',
            apply: { fill: '#countif-dynamic-too-large' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=SUMIF(INDIRECT("Sheet2!A1:A3"),"North*",B1:B3)=20',
            apply: { fill: '#sumif-dynamic-other-sheet' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#countif-offset');
    expect(overlay.get('0:0:3')?.fill).toBe('#countifs-dynamic');
    expect(overlay.get('0:0:4')?.fill).toBe('#sumif-dynamic');
    expect(overlay.get('0:0:5')?.fill).toBe('#averageif-dynamic');
    expect(overlay.get('0:0:6')?.fill).toBe('#sumifs-dynamic');
    expect(overlay.get('0:0:7')?.fill).toBe('#minmaxifs-dynamic');
    expect(overlay.get('0:0:8')?.fill).toBeUndefined();
    expect(overlay.get('0:0:9')?.fill).toBeUndefined();
  });

  it('formula rules evaluate exact MATCH over row and column ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'East' });
    s = seedCell(s, 3, 0, { kind: 'text', value: 'A*B' });
    s = seedNumber(s, 0, 3, 10);
    s = seedNumber(s, 0, 4, 20);
    s = seedNumber(s, 0, 5, 30);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 2, c1: 1 },
            formula: '=ISNUMBER(MATCH(A1,$A$1:$A$3,0))',
            apply: { fill: '#match-col' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=MATCH(20,$D$1:$F$1,0)=2',
            apply: { fill: '#match-row' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=ISNA(MATCH("West",$A$1:$A$3,0))',
            apply: { fill: '#match-na' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=ISNA(MATCH("North",$A$1:$A$3,1))',
            apply: { fill: '#match-unsupported' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=MATCH("No*",$A$1:$A$3,0)=1',
            apply: { fill: '#match-wildcard' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 10, r1: 0, c1: 10 },
            formula: '=MATCH("A~*B",$A$4:$A$4,0)=1',
            apply: { fill: '#match-escaped-wildcard' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#match-col');
    expect(overlay.get('0:1:1')?.fill).toBe('#match-col');
    expect(overlay.get('0:2:1')?.fill).toBe('#match-col');
    expect(overlay.get('0:0:6')?.fill).toBe('#match-row');
    expect(overlay.get('0:0:7')?.fill).toBe('#match-na');
    expect(overlay.get('0:0:8')?.fill).toBe('#match-unsupported');
    expect(overlay.get('0:0:9')?.fill).toBe('#match-wildcard');
    expect(overlay.get('0:0:10')?.fill).toBe('#match-escaped-wildcard');
  });

  it('formula rules evaluate approximate MATCH over monotonic ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 5);
    s = seedNumber(s, 1, 0, 10);
    s = seedNumber(s, 2, 0, 20);
    s = seedNumber(s, 3, 0, 40);
    s = seedNumber(s, 0, 1, 40);
    s = seedNumber(s, 1, 1, 20);
    s = seedNumber(s, 2, 1, 10);
    s = seedNumber(s, 3, 1, 5);
    s = seedCell(s, 0, 3, { kind: 'text', value: 'East' });
    s = seedCell(s, 0, 4, { kind: 'text', value: 'North' });
    s = seedCell(s, 0, 5, { kind: 'text', value: 'South' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=MATCH(17,$A$1:$A$4,1)=2',
            apply: { fill: '#match-ascending' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=MATCH(17,$B$1:$B$4,-1)=2',
            apply: { fill: '#match-descending' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=MATCH("Nor",$D$1:$F$1,1)=1',
            apply: { fill: '#match-text-ascending' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 6, r1: 1, c1: 6 },
            formula: '=MATCH(17,$A$1:$A$4)=2',
            apply: { fill: '#match-omitted-approx' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 3, c0: 2, r1: 3, c1: 2 },
            formula: '=ISNA(MATCH(17,$A$1:$B$2,1))',
            apply: { fill: '#match-not-one-dimensional' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#match-ascending');
    expect(overlay.get('0:1:2')?.fill).toBe('#match-descending');
    expect(overlay.get('0:0:6')?.fill).toBe('#match-text-ascending');
    expect(overlay.get('0:1:6')?.fill).toBe('#match-omitted-approx');
    expect(overlay.get('0:3:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate XMATCH exact, wildcard, reverse search, and INDEX composition', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'East' });
    s = seedCell(s, 3, 0, { kind: 'text', value: 'North' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 8);
    s = seedNumber(s, 2, 1, 4);
    s = seedNumber(s, 3, 1, 14);
    s = seedCell(s, 5, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 5, 1, { kind: 'text', value: 'South' });
    s = seedCell(s, 5, 2, { kind: 'text', value: 'East' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=XMATCH("South",$A$1:$A$4)=2',
            apply: { fill: '#xmatch-exact' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=XMATCH("No*",$A$1:$A$4,2)=1',
            apply: { fill: '#xmatch-wildcard' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=XMATCH("North",$A$1:$A$4,0,-1)=4',
            apply: { fill: '#xmatch-reverse' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=INDEX($B$1:$B$4,XMATCH("North",$A$1:$A$4,0,-1))=14',
            apply: { fill: '#xmatch-index' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=XMATCH("East",$A$6:$C$6)=3',
            apply: { fill: '#xmatch-horizontal' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=ISNA(XMATCH("South",$A$1:$A$4,-1))',
            apply: { fill: '#xmatch-unsupported' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=XMATCH("North",$A$1:$A$4,,-1)=4',
            apply: { fill: '#xmatch-omitted-match-mode' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#xmatch-exact');
    expect(overlay.get('0:0:4')?.fill).toBe('#xmatch-wildcard');
    expect(overlay.get('0:0:5')?.fill).toBe('#xmatch-reverse');
    expect(overlay.get('0:0:6')?.fill).toBe('#xmatch-index');
    expect(overlay.get('0:0:7')?.fill).toBe('#xmatch-horizontal');
    expect(overlay.get('0:0:8')?.fill).toBe('#xmatch-unsupported');
    expect(overlay.get('0:0:9')?.fill).toBe('#xmatch-omitted-match-mode');
  });

  it('formula rules evaluate approximate XMATCH over monotonic ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 5);
    s = seedNumber(s, 1, 0, 10);
    s = seedNumber(s, 2, 0, 20);
    s = seedNumber(s, 3, 0, 40);
    s = seedCell(s, 0, 1, { kind: 'text', value: 'East' });
    s = seedCell(s, 1, 1, { kind: 'text', value: 'North' });
    s = seedCell(s, 2, 1, { kind: 'text', value: 'South' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=XMATCH(17,$A$1:$A$4,-1)=2',
            apply: { fill: '#xmatch-next-smaller' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=XMATCH(17,$A$1:$A$4,1)=3',
            apply: { fill: '#xmatch-next-larger' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
            formula: '=XMATCH("Nor",$B$1:$B$3,-1)=1',
            apply: { fill: '#xmatch-text-next-smaller' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 3, c0: 2, r1: 3, c1: 2 },
            formula: '=XMATCH("Nor",$B$1:$B$3,1)=2',
            apply: { fill: '#xmatch-text-next-larger' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#xmatch-next-smaller');
    expect(overlay.get('0:1:2')?.fill).toBe('#xmatch-next-larger');
    expect(overlay.get('0:2:2')?.fill).toBe('#xmatch-text-next-smaller');
    expect(overlay.get('0:3:2')?.fill).toBe('#xmatch-text-next-larger');
  });

  it('formula rules evaluate lookup and rank functions over dynamic ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'East' });
    s = seedCell(s, 3, 0, { kind: 'text', value: 'West' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 8);
    s = seedNumber(s, 2, 1, 4);
    s = seedNumber(s, 3, 1, 14);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=MATCH("South",OFFSET(A1,0,0,4,1),0)=2',
            apply: { fill: '#match-offset' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=XMATCH("West",INDIRECT("A1:A4"))=4',
            apply: { fill: '#xmatch-indirect' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=INDEX(OFFSET(B1,0,0,4,1),XMATCH("West",A1:A4))=14',
            apply: { fill: '#index-offset' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=VLOOKUP("South",OFFSET(A1,0,0,4,2),2,FALSE)=8',
            apply: { fill: '#vlookup-offset' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=LARGE(INDIRECT("B1:B4"),2)=12',
            apply: { fill: '#large-indirect' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=RANK(8,OFFSET(B1,0,0,4,1),0)=3',
            apply: { fill: '#rank-offset' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=MATCH("South",OFFSET(A1,0,0,2,2),0)=2',
            apply: { fill: '#match-dynamic-not-one-dimensional' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#match-offset');
    expect(overlay.get('0:0:3')?.fill).toBe('#xmatch-indirect');
    expect(overlay.get('0:0:4')?.fill).toBe('#index-offset');
    expect(overlay.get('0:0:5')?.fill).toBe('#vlookup-offset');
    expect(overlay.get('0:0:6')?.fill).toBe('#large-indirect');
    expect(overlay.get('0:0:7')?.fill).toBe('#rank-offset');
    expect(overlay.get('0:0:8')?.fill).toBeUndefined();
  });

  it('formula rules evaluate scalar CHOOSE operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'East' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 8);
    s = seedNumber(s, 2, 1, 4);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=CHOOSE(2,"North","South","East")="South"',
            apply: { fill: '#choose-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=CHOOSE(XMATCH("South",$A$1:$A$3),$B$1,$B$2,$B$3)=8',
            apply: { fill: '#choose-xmatch' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=CHOOSE(3,1+1,2+2,3+3)=6',
            apply: { fill: '#choose-expression' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=ISERROR(CHOOSE(4,"North","South","East"))',
            apply: { fill: '#choose-error' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#choose-text');
    expect(overlay.get('0:0:4')?.fill).toBe('#choose-xmatch');
    expect(overlay.get('0:0:5')?.fill).toBe('#choose-expression');
    expect(overlay.get('0:0:6')?.fill).toBe('#choose-error');
  });

  it('formula rules evaluate scalar SWITCH operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'East' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 8);
    s = seedNumber(s, 2, 1, 4);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=SWITCH(A1,"North","Region N","South","Region S","Other")="Region N"',
            apply: { fill: '#switch-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=SWITCH(XMATCH("South",$A$1:$A$3),1,$B$1,2,$B$2,3,$B$3)=8',
            apply: { fill: '#switch-xmatch' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=SWITCH("West","North",1,"South",2,99)=99',
            apply: { fill: '#switch-default' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=ISNA(SWITCH("West","North",1,"South",2))',
            apply: { fill: '#switch-error' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#switch-text');
    expect(overlay.get('0:0:4')?.fill).toBe('#switch-xmatch');
    expect(overlay.get('0:0:5')?.fill).toBe('#switch-default');
    expect(overlay.get('0:0:6')?.fill).toBe('#switch-error');
  });

  it('formula rules evaluate scalar IFS operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'East' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 8);
    s = seedNumber(s, 2, 1, 4);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=IFS(A1="North","Region N",A1="South","Region S",TRUE,"Other")="Region N"',
            apply: { fill: '#ifs-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=IFS(AND(A1="North",B1>10),B1,TRUE,0)=12',
            apply: { fill: '#ifs-and' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 3, r1: 2, c1: 3 },
            formula: '=IFS(B3>10,"High",TRUE,"Low")="Low"',
            apply: { fill: '#ifs-fallback' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=ISNA(IFS(A1="West","Missing",A1="Central","Missing"))',
            apply: { fill: '#ifs-error' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#ifs-text');
    expect(overlay.get('0:0:4')?.fill).toBe('#ifs-and');
    expect(overlay.get('0:2:3')?.fill).toBe('#ifs-fallback');
    expect(overlay.get('0:0:5')?.fill).toBe('#ifs-error');
  });

  it('formula rules evaluate scalar IF result operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 8);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 1, c1: 2 },
            formula: '=IF(A1="North","Region N","Other")="Region N"',
            apply: { fill: '#if-text-result' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 1, c1: 3 },
            formula: '=IF(AND(A1="North",B1>10),B1,0)=12',
            apply: { fill: '#if-number-result' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=IF(TRUE(),42,1/0)=42',
            apply: { fill: '#if-short-circuit-true' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=IF(FALSE(),1/0,"fallback")="fallback"',
            apply: { fill: '#if-short-circuit-false' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=IF(A1="North",,99)=0',
            apply: { fill: '#if-omitted-true' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 6, r1: 1, c1: 6 },
            formula: '=IF(A2="North",99)=FALSE',
            apply: { fill: '#if-omitted-false' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#if-text-result');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBe('#if-number-result');
    expect(overlay.get('0:1:3')?.fill).toBeUndefined();
    expect(overlay.get('0:0:4')?.fill).toBe('#if-short-circuit-true');
    expect(overlay.get('0:0:5')?.fill).toBe('#if-short-circuit-false');
    expect(overlay.get('0:0:6')?.fill).toBe('#if-omitted-true');
    expect(overlay.get('0:1:6')?.fill).toBe('#if-omitted-false');
  });

  it('formula rules evaluate INDEX over vector and rectangular ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'East' });
    s = seedNumber(s, 0, 3, 10);
    s = seedNumber(s, 0, 4, 20);
    s = seedNumber(s, 0, 5, 30);
    s = seedCell(s, 3, 3, { kind: 'text', value: 'Red' });
    s = seedCell(s, 3, 4, { kind: 'text', value: 'Blue' });
    s = seedCell(s, 4, 3, { kind: 'text', value: 'Green' });
    s = seedCell(s, 4, 4, { kind: 'text', value: 'Gold' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=INDEX($A$1:$A$3,MATCH("South",$A$1:$A$3,0))="South"',
            apply: { fill: '#index-column' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=INDEX($D$1:$F$1,2)=20',
            apply: { fill: '#index-row' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=INDEX($D$4:$E$5,2,2)="Gold"',
            apply: { fill: '#index-rect' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=ISERROR(INDEX($D$4:$E$5,3,1))',
            apply: { fill: '#index-error' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#index-column');
    expect(overlay.get('0:0:6')?.fill).toBe('#index-row');
    expect(overlay.get('0:0:7')?.fill).toBe('#index-rect');
    expect(overlay.get('0:0:8')?.fill).toBe('#index-error');
  });

  it('formula rules evaluate scalar OFFSET references', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 18);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=OFFSET(A1,1,0)="South"',
            apply: { fill: '#offset-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 1, c1: 3 },
            formula: '=OFFSET(A1,0,1)=12',
            apply: { fill: '#offset-relative' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=OFFSET(A1,0,0,1,1)="North"',
            apply: { fill: '#offset-sized' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=OFFSET(A1,0,0,2,1)="North"',
            apply: { fill: '#offset-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=OFFSET(OFFSET(A1,1,0),0,0)="South"',
            apply: { fill: '#offset-nested-offset' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=OFFSET(INDIRECT("A1"),1,1)=18',
            apply: { fill: '#offset-nested-indirect' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#offset-text');
    expect(overlay.get('0:0:3')?.fill).toBe('#offset-relative');
    expect(overlay.get('0:1:3')?.fill).toBeUndefined();
    expect(overlay.get('0:0:4')?.fill).toBe('#offset-sized');
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
    expect(overlay.get('0:0:6')?.fill).toBe('#offset-nested-offset');
    expect(overlay.get('0:0:7')?.fill).toBe('#offset-nested-indirect');
  });

  it('formula rules evaluate aggregate dynamic ranges from OFFSET and INDIRECT', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 10);
    s = seedNumber(s, 1, 0, 20);
    s = seedCell(s, 2, 0, { kind: 'text', value: 'not numeric' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=SUM(OFFSET(A1,0,0,2,1))=30',
            apply: { fill: '#offset-range-sum' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=COUNT(OFFSET(A1,0,0,3,1))=2',
            apply: { fill: '#offset-range-count' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=SUM(INDIRECT("A1:A2"))=30',
            apply: { fill: '#indirect-range-a1' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=SUM(INDIRECT("R1C1:R2C1",FALSE))=30',
            apply: { fill: '#indirect-range-r1c1' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=SUM(INDIRECT("RC[-5]:R[1]C[-5]",FALSE))=30',
            apply: { fill: '#indirect-range-relative-r1c1' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=SUM(OFFSET(A1,0,0,10001,1))=30',
            apply: { fill: '#offset-range-too-large' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=SUM(INDIRECT("Sheet2!A1:A2"))=30',
            apply: { fill: '#indirect-range-other-sheet' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=SUM(OFFSET(OFFSET(A1,0,0,2,1),0,0,2,1))=30',
            apply: { fill: '#offset-range-nested-offset' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=SUM(OFFSET(INDIRECT("A1:A2"),0,0,2,1))=30',
            apply: { fill: '#offset-range-nested-indirect' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#offset-range-sum');
    expect(overlay.get('0:0:2')?.fill).toBe('#offset-range-count');
    expect(overlay.get('0:0:3')?.fill).toBe('#indirect-range-a1');
    expect(overlay.get('0:0:4')?.fill).toBe('#indirect-range-r1c1');
    expect(overlay.get('0:0:5')?.fill).toBe('#indirect-range-relative-r1c1');
    expect(overlay.get('0:0:6')?.fill).toBeUndefined();
    expect(overlay.get('0:0:7')?.fill).toBeUndefined();
    expect(overlay.get('0:0:8')?.fill).toBe('#offset-range-nested-offset');
    expect(overlay.get('0:0:9')?.fill).toBe('#offset-range-nested-indirect');
  });

  it('formula rules evaluate scalar INDIRECT A1 references', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedCell(s, 0, 1, { kind: 'text', value: 'A1' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=INDIRECT("A1")="North"',
            apply: { fill: '#indirect-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=INDIRECT(B1,TRUE)="North"',
            apply: { fill: '#indirect-ref-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 1, c1: 4 },
            formula: '=INDIRECT("A1")="North"',
            apply: { fill: '#indirect-fixed' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=INDIRECT("R1C1",FALSE)="North"',
            apply: { fill: '#indirect-r1c1' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=INDIRECT("Sheet2!A1")="North"',
            apply: { fill: '#indirect-other-sheet' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=INDIRECT("$A$1")="North"',
            apply: { fill: '#indirect-absolute' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=INDIRECT("Sheet1!A1")="North"',
            apply: { fill: '#indirect-current-sheet' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=INDIRECT("\'Sheet1\'!A1")="North"',
            apply: { fill: '#indirect-quoted-current-sheet' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 10, r1: 0, c1: 10 },
            formula: '=INDIRECT("Sheet1!R1C1",FALSE)="North"',
            apply: { fill: '#indirect-r1c1-current-sheet' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 11, r1: 0, c1: 11 },
            formula: '=INDIRECT("RC[-11]",FALSE)="North"',
            apply: { fill: '#indirect-r1c1-relative' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            formula: '=INDIRECT("RC",FALSE)="North"',
            apply: { fill: '#indirect-r1c1-current-cell' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 12, r1: 0, c1: 12 },
            formula: '=INDIRECT("R[-1]C",FALSE)="North"',
            apply: { fill: '#indirect-r1c1-out-of-bounds' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 13, r1: 0, c1: 13 },
            formula: '=INDIRECT("Sheet2!R1C1",FALSE)="North"',
            apply: { fill: '#indirect-r1c1-other-sheet' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#indirect-text');
    expect(overlay.get('0:0:3')?.fill).toBe('#indirect-ref-text');
    expect(overlay.get('0:0:4')?.fill).toBe('#indirect-fixed');
    expect(overlay.get('0:1:4')?.fill).toBe('#indirect-fixed');
    expect(overlay.get('0:0:5')?.fill).toBe('#indirect-r1c1');
    expect(overlay.get('0:0:6')?.fill).toBeUndefined();
    expect(overlay.get('0:0:7')?.fill).toBe('#indirect-absolute');
    expect(overlay.get('0:0:8')?.fill).toBe('#indirect-current-sheet');
    expect(overlay.get('0:0:9')?.fill).toBe('#indirect-quoted-current-sheet');
    expect(overlay.get('0:0:10')?.fill).toBe('#indirect-r1c1-current-sheet');
    expect(overlay.get('0:0:11')?.fill).toBe('#indirect-r1c1-relative');
    expect(overlay.get('0:0:0')?.fill).toBe('#indirect-r1c1-current-cell');
    expect(overlay.get('0:0:12')?.fill).toBeUndefined();
    expect(overlay.get('0:0:13')?.fill).toBeUndefined();
  });

  it('formula rules evaluate exact VLOOKUP and HLOOKUP table lookups', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedNumber(s, 0, 1, 12);
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedNumber(s, 1, 1, 8);
    s = seedCell(s, 2, 0, { kind: 'text', value: 'East' });
    s = seedNumber(s, 2, 1, 4);
    s = seedCell(s, 4, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 4, 1, { kind: 'text', value: 'South' });
    s = seedCell(s, 4, 2, { kind: 'text', value: 'East' });
    s = seedNumber(s, 5, 0, 12);
    s = seedNumber(s, 5, 1, 8);
    s = seedNumber(s, 5, 2, 4);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=VLOOKUP("South",$A$1:$B$3,2,FALSE)=8',
            apply: { fill: '#vlookup-exact' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=VLOOKUP("Nor*",$A$1:$B$3,2,0)=12',
            apply: { fill: '#vlookup-wildcard' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=HLOOKUP("East",$A$5:$C$6,2,FALSE)=4',
            apply: { fill: '#hlookup-exact' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=ISNA(VLOOKUP("South",$A$1:$B$3,2,TRUE))',
            apply: { fill: '#vlookup-unsupported' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#vlookup-exact');
    expect(overlay.get('0:0:4')?.fill).toBe('#vlookup-wildcard');
    expect(overlay.get('0:0:5')?.fill).toBe('#hlookup-exact');
    expect(overlay.get('0:0:6')?.fill).toBe('#vlookup-unsupported');
  });

  it('formula rules evaluate approximate VLOOKUP and HLOOKUP over sorted lookup axes', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 5);
    s = seedNumber(s, 0, 1, 50);
    s = seedNumber(s, 1, 0, 10);
    s = seedNumber(s, 1, 1, 100);
    s = seedNumber(s, 2, 0, 20);
    s = seedNumber(s, 2, 1, 200);
    s = seedNumber(s, 3, 0, 40);
    s = seedNumber(s, 3, 1, 400);
    s = seedCell(s, 5, 0, { kind: 'text', value: 'East' });
    s = seedCell(s, 5, 1, { kind: 'text', value: 'North' });
    s = seedCell(s, 5, 2, { kind: 'text', value: 'South' });
    s = seedNumber(s, 6, 0, 4);
    s = seedNumber(s, 6, 1, 12);
    s = seedNumber(s, 6, 2, 8);
    s = seedNumber(s, 0, 5, 10);
    s = seedNumber(s, 0, 6, 100);
    s = seedNumber(s, 1, 5, 5);
    s = seedNumber(s, 1, 6, 50);
    s = seedNumber(s, 2, 5, 20);
    s = seedNumber(s, 2, 6, 200);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=VLOOKUP(17,$A$1:$B$4,2,TRUE)=100',
            apply: { fill: '#vlookup-approx' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 3, r1: 1, c1: 3 },
            formula: '=HLOOKUP("Nor",$A$6:$C$7,2,TRUE)=4',
            apply: { fill: '#hlookup-approx-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 3, r1: 2, c1: 3 },
            formula: '=ISNA(VLOOKUP(17,$F$1:$G$3,2,TRUE))',
            apply: { fill: '#vlookup-approx-unsorted' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=VLOOKUP(17,$A$1:$B$4,2)=100',
            apply: { fill: '#vlookup-omitted-approx' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 4, r1: 1, c1: 4 },
            formula: '=HLOOKUP("Nor",$A$6:$C$7,2)=4',
            apply: { fill: '#hlookup-omitted-approx' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#vlookup-approx');
    expect(overlay.get('0:1:3')?.fill).toBe('#hlookup-approx-text');
    expect(overlay.get('0:2:3')?.fill).toBe('#vlookup-approx-unsorted');
    expect(overlay.get('0:0:4')?.fill).toBe('#vlookup-omitted-approx');
    expect(overlay.get('0:1:4')?.fill).toBe('#hlookup-omitted-approx');
  });

  it('formula rules evaluate XLOOKUP exact, wildcard, fallback, and reverse search', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'East' });
    s = seedCell(s, 3, 0, { kind: 'text', value: 'North' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 8);
    s = seedNumber(s, 2, 1, 4);
    s = seedNumber(s, 3, 1, 14);
    s = seedCell(s, 5, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 5, 1, { kind: 'text', value: 'South' });
    s = seedCell(s, 5, 2, { kind: 'text', value: 'East' });
    s = seedNumber(s, 6, 0, 12);
    s = seedNumber(s, 6, 1, 8);
    s = seedNumber(s, 6, 2, 4);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=XLOOKUP("South",$A$1:$A$4,$B$1:$B$4)=8',
            apply: { fill: '#xlookup-exact' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=XLOOKUP("No*",$A$1:$A$4,$B$1:$B$4,0,2)=12',
            apply: { fill: '#xlookup-wildcard' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=XLOOKUP("West",$A$1:$A$4,$B$1:$B$4,99)=99',
            apply: { fill: '#xlookup-fallback' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=XLOOKUP("North",$A$1:$A$4,$B$1:$B$4,0,0,-1)=14',
            apply: { fill: '#xlookup-reverse' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=XLOOKUP("East",$A$6:$C$6,$A$7:$C$7)=4',
            apply: { fill: '#xlookup-horizontal' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=ISNA(XLOOKUP("South",$A$1:$A$4,$B$1:$B$4,0,0,2))',
            apply: { fill: '#xlookup-unsupported' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=XLOOKUP("North",$A$1:$A$4,$B$1:$B$4,,0,-1)=14',
            apply: { fill: '#xlookup-omitted-if-not-found' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#xlookup-exact');
    expect(overlay.get('0:0:4')?.fill).toBe('#xlookup-wildcard');
    expect(overlay.get('0:0:5')?.fill).toBe('#xlookup-fallback');
    expect(overlay.get('0:0:6')?.fill).toBe('#xlookup-reverse');
    expect(overlay.get('0:0:7')?.fill).toBe('#xlookup-horizontal');
    expect(overlay.get('0:0:8')?.fill).toBe('#xlookup-unsupported');
    expect(overlay.get('0:0:9')?.fill).toBe('#xlookup-omitted-if-not-found');
  });

  it('formula rules evaluate XLOOKUP next smaller and next larger over monotonic ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 5);
    s = seedNumber(s, 1, 0, 10);
    s = seedNumber(s, 2, 0, 20);
    s = seedNumber(s, 3, 0, 40);
    s = seedNumber(s, 0, 1, 50);
    s = seedNumber(s, 1, 1, 100);
    s = seedNumber(s, 2, 1, 200);
    s = seedNumber(s, 3, 1, 400);
    s = seedCell(s, 5, 0, { kind: 'text', value: 'East' });
    s = seedCell(s, 5, 1, { kind: 'text', value: 'North' });
    s = seedCell(s, 5, 2, { kind: 'text', value: 'South' });
    s = seedNumber(s, 6, 0, 4);
    s = seedNumber(s, 6, 1, 12);
    s = seedNumber(s, 6, 2, 8);
    s = seedNumber(s, 0, 5, 10);
    s = seedNumber(s, 1, 5, 5);
    s = seedNumber(s, 2, 5, 20);
    s = seedNumber(s, 0, 6, 100);
    s = seedNumber(s, 1, 6, 50);
    s = seedNumber(s, 2, 6, 200);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=XLOOKUP(17,$A$1:$A$4,$B$1:$B$4,0,-1)=100',
            apply: { fill: '#xlookup-next-smaller' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=XLOOKUP(17,$A$1:$A$4,$B$1:$B$4,0,1)=200',
            apply: { fill: '#xlookup-next-larger' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
            formula: '=XLOOKUP("Nor",$A$6:$C$6,$A$7:$C$7,0,-1)=4',
            apply: { fill: '#xlookup-text-next-smaller' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 3, c0: 2, r1: 3, c1: 2 },
            formula: '=XLOOKUP("Nor",$A$6:$C$6,$A$7:$C$7,0,1)=12',
            apply: { fill: '#xlookup-text-next-larger' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 4, c0: 2, r1: 4, c1: 2 },
            formula: '=XLOOKUP(17,$F$1:$F$3,$G$1:$G$3,"missing",-1)="missing"',
            apply: { fill: '#xlookup-next-smaller-unsorted' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#xlookup-next-smaller');
    expect(overlay.get('0:1:2')?.fill).toBe('#xlookup-next-larger');
    expect(overlay.get('0:2:2')?.fill).toBe('#xlookup-text-next-smaller');
    expect(overlay.get('0:3:2')?.fill).toBe('#xlookup-text-next-larger');
    expect(overlay.get('0:4:2')?.fill).toBe('#xlookup-next-smaller-unsorted');
  });

  it('formula rules evaluate vector LOOKUP over monotonic ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 5);
    s = seedNumber(s, 1, 0, 10);
    s = seedNumber(s, 2, 0, 20);
    s = seedNumber(s, 3, 0, 40);
    s = seedNumber(s, 0, 1, 50);
    s = seedNumber(s, 1, 1, 100);
    s = seedNumber(s, 2, 1, 200);
    s = seedNumber(s, 3, 1, 400);
    s = seedCell(s, 5, 0, { kind: 'text', value: 'East' });
    s = seedCell(s, 5, 1, { kind: 'text', value: 'North' });
    s = seedCell(s, 5, 2, { kind: 'text', value: 'South' });
    s = seedNumber(s, 6, 0, 4);
    s = seedNumber(s, 6, 1, 12);
    s = seedNumber(s, 6, 2, 8);
    s = seedNumber(s, 0, 5, 10);
    s = seedNumber(s, 1, 5, 5);
    s = seedNumber(s, 2, 5, 20);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=LOOKUP(17,$A$1:$A$4,$B$1:$B$4)=100',
            apply: { fill: '#lookup-vector' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=LOOKUP(17,$A$1:$A$4)=10',
            apply: { fill: '#lookup-omitted-result' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
            formula: '=LOOKUP("Nor",$A$6:$C$6,$A$7:$C$7)=4',
            apply: { fill: '#lookup-horizontal-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 3, c0: 2, r1: 3, c1: 2 },
            formula: '=LOOKUP(1,$A$1:$A$4,$B$1:$B$4)=50',
            apply: { fill: '#lookup-too-small' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 4, c0: 2, r1: 4, c1: 2 },
            formula: '=LOOKUP(17,$F$1:$F$3,$B$1:$B$3)=100',
            apply: { fill: '#lookup-unsorted' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 5, c0: 2, r1: 5, c1: 2 },
            formula: '=LOOKUP(17,$A$1:$A$4,$B$1:$B$3)=100',
            apply: { fill: '#lookup-mismatch' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#lookup-vector');
    expect(overlay.get('0:1:2')?.fill).toBe('#lookup-omitted-result');
    expect(overlay.get('0:2:2')?.fill).toBe('#lookup-horizontal-text');
    expect(overlay.get('0:3:2')?.fill).toBeUndefined();
    expect(overlay.get('0:4:2')?.fill).toBeUndefined();
    expect(overlay.get('0:5:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate XLOOKUP and LOOKUP over dynamic ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'East' });
    s = seedCell(s, 3, 0, { kind: 'text', value: 'North' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 8);
    s = seedNumber(s, 2, 1, 4);
    s = seedNumber(s, 3, 1, 14);
    s = seedNumber(s, 0, 4, 5);
    s = seedNumber(s, 1, 4, 10);
    s = seedNumber(s, 2, 4, 20);
    s = seedNumber(s, 3, 4, 40);
    s = seedNumber(s, 0, 5, 50);
    s = seedNumber(s, 1, 5, 100);
    s = seedNumber(s, 2, 5, 200);
    s = seedNumber(s, 3, 5, 400);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=XLOOKUP("South",OFFSET(A1,0,0,4,1),OFFSET(B1,0,0,4,1))=8',
            apply: { fill: '#xlookup-offset' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=XLOOKUP("North",INDIRECT("A1:A4"),INDIRECT("B1:B4"),0,0,-1)=14',
            apply: { fill: '#xlookup-indirect-reverse' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=XLOOKUP(17,OFFSET(E1,0,0,4,1),INDIRECT("F1:F4"),0,-1)=100',
            apply: { fill: '#xlookup-dynamic-approx' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=LOOKUP(17,OFFSET(E1,0,0,4,1),OFFSET(F1,0,0,4,1))=100',
            apply: { fill: '#lookup-offset' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=LOOKUP(17,INDIRECT("E1:E4"))=10',
            apply: { fill: '#lookup-indirect-self' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=XLOOKUP("South",OFFSET(A1,0,0,4,1),OFFSET(B1,0,0,2,1),0)=8',
            apply: { fill: '#xlookup-dynamic-size-mismatch' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#xlookup-offset');
    expect(overlay.get('0:0:3')?.fill).toBe('#xlookup-indirect-reverse');
    expect(overlay.get('0:0:6')?.fill).toBe('#xlookup-dynamic-approx');
    expect(overlay.get('0:0:7')?.fill).toBe('#lookup-offset');
    expect(overlay.get('0:0:8')?.fill).toBe('#lookup-indirect-self');
    expect(overlay.get('0:0:9')?.fill).toBeUndefined();
  });

  it('formula rules evaluate SEARCH/FIND/LEN text operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North Region' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'south region' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'NE' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 2, c1: 1 },
            formula: '=AND(ISNUMBER(SEARCH("region",A1)),LEN(A1)>5)',
            apply: { fill: '#search' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 2, c1: 2 },
            formula: '=ISNUMBER(FIND("Region",A1))',
            apply: { fill: '#find' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#search');
    expect(overlay.get('0:1:1')?.fill).toBe('#search');
    expect(overlay.get('0:2:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBe('#find');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
  });

  it('formula SEARCH start position is one-based and fails closed when missing', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'north north' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'north south' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 1, c1: 1 },
            formula: '=SEARCH("north",A1,7)=7',
            apply: { fill: '#search-start' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=SEARCH("north",A1,)=1',
            apply: { fill: '#search-omitted-start' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#search-start');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBe('#search-omitted-start');
  });

  it('formula SEARCH supports Excel wildcards while FIND treats them literally', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North Region' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'Q1*Plan' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(SEARCH("N*r",A1)=1,SEARCH("r?g",A1)=7)',
            apply: { fill: '#search-wildcard' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=SEARCH("~*",A2)=3',
            apply: { fill: '#search-escape' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=FIND("*",A2)=3',
            apply: { fill: '#find-literal' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=FIND("Region",A1,)=7',
            apply: { fill: '#find-omitted-start' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#search-wildcard');
    expect(overlay.get('0:1:1')?.fill).toBe('#search-escape');
    expect(overlay.get('0:1:2')?.fill).toBe('#find-literal');
    expect(overlay.get('0:0:3')?.fill).toBe('#find-omitted-start');
  });

  it('formula rules evaluate scalar HYPERLINK display values', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            formula:
              '=AND(HYPERLINK("https://example.test","Example")="Example",HYPERLINK("#A1")="#A1",HYPERLINK("#A1",42)=42)',
            apply: { fill: '#hyperlink' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=HYPERLINK(NA())=""',
            apply: { fill: '#hyperlink-invalid' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#hyperlink');
    expect(overlay.get('0:0:1')?.fill).toBeUndefined();
  });

  it('formula rules evaluate limited CELL information operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedNumber(s, 1, 0, 42);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula:
              '=AND(CELL("address",A1)="$A$1",CELL("row",A1)=1,CELL("col",A1)=1,CELL("contents",A1)="North",CELL("type",A1)="l")',
            apply: { fill: '#cell-info-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=AND(CELL("contents",A2)=42,CELL("type",A2)="v")',
            apply: { fill: '#cell-info-number' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
            formula: '=AND(CELL("address")="$B$3",CELL("type")="b")',
            apply: { fill: '#cell-info-current' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 3, c0: 1, r1: 3, c1: 1 },
            formula: '=CELL("filename",A1)=""',
            apply: { fill: '#cell-info-unsupported' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 4, c0: 1, r1: 4, c1: 1 },
            formula:
              '=AND(CELL("address",OFFSET(A1,1,0))="$A$2",CELL("contents",INDIRECT("A2"))=42)',
            apply: { fill: '#cell-info-dynamic' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 5, c0: 1, r1: 5, c1: 1 },
            formula: '=CELL("address",OFFSET(A1,0,0,2,1))="$A$1"',
            apply: { fill: '#cell-info-dynamic-multi' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#cell-info-text');
    expect(overlay.get('0:1:1')?.fill).toBe('#cell-info-number');
    expect(overlay.get('0:2:1')?.fill).toBe('#cell-info-current');
    expect(overlay.get('0:3:1')?.fill).toBeUndefined();
    expect(overlay.get('0:4:1')?.fill).toBe('#cell-info-dynamic');
    expect(overlay.get('0:5:1')?.fill).toBeUndefined();
  });

  it('formula rules evaluate limited SHEET and SHEETS information operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            formula: '=AND(SHEET()=1,SHEET(A1)=1,SHEET(A1:B2)=1)',
            apply: { fill: '#sheet-info' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=SHEETS(A1:B2)=1',
            apply: { fill: '#sheets-info' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=SHEETS()=1',
            apply: { fill: '#sheets-workbook-count' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=SHEET(Sheet2!A1)=2',
            apply: { fill: '#sheet-other' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#sheet-info');
    expect(overlay.get('0:0:1')?.fill).toBe('#sheets-info');
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
  });

  it('formula rules evaluate limited SHEET and SHEETS over dynamic ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            formula: '=AND(SHEET(OFFSET(A1,0,0,2,2))=1,SHEETS(INDIRECT("A1:B2"))=1)',
            apply: { fill: '#sheet-info-dynamic' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=SHEET(INDIRECT("Sheet2!A1"))=2',
            apply: { fill: '#sheet-info-dynamic-other' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#sheet-info-dynamic');
    expect(overlay.get('0:0:1')?.fill).toBeUndefined();
  });

  it('formula rules evaluate ISLOGICAL/ISNONTEXT/ISFORMULA predicates', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'bool', value: true });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'text' });
    const cells = new Map(s.data.cells);
    cells.set('0:2:0', { value: { kind: 'number', value: 3 }, formula: '=SUM(1,2)' });
    s = {
      ...s,
      data: { ...s.data, cells },
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 2, c1: 1 },
            formula: '=ISLOGICAL(A1)',
            apply: { fill: '#logical' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 2, c1: 2 },
            formula: '=ISNONTEXT(A1)',
            apply: { fill: '#nontext' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 2, c1: 3 },
            formula: '=ISFORMULA(A1)',
            apply: { fill: '#formula' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 2, c1: 4 },
            formula: '=EXACT(FORMULATEXT(A1),"=SUM(1,2)")',
            apply: { fill: '#formulatext' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=AND(TYPE(A1)=4,TYPE(A2)=2,TYPE(A3)=1,TYPE(NA())=16)',
            apply: { fill: '#type-codes' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 5, r1: 1, c1: 5 },
            formula: '=EXACT(FORMULATEXT(OFFSET(A1,2,0)),"=SUM(1,2)")',
            apply: { fill: '#formulatext-dynamic' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 5, r1: 2, c1: 5 },
            formula: '=FORMULATEXT(OFFSET(A1,0,0,2,1))=""',
            apply: { fill: '#formulatext-dynamic-multi' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=ISFORMULA(1)',
            apply: { fill: '#formula-literal' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=ISFORMULA(OFFSET($A$1,2,0))',
            apply: { fill: '#formula-dynamic-offset' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=ISFORMULA(INDIRECT("$A$3"))',
            apply: { fill: '#formula-dynamic-indirect' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=ISFORMULA(OFFSET($A$1,0,0,2,1))',
            apply: { fill: '#formula-dynamic-multi' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#logical');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
    expect(overlay.get('0:2:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBe('#nontext');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
    expect(overlay.get('0:2:2')?.fill).toBe('#nontext');
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
    expect(overlay.get('0:1:3')?.fill).toBeUndefined();
    expect(overlay.get('0:2:3')?.fill).toBe('#formula');
    expect(overlay.get('0:0:4')?.fill).toBeUndefined();
    expect(overlay.get('0:1:4')?.fill).toBeUndefined();
    expect(overlay.get('0:2:4')?.fill).toBe('#formulatext');
    expect(overlay.get('0:0:5')?.fill).toBe('#type-codes');
    expect(overlay.get('0:1:5')?.fill).toBe('#formulatext-dynamic');
    expect(overlay.get('0:2:5')?.fill).toBeUndefined();
    expect(overlay.get('0:0:6')?.fill).toBeUndefined();
    expect(overlay.get('0:0:7')?.fill).toBe('#formula-dynamic-offset');
    expect(overlay.get('0:0:8')?.fill).toBe('#formula-dynamic-indirect');
    expect(overlay.get('0:0:9')?.fill).toBeUndefined();
  });

  it('formula rules evaluate ISREF and NA error operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'error', code: 6, text: '#N/A' });
    s = seedCell(s, 1, 0, { kind: 'error', code: 1, text: '#DIV/0!' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(ISREF(A1),ISREF(A1:B2),NOT(ISREF(1)))',
            apply: { fill: '#isref' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=ISERROR(NA())',
            apply: { fill: '#na-error' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=NA()=A1',
            apply: { fill: '#na-equals' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 1, c1: 4 },
            formula: '=ISNA(A1)',
            apply: { fill: '#isna' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 1, c1: 5 },
            formula: '=ISERR(A1)',
            apply: { fill: '#iserr' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula:
              '=AND(ERROR.TYPE(A1)=7,ERROR.TYPE(A2)=2,ERROR.TYPE(1/0)=2,ERROR.TYPE(SQRT(-1))=6)',
            apply: { fill: '#error-type' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=ERROR.TYPE(1)=0',
            apply: { fill: '#error-type-non-error' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=AND(ISREF(OFFSET(A1,0,0,1,1)),ISREF(INDIRECT("A1:B2")),NOT(ISREF(1)))',
            apply: { fill: '#isref-dynamic' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#isref');
    expect(overlay.get('0:0:2')?.fill).toBe('#na-error');
    expect(overlay.get('0:0:3')?.fill).toBe('#na-equals');
    expect(overlay.get('0:0:4')?.fill).toBe('#isna');
    expect(overlay.get('0:1:4')?.fill).toBeUndefined();
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
    expect(overlay.get('0:1:5')?.fill).toBe('#iserr');
    expect(overlay.get('0:0:6')?.fill).toBe('#error-type');
    expect(overlay.get('0:0:7')?.fill).toBeUndefined();
    expect(overlay.get('0:0:8')?.fill).toBe('#isref-dynamic');
  });

  it('formula rules evaluate IFERROR and IFNA fallback operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North Region' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South Area' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 1, c1: 1 },
            formula: '=IFERROR(SEARCH("region",A1),0)>0',
            apply: { fill: '#iferror-search' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=IFNA(NA(),42)=42',
            apply: { fill: '#ifna-na' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=ISERR(IFNA(1/0,42))',
            apply: { fill: '#ifna-preserves-other-errors' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#iferror-search');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBe('#ifna-na');
    expect(overlay.get('0:0:3')?.fill).toBe('#ifna-preserves-other-errors');
  });

  it('formula rules evaluate LEFT/RIGHT/MID text operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North Region' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South Region' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'North Area' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 2, c1: 1 },
            formula: '=LEFT(A1,5)="North"',
            apply: { fill: '#left' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 2, c1: 2 },
            formula: '=RIGHT(A1,6)="Region"',
            apply: { fill: '#right' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 2, c1: 3 },
            formula: '=MID(A1,7,6)="Region"',
            apply: { fill: '#mid' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#left');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
    expect(overlay.get('0:2:1')?.fill).toBe('#left');
    expect(overlay.get('0:0:2')?.fill).toBe('#right');
    expect(overlay.get('0:1:2')?.fill).toBe('#right');
    expect(overlay.get('0:2:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBe('#mid');
    expect(overlay.get('0:1:3')?.fill).toBe('#mid');
    expect(overlay.get('0:2:3')?.fill).toBeUndefined();
  });

  it('formula text slices default to one character and fail closed on invalid arguments', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=LEFT(A1)="N"',
            apply: { fill: '#left-default' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=AND(LEFT(A1,)="N",RIGHT(A1,)="h")',
            apply: { fill: '#slice-omitted-count' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=RIGHT(A1,-1)=""',
            apply: { fill: '#right-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=MID(A1,0,1)="N"',
            apply: { fill: '#mid-invalid' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#left-default');
    expect(overlay.get('0:0:4')?.fill).toBe('#slice-omitted-count');
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
  });

  it('formula rules evaluate LOWER/UPPER/TRIM text transforms', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: '  North   Region  ' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South Region' });
    s = seedCell(s, 2, 0, { kind: 'text', value: ' 1,234.5 ' });
    s = seedCell(s, 3, 0, { kind: 'text', value: '12.5%' });
    s = seedCell(s, 4, 0, { kind: 'text', value: 'north' });
    s = seedCell(s, 5, 0, { kind: 'text', value: '1.234,5' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 1, c1: 1 },
            formula: '=LOWER(TRIM(A1))="north region"',
            apply: { fill: '#lower-trim' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 1, c1: 2 },
            formula: '=UPPER(LEFT(TRIM(A1),5))="NORTH"',
            apply: { fill: '#upper-left' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
            formula: '=VALUE(A3)=1234.5',
            apply: { fill: '#value-thousands' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 3, c0: 1, r1: 3, c1: 1 },
            formula: '=VALUE(A4)=0.125',
            apply: { fill: '#value-percent' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 4, c0: 1, r1: 4, c1: 1 },
            formula: '=VALUE(A5)=0',
            apply: { fill: '#value-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 5, c0: 1, r1: 5, c1: 1 },
            formula: '=NUMBERVALUE(A6,",",".")=1234.5',
            apply: { fill: '#numbervalue-separators' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
            formula: '=NUMBERVALUE(A3,,)=1234.5',
            apply: { fill: '#numbervalue-omitted-separators' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 5, c0: 2, r1: 5, c1: 2 },
            formula: '=NUMBERVALUE(A6,",",",")=1234.5',
            apply: { fill: '#numbervalue-invalid-separators' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#lower-trim');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBe('#upper-left');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
    expect(overlay.get('0:2:1')?.fill).toBe('#value-thousands');
    expect(overlay.get('0:3:1')?.fill).toBe('#value-percent');
    expect(overlay.get('0:4:1')?.fill).toBeUndefined();
    expect(overlay.get('0:5:1')?.fill).toBe('#numbervalue-separators');
    expect(overlay.get('0:2:2')?.fill).toBe('#numbervalue-omitted-separators');
    expect(overlay.get('0:5:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate text concatenation with ampersand', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North Region' });
    s = seedNumber(s, 0, 1, 12);
    s = seedCell(s, 0, 2, { kind: 'bool', value: true });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=LEFT(A1,5)&"-"&RIGHT(A1,6)="North-Region"',
            apply: { fill: '#concat-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '="N"&"o"&"r"&"t"&"h"="North"',
            apply: { fill: '#concat-left-assoc' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=A1&" "&B1&" "&C1="North Region 12 TRUE"',
            apply: { fill: '#concat-coerce' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=D10&"x"="x"',
            apply: { fill: '#concat-blank' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#concat-text');
    expect(overlay.get('0:0:4')?.fill).toBe('#concat-left-assoc');
    expect(overlay.get('0:0:5')?.fill).toBe('#concat-coerce');
    expect(overlay.get('0:0:6')?.fill).toBe('#concat-blank');
  });

  it('formula rules evaluate scalar CONCATENATE and CONCAT text functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedNumber(s, 0, 1, 12);
    s = seedCell(s, 0, 2, { kind: 'bool', value: true });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=CONCATENATE(A1,"-",B1,"-",C1)="North-12-TRUE"',
            apply: { fill: '#concatenate' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=CONCAT(A1,"-",D10)="North-"',
            apply: { fill: '#concat-function' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=CONCATENATE(A1,,"-",B1)="North-12"',
            apply: { fill: '#concatenate-omitted' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=CONCAT(A1,,D10)="North"',
            apply: { fill: '#concat-omitted' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#concatenate');
    expect(overlay.get('0:0:4')?.fill).toBe('#concat-function');
    expect(overlay.get('0:0:5')?.fill).toBe('#concatenate-omitted');
    expect(overlay.get('0:0:6')?.fill).toBe('#concat-omitted');
  });

  it('formula rules evaluate scalar SUBSTITUTE/REPLACE/REPT/TEXTJOIN text functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North Region North' });
    s = seedCell(s, 0, 1, { kind: 'text', value: 'East' });
    s = seedNumber(s, 0, 2, 12);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=SUBSTITUTE(A1,"North","South",2)="North Region South"',
            apply: { fill: '#substitute-instance' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=REPLACE(A1,7,6,"Area")="North Area North"',
            apply: { fill: '#replace' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=REPT(LEFT(B1,1),3)="EEE"',
            apply: { fill: '#rept' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=TEXTJOIN("-",TRUE,B1,D10,C1)="East-12"',
            apply: { fill: '#textjoin-ignore-empty' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=TEXTJOIN("-",FALSE,B1,D10,C1)="East--12"',
            apply: { fill: '#textjoin-keep-empty' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=SUBSTITUTE(A1,"North","South",)="South Region South"',
            apply: { fill: '#substitute-omitted-instance' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=TEXTJOIN("-",TRUE,B1,,C1)="East-12"',
            apply: { fill: '#textjoin-omitted-ignore-empty' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 10, r1: 0, c1: 10 },
            formula: '=TEXTJOIN("-",FALSE,B1,,C1)="East--12"',
            apply: { fill: '#textjoin-omitted-keep-empty' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#substitute-instance');
    expect(overlay.get('0:0:4')?.fill).toBe('#replace');
    expect(overlay.get('0:0:5')?.fill).toBe('#rept');
    expect(overlay.get('0:0:6')?.fill).toBe('#textjoin-ignore-empty');
    expect(overlay.get('0:0:7')?.fill).toBe('#textjoin-keep-empty');
    expect(overlay.get('0:0:8')?.fill).toBe('#substitute-omitted-instance');
    expect(overlay.get('0:0:9')?.fill).toBe('#textjoin-omitted-ignore-empty');
    expect(overlay.get('0:0:10')?.fill).toBe('#textjoin-omitted-keep-empty');
  });

  it('formula rules evaluate scalar TEXTBEFORE and TEXTAFTER text functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North-East-West' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(TEXTBEFORE(A1,"-")="North",TEXTAFTER(A1,"-")="East-West")',
            apply: { fill: '#text-before-after' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=AND(TEXTBEFORE(A1,"-",-1)="North-East",TEXTAFTER(A1,"-",-1)="West")',
            apply: { fill: '#text-before-after-negative' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=AND(TEXTBEFORE(A1,"east",1,1)="North-",TEXTAFTER(A1,"east",1,1)="-West")',
            apply: { fill: '#text-before-after-ignore-case' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=TEXTAFTER(A1,"/",1,0,0,"missing")="missing"',
            apply: { fill: '#textafter-fallback' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#text-before-after');
    expect(overlay.get('0:0:2')?.fill).toBe('#text-before-after-negative');
    expect(overlay.get('0:0:3')?.fill).toBe('#text-before-after-ignore-case');
    expect(overlay.get('0:0:4')?.fill).toBe('#textafter-fallback');
  });

  it('formula rules evaluate scalar TEXT number formatting function', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 1234.567);
    s = seedNumber(s, 0, 1, 0.25);
    s = seedNumber(s, 0, 2, 45651);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=AND(TEXT(A1,"#,##0.00")="1,234.57",TEXT(B1,"0%")="25%")',
            apply: { fill: '#text-number-format' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=TEXT(C1,"yyyy-mm-dd")="2024-12-25"',
            apply: { fill: '#text-date-format' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=TEXT("North","0")="North"',
            apply: { fill: '#text-format-nonnumeric' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#text-number-format');
    expect(overlay.get('0:0:4')?.fill).toBe('#text-date-format');
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
  });

  it('formula rules evaluate scalar DOLLAR and FIXED text formatting functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 1234.567);
    s = seedNumber(s, 0, 1, -1234.567);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=AND(DOLLAR(A1)="$1,234.57",DOLLAR(A1,0)="$1,235")',
            apply: { fill: '#dollar' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=AND(FIXED(A1,1)="1,234.6",FIXED(A1,1,TRUE())="1234.6")',
            apply: { fill: '#fixed' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=AND(DOLLAR(A1,-2)="$1,200",FIXED(B1,-2)="-1,200")',
            apply: { fill: '#fixed-negative-decimals' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=DOLLAR("North")="$0.00"',
            apply: { fill: '#dollar-nonnumeric' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#dollar');
    expect(overlay.get('0:0:3')?.fill).toBe('#fixed');
    expect(overlay.get('0:0:4')?.fill).toBe('#fixed-negative-decimals');
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
  });

  it('formula rules evaluate scalar VALUETOTEXT text conversion', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 12);
    s = seedCell(s, 0, 1, { kind: 'bool', value: true });
    s = seedCell(s, 0, 2, { kind: 'text', value: 'North' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula:
              '=AND(VALUETOTEXT(A1)="12",VALUETOTEXT(B1)="TRUE",LEN(VALUETOTEXT(C1,1))=7,FIND("North",VALUETOTEXT(C1,1))=2,VALUETOTEXT(NA())="#N/A")',
            apply: { fill: '#value-to-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=VALUETOTEXT(A1,2)="12"',
            apply: { fill: '#value-to-text-invalid' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#value-to-text');
    expect(overlay.get('0:0:4')?.fill).toBeUndefined();
  });

  it('formula scalar text functions fail closed on invalid arguments', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=SUBSTITUTE(A1,"o","0",0)="N0rth"',
            apply: { fill: '#substitute-invalid-instance' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=REPLACE(A1,0,1,"S")="Sorth"',
            apply: { fill: '#replace-invalid-start' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=REPT(A1,-1)=""',
            apply: { fill: '#rept-invalid-count' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=TEXTJOIN("-",A1,A1)="North"',
            apply: { fill: '#textjoin-invalid-ignore-empty' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=TEXTBEFORE(A1,"",1)="North"',
            apply: { fill: '#textbefore-empty-delimiter' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=TEXTAFTER(A1,"x")="North"',
            apply: { fill: '#textafter-missing' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=TEXT(12,"")="12"',
            apply: { fill: '#text-empty-pattern' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
    expect(overlay.get('0:0:4')?.fill).toBeUndefined();
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
    expect(overlay.get('0:0:6')?.fill).toBeUndefined();
    expect(overlay.get('0:0:7')?.fill).toBeUndefined();
  });

  it('formula rules evaluate EXACT as a case-sensitive boolean operand', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'north' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 1, c1: 1 },
            formula: '=EXACT(A1,"North")',
            apply: { fill: '#exact' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 1, c1: 2 },
            formula: '=NOT(EXACT(A1,"North"))',
            apply: { fill: '#not-exact' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#exact');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:1:2')?.fill).toBe('#not-exact');
  });

  it('formula rules evaluate N/T scalar coercion operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 12);
    s = seedCell(s, 0, 1, { kind: 'bool', value: true });
    s = seedCell(s, 0, 2, { kind: 'text', value: 'North' });
    s = seedCell(s, 0, 4, { kind: 'error', code: 6, text: '#N/A' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=AND(N(A1)=12,N(B1)=1,N(C1)=0,N(D1)=0)',
            apply: { fill: '#n-coerce' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=AND(T(C1)="North",T(A1)="",T(D1)="")',
            apply: { fill: '#t-coerce' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=ISNA(N(E1))',
            apply: { fill: '#n-error' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:5')?.fill).toBe('#n-coerce');
    expect(overlay.get('0:0:6')?.fill).toBe('#t-coerce');
    expect(overlay.get('0:0:7')?.fill).toBe('#n-error');
  });

  it('formula rules evaluate ADDRESS text references', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            formula:
              '=AND(ADDRESS(3,2)="$B$3",ADDRESS(3,2,4)="B3",ADDRESS(3,2,2,FALSE())="R3C[2]",ADDRESS(3,2,,,"Sheet 1")="\'Sheet 1\'!$B$3")',
            apply: { fill: '#address' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=ADDRESS(0,1)=""',
            apply: { fill: '#invalid-address' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#address');
    expect(overlay.get('0:0:1')?.fill).toBeUndefined();
  });

  it('formula text comparisons are case-insensitive unless EXACT is used', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'south' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 1, c1: 1 },
            formula: '=A1="north"',
            apply: { fill: '#case-insensitive-eq' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 1, c1: 2 },
            formula: '=A1<>"NORTH"',
            apply: { fill: '#case-insensitive-ne' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=EXACT(A1,"north")',
            apply: { fill: '#exact-case' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#case-insensitive-eq');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:1:2')?.fill).toBe('#case-insensitive-ne');
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
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

  it('formula rules evaluate ROW/COLUMN position operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 },
            formula: '=MOD(ROW(),2)=0',
            apply: { fill: '#even-row' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 3 },
            formula: '=COLUMN()=3',
            apply: { fill: '#column-c' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 1, c1: 4 },
            formula: '=ROW(A1)=2',
            apply: { fill: '#relative-row-ref' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=AND(ROWS(A1:C5)=5,COLUMNS(A1:C5)=3,AREAS(A1:C5)=1,AREAS(A1)=1)',
            apply: { fill: '#range-dimensions' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 5, r1: 1, c1: 5 },
            formula:
              '=AND(ROWS(OFFSET(A1,0,0,5,3))=5,COLUMNS(INDIRECT("A1:C5"))=3,AREAS(OFFSET(A1,0,0,5,3))=1)',
            apply: { fill: '#range-dimensions-dynamic' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 5, r1: 2, c1: 5 },
            formula: '=AND(ROW(OFFSET(A1,4,0))=5,COLUMN(INDIRECT("C1"))=3)',
            apply: { fill: '#position-dynamic' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 3, c0: 5, r1: 3, c1: 5 },
            formula: '=ROW(OFFSET(A1,0,0,2,1))=1',
            apply: { fill: '#position-dynamic-multi' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBeUndefined();
    expect(overlay.get('0:1:0')?.fill).toBe('#even-row');
    expect(overlay.get('0:2:0')?.fill).toBeUndefined();
    expect(overlay.get('0:3:0')?.fill).toBe('#even-row');
    expect(overlay.get('0:0:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBe('#column-c');
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
    expect(overlay.get('0:0:4')?.fill).toBeUndefined();
    expect(overlay.get('0:1:4')?.fill).toBe('#relative-row-ref');
    expect(overlay.get('0:0:5')?.fill).toBe('#range-dimensions');
    expect(overlay.get('0:1:5')?.fill).toBe('#range-dimensions-dynamic');
    expect(overlay.get('0:2:5')?.fill).toBe('#position-dynamic');
    expect(overlay.get('0:3:5')?.fill).toBeUndefined();
  });

  it('formula rules evaluate ISEVEN/ISODD boolean operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 },
            formula: '=ISEVEN(ROW())',
            apply: { fill: '#is-even-row' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 },
            formula: '=ISODD(COLUMN())',
            apply: { fill: '#is-odd-column' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#is-odd-column');
    expect(overlay.get('0:1:0')?.fill).toBe('#is-even-row');
    expect(overlay.get('0:2:0')?.fill).toBeUndefined();
    expect(overlay.get('0:0:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBe('#is-odd-column');
  });

  it('formula rules evaluate ABS/MOD/ROUND numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, -3.4);
    s = seedNumber(s, 1, 0, 125);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(ABS(A1)>3,MOD(A1,2)>0.5,ROUND(A1,0)=-3)',
            apply: { fill: '#numeric-fns' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=ROUND(A2,-1)=130',
            apply: { fill: '#round-negative-digits' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=MOD(A1,0)=0',
            apply: { fill: '#mod-zero' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#numeric-fns');
    expect(overlay.get('0:1:1')?.fill).toBe('#round-negative-digits');
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate INT/TRUNC/SQRT/POWER/SIGN numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, -3.75);
    s = seedNumber(s, 1, 0, 125.987);
    s = seedNumber(s, 2, 0, 16);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(INT(A1)=-4,TRUNC(A1)=-3,SIGN(A1)=-1)',
            apply: { fill: '#int-trunc-sign' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=AND(TRUNC(A2,1)=125.9,TRUNC(A2,-1)=120)',
            apply: { fill: '#trunc-digits' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 3, r1: 1, c1: 3 },
            formula: '=TRUNC(A2,)=125',
            apply: { fill: '#trunc-omitted-digits' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
            formula: '=AND(SQRT(A3)=4,POWER(A3,2)=256)',
            apply: { fill: '#sqrt-power' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=SQRT(A1)=0',
            apply: { fill: '#sqrt-negative' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#int-trunc-sign');
    expect(overlay.get('0:1:1')?.fill).toBe('#trunc-digits');
    expect(overlay.get('0:1:3')?.fill).toBe('#trunc-omitted-digits');
    expect(overlay.get('0:2:1')?.fill).toBe('#sqrt-power');
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate ROUNDUP/ROUNDDOWN numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, -3.275);
    s = seedNumber(s, 1, 0, 125.987);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(ROUNDUP(A1,2)=-3.28,ROUNDDOWN(A1,2)=-3.27)',
            apply: { fill: '#roundup-rounddown' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=AND(ROUNDUP(A2,-1)=130,ROUNDDOWN(A2,-1)=120)',
            apply: { fill: '#roundup-rounddown-negative-digits' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#roundup-rounddown');
    expect(overlay.get('0:1:1')?.fill).toBe('#roundup-rounddown-negative-digits');
  });

  it('formula rules evaluate EVEN/ODD numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 3.2);
    s = seedNumber(s, 1, 0, -3.2);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(EVEN(A1)=4,ODD(A1)=5)',
            apply: { fill: '#even-odd-positive' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=AND(EVEN(A2)=-4,ODD(A2)=-5)',
            apply: { fill: '#even-odd-negative' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#even-odd-positive');
    expect(overlay.get('0:1:1')?.fill).toBe('#even-odd-negative');
  });

  it('formula rules evaluate CEILING.MATH/FLOOR.MATH numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 4.3);
    s = seedNumber(s, 1, 0, -4.3);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(CEILING.MATH(A1,2)=6,FLOOR.MATH(A1,2)=4)',
            apply: { fill: '#math-round-positive' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=AND(CEILING.MATH(A2,2)=-4,FLOOR.MATH(A2,2)=-6)',
            apply: { fill: '#math-round-negative-default' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=AND(CEILING.MATH(A2,2,1)=-6,FLOOR.MATH(A2,2,1)=-4)',
            apply: { fill: '#math-round-negative-mode' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=AND(CEILING.MATH(A1,)=5,FLOOR.MATH(A1,)=4)',
            apply: { fill: '#math-round-omitted-significance' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 3, r1: 1, c1: 3 },
            formula: '=AND(CEILING.MATH(A2,,1)=-5,FLOOR.MATH(A2,,1)=-4)',
            apply: { fill: '#math-round-omitted-significance-mode' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#math-round-positive');
    expect(overlay.get('0:1:1')?.fill).toBe('#math-round-negative-default');
    expect(overlay.get('0:1:2')?.fill).toBe('#math-round-negative-mode');
    expect(overlay.get('0:0:3')?.fill).toBe('#math-round-omitted-significance');
    expect(overlay.get('0:1:3')?.fill).toBe('#math-round-omitted-significance-mode');
  });

  it('formula rules evaluate CEILING.PRECISE/FLOOR.PRECISE numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 4.3);
    s = seedNumber(s, 1, 0, -4.3);
    s = seedCell(s, 2, 0, { kind: 'text', value: 'north' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(CEILING.PRECISE(A1,2)=6,FLOOR.PRECISE(A1,2)=4)',
            apply: { fill: '#precise-round-positive' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=AND(CEILING.PRECISE(A2,2)=-4,FLOOR.PRECISE(A2,2)=-6)',
            apply: { fill: '#precise-round-negative' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=AND(CEILING.PRECISE(A1,)=5,FLOOR.PRECISE(A1,)=4,ISO.CEILING(A1)=5)',
            apply: { fill: '#precise-round-omitted-significance' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
            formula: '=CEILING.PRECISE(A3)=0',
            apply: { fill: '#precise-round-text' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#precise-round-positive');
    expect(overlay.get('0:1:1')?.fill).toBe('#precise-round-negative');
    expect(overlay.get('0:0:2')?.fill).toBe('#precise-round-omitted-significance');
    expect(overlay.get('0:2:1')?.fill).toBeUndefined();
  });

  it('formula rules evaluate legacy CEILING/FLOOR numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 4.3);
    s = seedNumber(s, 1, 0, -4.3);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(CEILING(A1,2)=6,FLOOR(A1,2)=4)',
            apply: { fill: '#legacy-round-positive' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=AND(CEILING(A2,-2)=-6,FLOOR(A2,-2)=-4)',
            apply: { fill: '#legacy-round-negative' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=CEILING(A2,2)=0',
            apply: { fill: '#legacy-round-sign-mismatch' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#legacy-round-positive');
    expect(overlay.get('0:1:1')?.fill).toBe('#legacy-round-negative');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate MROUND numeric function', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 10);
    s = seedNumber(s, 1, 0, -10);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(MROUND(A1,3)=9,MROUND(10.5,3)=12)',
            apply: { fill: '#mround-positive' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=MROUND(A2,-3)=-9',
            apply: { fill: '#mround-negative' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=MROUND(A2,3)=0',
            apply: { fill: '#mround-sign-mismatch' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#mround-positive');
    expect(overlay.get('0:1:1')?.fill).toBe('#mround-negative');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate GCD/LCM numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 24);
    s = seedNumber(s, 0, 1, 36);
    s = seedNumber(s, 0, 2, 54);
    s = seedNumber(s, 1, 0, -12);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=AND(GCD(A1,B1,C1)=6,LCM(4,6,8)=24)',
            apply: { fill: '#gcd-lcm' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=LCM(A1,0)=0',
            apply: { fill: '#lcm-zero' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=GCD(A2,6)=6',
            apply: { fill: '#gcd-negative' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#gcd-lcm');
    expect(overlay.get('0:0:4')?.fill).toBe('#lcm-zero');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
  });

  it('formula rules evaluate factorial, gamma, and combinatoric numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            formula: '=AND(FACT(5)=120,FACTDOUBLE(6)=48,FACT(5.9)=120)',
            apply: { fill: '#factorial' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula:
              '=AND(COMBIN(10,3)=120,COMBINA(3,2)=6,PERMUT(10,3)=720,PERMUTATIONA(4,3)=64,MULTINOMIAL(2,3,4)=1260)',
            apply: { fill: '#combinatoric' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=AND(ROUND(GAMMALN(4),6)=1.791759,ROUND(GAMMALN.PRECISE(4),6)=1.791759)',
            apply: { fill: '#gammaln' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula:
              '=AND(ROUND(GAMMA(5),6)=24,ROUND(GAMMA(0.5),6)=1.772454,ROUND(GAMMA(-0.5),6)=-3.544908)',
            apply: { fill: '#gamma' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=FACT(-1)=0',
            apply: { fill: '#fact-negative' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=COMBIN(2,3)=0',
            apply: { fill: '#combin-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=GAMMALN(0)=0',
            apply: { fill: '#gammaln-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=MULTINOMIAL(2,-1)=0',
            apply: { fill: '#multinomial-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=GAMMA(0)=0',
            apply: { fill: '#gamma-invalid' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#factorial');
    expect(overlay.get('0:0:1')?.fill).toBe('#combinatoric');
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
    expect(overlay.get('0:0:4')?.fill).toBe('#gammaln');
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
    expect(overlay.get('0:0:6')?.fill).toBeUndefined();
    expect(overlay.get('0:0:7')?.fill).toBe('#gamma');
    expect(overlay.get('0:0:8')?.fill).toBeUndefined();
  });

  it('formula rules evaluate QUOTIENT numeric function', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 17);
    s = seedNumber(s, 1, 0, -17);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=QUOTIENT(A1,5)=3',
            apply: { fill: '#quotient-positive' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=QUOTIENT(A2,5)=-3',
            apply: { fill: '#quotient-negative' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=QUOTIENT(A1,0)=0',
            apply: { fill: '#quotient-zero-divisor' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#quotient-positive');
    expect(overlay.get('0:1:1')?.fill).toBe('#quotient-negative');
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate trigonometric numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 30);
    s = seedNumber(s, 1, 0, 45);
    s = seedCell(s, 2, 0, { kind: 'text', value: 'north' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(ROUND(SIN(RADIANS(A1)),6)=0.5,ROUND(COS(RADIANS(60)),6)=0.5)',
            apply: { fill: '#trig-sin-cos' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula:
              '=AND(ROUND(TAN(RADIANS(A2)),6)=1,ROUND(DEGREES(PI()),6)=180,ROUND(DEGREES(ATAN2(0,-1)),6)=-90)',
            apply: { fill: '#trig-tan-pi' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=ATAN2(0,0)=0',
            apply: { fill: '#atan2-zero-origin' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula:
              '=AND(ROUND(SEC(RADIANS(60)),6)=2,ROUND(CSC(RADIANS(30)),6)=2,ROUND(COT(RADIANS(45)),6)=1)',
            apply: { fill: '#reciprocal-trig' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 3, r1: 1, c1: 3 },
            formula: '=COT(0)=0',
            apply: { fill: '#cot-zero' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
            formula: '=SIN(A3)=0',
            apply: { fill: '#trig-text' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#trig-sin-cos');
    expect(overlay.get('0:1:1')?.fill).toBe('#trig-tan-pi');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBe('#reciprocal-trig');
    expect(overlay.get('0:1:3')?.fill).toBeUndefined();
    expect(overlay.get('0:2:1')?.fill).toBeUndefined();
  });

  it('formula rules evaluate inverse and hyperbolic numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 0.5);
    s = seedNumber(s, 1, 0, 2);
    s = seedNumber(s, 2, 0, 4);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula:
              '=AND(ROUND(DEGREES(ASIN(A1)),6)=30,ROUND(DEGREES(ACOS(A1)),6)=60,ROUND(DEGREES(ATAN(1)),6)=45,ROUND(DEGREES(ACOT(1)),6)=45)',
            apply: { fill: '#inverse-trig' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula:
              '=AND(ROUND(SINH(1),6)=1.175201,ROUND(COSH(1),6)=1.543081,ROUND(TANH(1),6)=0.761594,ROUND(COTH(1),6)=1.313035,ROUND(SECH(1),6)=0.648054,ROUND(CSCH(1),6)=0.850918,ROUND(ASINH(1),6)=0.881374,ROUND(ACOSH(2),6)=1.316958,ROUND(ATANH(0.5),6)=0.549306,ROUND(ACOTH(2),6)=0.549306)',
            apply: { fill: '#hyperbolic' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
            formula: '=ROUND(SQRTPI(A3),6)=3.544908',
            apply: { fill: '#sqrtpi' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=ASIN(A2)=0',
            apply: { fill: '#asin-out-of-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
            formula: '=ACOSH(0)=0',
            apply: { fill: '#acosh-out-of-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 3, r1: 2, c1: 3 },
            formula: '=COTH(0)=0',
            apply: { fill: '#coth-zero' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 4, r1: 2, c1: 4 },
            formula: '=ACOTH(0.5)=0',
            apply: { fill: '#acoth-out-of-range' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#inverse-trig');
    expect(overlay.get('0:1:1')?.fill).toBe('#hyperbolic');
    expect(overlay.get('0:2:1')?.fill).toBe('#sqrtpi');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
    expect(overlay.get('0:2:2')?.fill).toBeUndefined();
    expect(overlay.get('0:2:3')?.fill).toBeUndefined();
    expect(overlay.get('0:2:4')?.fill).toBeUndefined();
  });

  it('formula rules evaluate SUMSQ and SERIESSUM numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 3);
    s = seedNumber(s, 0, 1, 4);
    s = seedCell(s, 1, 0, { kind: 'text', value: 'north' });
    s = seedNumber(s, 2, 0, 1);
    s = seedNumber(s, 2, 1, 2);
    s = seedNumber(s, 2, 2, 3);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=SUMSQ(A1,B1,12)=169',
            apply: { fill: '#sumsq' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=SUMSQ(A2,1)=1',
            apply: { fill: '#sumsq-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=SERIESSUM(2,1,1,$A$3:$C$3)=34',
            apply: { fill: '#series-sum' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=SERIESSUM(2,0,2,3)=3',
            apply: { fill: '#series-sum-scalar' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=SERIESSUM(2,1,1,A2)=0',
            apply: { fill: '#series-sum-text' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#sumsq');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBe('#series-sum');
    expect(overlay.get('0:0:4')?.fill).toBe('#series-sum-scalar');
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
  });

  it('formula rules evaluate logarithmic numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 7);
    s = seedNumber(s, 1, 0, 8);
    s = seedNumber(s, 2, 0, -1);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(ROUND(EXP(LN(A1)),6)=7,ROUND(LOG10(1000),6)=3)',
            apply: { fill: '#exp-ln-log10' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=AND(ROUND(LOG(A2,2),6)=3,ROUND(LOG(100),6)=2)',
            apply: { fill: '#log' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=ROUND(LOG(100,),6)=2',
            apply: { fill: '#log-omitted-base' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
            formula: '=LN(A3)=0',
            apply: { fill: '#ln-negative' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=LOG(A2,1)=0',
            apply: { fill: '#log-base-one' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#exp-ln-log10');
    expect(overlay.get('0:1:1')?.fill).toBe('#log');
    expect(overlay.get('0:0:3')?.fill).toBe('#log-omitted-base');
    expect(overlay.get('0:2:1')?.fill).toBeUndefined();
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate error function numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'north' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula:
              '=AND(ROUND(ERF(1),6)=0.842701,ROUND(ERF(0,1),6)=0.842701,ROUND(ERF.PRECISE(1),6)=0.842701)',
            apply: { fill: '#erf' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula:
              '=AND(ROUND(ERFC(1),6)=0.157299,ROUND(ERFC.PRECISE(1),6)=0.157299,ROUND(GAUSS(1),6)=0.341345)',
            apply: { fill: '#erfc-gauss' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=ERF(A1)=0',
            apply: { fill: '#erf-text' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#erf');
    expect(overlay.get('0:0:2')?.fill).toBe('#erfc-gauss');
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
  });

  it('formula rules evaluate base conversion numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            formula:
              '=AND(BASE(255,16)="FF",BASE(5,2,8)="00000101",DECIMAL("FF",16)=255,BIN2DEC("1111111111")=-1,DEC2BIN(5,8)="00000101",DEC2BIN(-1)="1111111111",HEX2DEC("FFFFFFFFFF")=-1,DEC2HEX(255,4)="00FF",DEC2HEX(-1)="FFFFFFFFFF",OCT2DEC("7777777777")=-1,DEC2OCT(64,4)="0100",DEC2OCT(-1)="7777777777",BIN2HEX("1111",4)="000F",HEX2BIN("F",8)="00001111",BIN2OCT("1111",4)="0017",OCT2BIN("17",8)="00001111",HEX2OCT("F",4)="0017",OCT2HEX("17",4)="000F")',
            apply: { fill: '#base-decimal' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=BASE(10,1)="10"',
            apply: { fill: '#base-invalid-radix' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=DECIMAL("2",2)=2',
            apply: { fill: '#decimal-invalid-digit' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=BIN2DEC("10000000000")=0',
            apply: { fill: '#bin2dec-too-long' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=DEC2BIN(512)="1000000000"',
            apply: { fill: '#dec2bin-out-of-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=HEX2DEC("G")=16',
            apply: { fill: '#hex2dec-invalid-digit' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=DEC2HEX(549755813888)="8000000000"',
            apply: { fill: '#dec2hex-out-of-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=OCT2DEC("8")=8',
            apply: { fill: '#oct2dec-invalid-digit' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=DEC2OCT(536870912)="4000000000"',
            apply: { fill: '#dec2oct-out-of-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=HEX2BIN("200")="1000000000"',
            apply: { fill: '#hex2bin-out-of-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 10, r1: 0, c1: 10 },
            formula: '=BIN2HEX("10000000000")="400"',
            apply: { fill: '#bin2hex-too-long' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 11, r1: 0, c1: 11 },
            formula: '=OCT2HEX("8")="8"',
            apply: { fill: '#oct2hex-invalid-digit' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#base-decimal');
    expect(overlay.get('0:0:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
    expect(overlay.get('0:0:4')?.fill).toBeUndefined();
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
    expect(overlay.get('0:0:6')?.fill).toBeUndefined();
    expect(overlay.get('0:0:7')?.fill).toBeUndefined();
    expect(overlay.get('0:0:8')?.fill).toBeUndefined();
    expect(overlay.get('0:0:9')?.fill).toBeUndefined();
    expect(overlay.get('0:0:10')?.fill).toBeUndefined();
    expect(overlay.get('0:0:11')?.fill).toBeUndefined();
  });

  it('formula rules evaluate roman numeral numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            formula: '=AND(ROMAN(1999)="MCMXCIX",ROMAN(944,0)="CMXLIV",ARABIC("MCMXCIX")=1999)',
            apply: { fill: '#roman-arabic' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=ROMAN(0)="N"',
            apply: { fill: '#roman-zero' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=ARABIC("IIII")=4',
            apply: { fill: '#arabic-invalid' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#roman-arabic');
    expect(overlay.get('0:0:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate threshold numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'north' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(DELTA(4,4)=1,DELTA(4,5)=0,DELTA(0)=1,DELTA(0,)=1)',
            apply: { fill: '#delta' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=AND(GESTEP(5,4)=1,GESTEP(4,5)=0,GESTEP(0)=1,GESTEP(0,)=1)',
            apply: { fill: '#gestep' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=DELTA(A1,0)=0',
            apply: { fill: '#delta-text' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#delta');
    expect(overlay.get('0:0:2')?.fill).toBe('#gestep');
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
  });

  it('formula rules evaluate bitwise numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            formula: '=AND(BITAND(13,7)=5,BITOR(9,6)=15,BITXOR(9,6)=15)',
            apply: { fill: '#bitwise' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula:
              '=AND(BITLSHIFT(3,4)=48,BITRSHIFT(48,4)=3,BITLSHIFT(48,-4)=3,BITRSHIFT(3,-4)=48)',
            apply: { fill: '#bitshift' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=BITAND(-1,1)=1',
            apply: { fill: '#bit-negative' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=BITLSHIFT(1,54)=0',
            apply: { fill: '#bit-shift-too-large' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#bitwise');
    expect(overlay.get('0:0:1')?.fill).toBe('#bitshift');
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
  });

  it('formula rules evaluate character and text cleanup functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            formula: '=AND(CHAR(65)="A",CODE("Apple")=65,UNICODE(UNICHAR(9731))=9731)',
            apply: { fill: '#char-code' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(CLEAN("A"&CHAR(10)&"B")="AB",PROPER("north region")="North Region")',
            apply: { fill: '#clean-proper' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=CHAR(0)=""',
            apply: { fill: '#char-zero' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=UNICODE("")=0',
            apply: { fill: '#unicode-empty' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=ENCODEURL("North Region/東京")="North%20Region%2F%E6%9D%B1%E4%BA%AC"',
            apply: { fill: '#encodeurl' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=ENCODEURL(NA())=""',
            apply: { fill: '#encodeurl-error' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#char-code');
    expect(overlay.get('0:0:1')?.fill).toBe('#clean-proper');
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
    expect(overlay.get('0:0:4')?.fill).toBe('#encodeurl');
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
  });

  it('formula rules evaluate financial numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 1, 0, 0.09);
    s = seedNumber(s, 1, 1, 0.11);
    s = seedNumber(s, 1, 2, 0.1);
    s = seedNumber(s, 2, 0, 1000);
    s = seedNumber(s, 2, 1, 2000);
    s = seedNumber(s, 2, 2, 3000);
    s = seedNumber(s, 2, 3, 4000);
    s = seedNumber(s, 3, 0, -120000);
    s = seedNumber(s, 3, 1, 39000);
    s = seedNumber(s, 3, 2, 30000);
    s = seedNumber(s, 3, 3, 21000);
    s = seedNumber(s, 3, 4, 37000);
    s = seedNumber(s, 3, 5, 46000);
    s = seedNumber(s, 4, 0, -10000);
    s = seedNumber(s, 4, 1, 2750);
    s = seedNumber(s, 4, 2, 4250);
    s = seedNumber(s, 4, 3, 3250);
    s = seedNumber(s, 4, 4, 2750);
    s = seedNumber(s, 5, 0, dateSerial(2024, 1, 1));
    s = seedNumber(s, 5, 1, dateSerial(2024, 3, 1));
    s = seedNumber(s, 5, 2, dateSerial(2024, 10, 30));
    s = seedNumber(s, 5, 3, dateSerial(2025, 2, 15));
    s = seedNumber(s, 5, 4, dateSerial(2025, 4, 1));
    s = seedNumber(s, 5, 5, dateSerial(2024, 7, 1));
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            formula:
              '=AND(ROUND(PMT(0.05/12,60,10000),6)=-188.712336,ROUND(PMT(0.05/12,60,10000,,1),6)=-187.929298)',
            apply: { fill: '#pmt' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(ROUND(PV(0.05/12,60,-188.712336),6)=9999.999977,PV(0,10,-100,0,0)=1000)',
            apply: { fill: '#pv' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula:
              '=AND(ROUND(FV(0.05/12,60,-100),6)=6800.608284,ROUND(FV(0.05/12,60,-100,,1),6)=6828.944152,FV(0,10,-100)=1000)',
            apply: { fill: '#fv' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula:
              '=AND(ROUND(NPER(0.05/12,-188.712336,10000),6)=60,NPER(0,-100,1000)=10,ROUND(NPER(0.05/12,-187.929298,10000,,1),6)=60)',
            apply: { fill: '#nper' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula:
              '=AND(ROUND(RATE(60,-188.712336,10000),9)=0.004166667,ROUND(RATE(60,-187.929298,10000,,1),9)=0.004166667,ROUND(RATE(10,-100,1000),9)=0)',
            apply: { fill: '#rate' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula:
              '=AND(ROUND(IPMT(0.05/12,1,60,10000),6)=-41.666667,ROUND(PPMT(0.05/12,1,60,10000),6)=-147.04567,ROUND(IPMT(0.05/12,2,60,10000),6)=-41.053976)',
            apply: { fill: '#ipmt-ppmt' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula:
              '=AND(IPMT(0.05/12,1,60,10000,,1)=0,ROUND(PPMT(0.05/12,1,60,10000,,1),6)=-187.929298,ROUND(IPMT(0.05/12,2,60,10000,,1),6)=-40.100589,PPMT(0,1,10,1000)=-100)',
            apply: { fill: '#ipmt-ppmt-type' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=PMT(0.05/12,0,10000)=0',
            apply: { fill: '#pmt-zero-periods' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=FV(0.05/12,60,-100,0,2)=0',
            apply: { fill: '#fv-invalid-type' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=NPER(0.05/12,0,10000)=0',
            apply: { fill: '#nper-zero-payment' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 10, r1: 0, c1: 10 },
            formula: '=RATE(0,-100,1000)=0',
            apply: { fill: '#rate-zero-periods' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 11, r1: 0, c1: 11 },
            formula: '=IPMT(0.05/12,0,60,10000)=0',
            apply: { fill: '#ipmt-invalid-period' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 12, r1: 0, c1: 12 },
            formula:
              '=AND(SLN(30000,7500,10)=2250,ROUND(SYD(30000,7500,10,1),6)=4090.909091,ROUND(SYD(30000,7500,10,10),6)=409.090909)',
            apply: { fill: '#depreciation-linear-syd' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 13, r1: 0, c1: 13 },
            formula:
              '=AND(DDB(2400,300,10,1)=480,DDB(2400,300,10,2)=384,DDB(2400,300,10,1,1.5)=360)',
            apply: { fill: '#depreciation-ddb' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 14, r1: 0, c1: 14 },
            formula: '=DDB(2400,300,10,0)=0',
            apply: { fill: '#ddb-invalid-period' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 15, r1: 0, c1: 15 },
            formula:
              '=AND(ROUND(DB(1000000,100000,6,1,7),2)=186083.33,ROUND(DB(1000000,100000,6,2,7),2)=259639.42,ROUND(DB(1000000,100000,6,7,7),2)=15845.1)',
            apply: { fill: '#depreciation-db' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 16, r1: 0, c1: 16 },
            formula: '=DB(1000000,100000,6,8,7)=0',
            apply: { fill: '#db-invalid-period' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 17, r1: 0, c1: 17 },
            formula: '=DB(1000000,100000,6,1,13)=0',
            apply: { fill: '#db-invalid-month' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 18, r1: 0, c1: 18 },
            formula:
              '=AND(ROUND(CUMIPMT(0.05/12,60,10000,1,12,0),6)=-458.995507,ROUND(CUMPRINC(0.05/12,60,10000,1,12,0),6)=-1805.55253)',
            apply: { fill: '#cumipmt-cumprinc' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 19, r1: 0, c1: 19 },
            formula:
              '=AND(ROUND(CUMIPMT(0.05/12,60,10000,1,12,1),6)=-406.802051,ROUND(CUMPRINC(0.05/12,60,10000,1,12,1),6)=-1848.349521)',
            apply: { fill: '#cumipmt-cumprinc-type' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 20, r1: 0, c1: 20 },
            formula: '=CUMIPMT(0,60,10000,1,12,0)=0',
            apply: { fill: '#cumipmt-invalid-rate' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 21, r1: 0, c1: 21 },
            formula: '=CUMPRINC(0.05/12,60,10000,12,1,0)=0',
            apply: { fill: '#cumprinc-invalid-periods' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 22, r1: 0, c1: 22 },
            formula:
              '=AND(ROUND(ISPMT(0.1/12,1,36,8000),6)=-64.814815,ROUND(ISPMT(0.1/12,0,36,8000),6)=-66.666667)',
            apply: { fill: '#ispmt' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 23, r1: 0, c1: 23 },
            formula:
              '=AND(ROUND(EFFECT(0.0525,4),9)=0.053542667,ROUND(NOMINAL(0.05354266737075822,4),9)=0.0525)',
            apply: { fill: '#effect-nominal' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 24, r1: 0, c1: 24 },
            formula: '=ISPMT(0.1/12,1,0,8000)=0',
            apply: { fill: '#ispmt-zero-periods' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 25, r1: 0, c1: 25 },
            formula: '=EFFECT(0.0525,0)=0',
            apply: { fill: '#effect-invalid-periods' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 26, r1: 0, c1: 26 },
            formula: '=NOMINAL(0,4)=0',
            apply: { fill: '#nominal-invalid-rate' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 27, r1: 0, c1: 27 },
            formula:
              '=AND(FVSCHEDULE(1000,$A$2:$C$2)=1330.89,ROUND(RRI(10,1000,10000),6)=0.258925,ROUND(PDURATION(0.1,1000,10000),6)=24.158858)',
            apply: { fill: '#fvschedule-rri-pduration' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 28, r1: 0, c1: 28 },
            formula: '=FVSCHEDULE("principal",$A$2:$C$2)=0',
            apply: { fill: '#fvschedule-invalid-principal' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 29, r1: 0, c1: 29 },
            formula: '=RRI(0,1000,10000)=0',
            apply: { fill: '#rri-invalid-periods' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 30, r1: 0, c1: 30 },
            formula: '=PDURATION(0,1000,10000)=0',
            apply: { fill: '#pduration-invalid-rate' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 31, r1: 0, c1: 31 },
            formula:
              '=AND(ROUND(NPV(0.1,-10000,3000,4200,6800),6)=1188.443412,ROUND(NPV(0.08,$A$3:$C$3,D3),6)=7962.219701)',
            apply: { fill: '#npv' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 32, r1: 0, c1: 32 },
            formula: '=NPV(-1,$A$3:$C$3)=0',
            apply: { fill: '#npv-invalid-rate' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 33, r1: 0, c1: 33 },
            formula: '=NPV(0.1,"bad")=0',
            apply: { fill: '#npv-invalid-value' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 34, r1: 0, c1: 34 },
            formula: '=ROUND(MIRR($A$4:$F$4,0.1,0.12),6)=0.126094',
            apply: { fill: '#mirr' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 35, r1: 0, c1: 35 },
            formula: '=MIRR($A$3:$D$3,0.1,0.12)=0',
            apply: { fill: '#mirr-no-negative' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 36, r1: 0, c1: 36 },
            formula: '=MIRR($A$4:$F$4,-1,0.12)=0',
            apply: { fill: '#mirr-invalid-rate' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 37, r1: 0, c1: 37 },
            formula: '=ROUND(XNPV(0.09,$A$5:$E$5,$A$6:$E$6),6)=2086.647602',
            apply: { fill: '#xnpv' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 38, r1: 0, c1: 38 },
            formula: '=XNPV(0.09,$A$5:$E$5,$A$6:$D$6)=0',
            apply: { fill: '#xnpv-mismatched-ranges' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 39, r1: 0, c1: 39 },
            formula: '=XNPV(0.09,$A$5:$E$5,$B$6:$F$6)=0',
            apply: { fill: '#xnpv-invalid-date' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 40, r1: 0, c1: 40 },
            formula:
              '=AND(ROUND(XIRR($A$5:$E$5,$A$6:$E$6),9)=0.373362534,ROUND(XIRR($A$5:$E$5,$A$6:$E$6,0.2),9)=0.373362534)',
            apply: { fill: '#xirr' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 41, r1: 0, c1: 41 },
            formula: '=XIRR($A$5:$E$5,$A$6:$D$6)=0',
            apply: { fill: '#xirr-mismatched-ranges' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 42, r1: 0, c1: 42 },
            formula: '=XIRR($A$3:$D$3,$A$6:$D$6)=0',
            apply: { fill: '#xirr-no-negative' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 43, r1: 0, c1: 43 },
            formula:
              '=AND(ROUND(IRR($A$4:$F$4),9)=0.130735539,ROUND(IRR($A$4:$F$4,0.2),9)=0.130735539)',
            apply: { fill: '#irr' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 44, r1: 0, c1: 44 },
            formula: '=IRR($A$3:$D$3)=0',
            apply: { fill: '#irr-no-negative' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 45, r1: 0, c1: 45 },
            formula: '=IRR($A$4:$F$4,-1)=0',
            apply: { fill: '#irr-invalid-guess' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 46, r1: 0, c1: 46 },
            formula:
              '=AND(DOLLARDE(1.02,16)=1.125,DOLLARFR(1.125,16)=1.02,DOLLARDE(-1.02,16)=-1.125)',
            apply: { fill: '#dollarde-dollarfr' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 47, r1: 0, c1: 47 },
            formula: '=DOLLARDE(1.02,0)=0',
            apply: { fill: '#dollarde-invalid-fraction' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 48, r1: 0, c1: 48 },
            formula: '=DOLLARFR(1.125,0)=0',
            apply: { fill: '#dollarfr-invalid-fraction' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 49, r1: 0, c1: 49 },
            formula:
              '=AND(ROUND(DISC($A$6,$F$6,97.5,100,3),9)=0.050137363,ROUND(INTRATE($A$6,$F$6,9700,10000,3),9)=0.062025603)',
            apply: { fill: '#disc-intrate' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 50, r1: 0, c1: 50 },
            formula: '=DISC($F$6,$A$6,97.5,100,3)=0',
            apply: { fill: '#disc-invalid-dates' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 51, r1: 0, c1: 51 },
            formula: '=INTRATE($A$6,$F$6,9700,10000,5)=0',
            apply: { fill: '#intrate-invalid-basis' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 52, r1: 0, c1: 52 },
            formula:
              '=AND(ROUND(PRICEDISC($A$6,$F$6,0.05,100,3),9)=97.506849315,ROUND(RECEIVED($A$6,$F$6,9700,0.05,3),6)=9948.019106)',
            apply: { fill: '#pricedisc-received' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 53, r1: 0, c1: 53 },
            formula: '=PRICEDISC($A$6,$F$6,0,100,3)=0',
            apply: { fill: '#pricedisc-invalid-discount' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 54, r1: 0, c1: 54 },
            formula: '=RECEIVED($A$6,$F$6,9700,0,3)=0',
            apply: { fill: '#received-invalid-discount' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 55, r1: 0, c1: 55 },
            formula:
              '=AND(ROUND(TBILLPRICE($A$6,$F$6,0.05),9)=97.472222222,ROUND(TBILLYIELD($A$6,$F$6,97.5),9)=0.050718512,ROUND(TBILLEQ($A$6,$F$6,0.05),9)=0.052009119)',
            apply: { fill: '#tbill' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 56, r1: 0, c1: 56 },
            formula: '=TBILLPRICE($F$6,$A$6,0.05)=0',
            apply: { fill: '#tbill-invalid-dates' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 57, r1: 0, c1: 57 },
            formula: '=TBILLYIELD($A$6,$F$6,0)=0',
            apply: { fill: '#tbill-invalid-price' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 58, r1: 0, c1: 58 },
            formula: '=TBILLEQ($A$6,$E$6,0.05)=0',
            apply: { fill: '#tbill-too-long' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 59, r1: 0, c1: 59 },
            formula:
              '=AND(ACCRINTM($A$6,$F$6,0.05)=25,ROUND(ACCRINTM($A$6,$F$6,0.05,1000,3),9)=24.931506849,ACCRINTM($A$6,$F$6,0.05,,)=25)',
            apply: { fill: '#accrintm' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 60, r1: 0, c1: 60 },
            formula: '=ACCRINTM($F$6,$A$6,0.05)=0',
            apply: { fill: '#accrintm-invalid-dates' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 61, r1: 0, c1: 61 },
            formula: '=ACCRINTM($A$6,$F$6,0)=0',
            apply: { fill: '#accrintm-invalid-rate' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#pmt');
    expect(overlay.get('0:0:1')?.fill).toBe('#pv');
    expect(overlay.get('0:0:2')?.fill).toBe('#fv');
    expect(overlay.get('0:0:3')?.fill).toBe('#nper');
    expect(overlay.get('0:0:4')?.fill).toBe('#rate');
    expect(overlay.get('0:0:5')?.fill).toBe('#ipmt-ppmt');
    expect(overlay.get('0:0:6')?.fill).toBe('#ipmt-ppmt-type');
    expect(overlay.get('0:0:7')?.fill).toBeUndefined();
    expect(overlay.get('0:0:8')?.fill).toBeUndefined();
    expect(overlay.get('0:0:9')?.fill).toBeUndefined();
    expect(overlay.get('0:0:10')?.fill).toBeUndefined();
    expect(overlay.get('0:0:11')?.fill).toBeUndefined();
    expect(overlay.get('0:0:12')?.fill).toBe('#depreciation-linear-syd');
    expect(overlay.get('0:0:13')?.fill).toBe('#depreciation-ddb');
    expect(overlay.get('0:0:14')?.fill).toBeUndefined();
    expect(overlay.get('0:0:15')?.fill).toBe('#depreciation-db');
    expect(overlay.get('0:0:16')?.fill).toBeUndefined();
    expect(overlay.get('0:0:17')?.fill).toBeUndefined();
    expect(overlay.get('0:0:18')?.fill).toBe('#cumipmt-cumprinc');
    expect(overlay.get('0:0:19')?.fill).toBe('#cumipmt-cumprinc-type');
    expect(overlay.get('0:0:20')?.fill).toBeUndefined();
    expect(overlay.get('0:0:21')?.fill).toBeUndefined();
    expect(overlay.get('0:0:22')?.fill).toBe('#ispmt');
    expect(overlay.get('0:0:23')?.fill).toBe('#effect-nominal');
    expect(overlay.get('0:0:24')?.fill).toBeUndefined();
    expect(overlay.get('0:0:25')?.fill).toBeUndefined();
    expect(overlay.get('0:0:26')?.fill).toBeUndefined();
    expect(overlay.get('0:0:27')?.fill).toBe('#fvschedule-rri-pduration');
    expect(overlay.get('0:0:28')?.fill).toBeUndefined();
    expect(overlay.get('0:0:29')?.fill).toBeUndefined();
    expect(overlay.get('0:0:30')?.fill).toBeUndefined();
    expect(overlay.get('0:0:31')?.fill).toBe('#npv');
    expect(overlay.get('0:0:32')?.fill).toBeUndefined();
    expect(overlay.get('0:0:33')?.fill).toBeUndefined();
    expect(overlay.get('0:0:34')?.fill).toBe('#mirr');
    expect(overlay.get('0:0:35')?.fill).toBeUndefined();
    expect(overlay.get('0:0:36')?.fill).toBeUndefined();
    expect(overlay.get('0:0:37')?.fill).toBe('#xnpv');
    expect(overlay.get('0:0:38')?.fill).toBeUndefined();
    expect(overlay.get('0:0:39')?.fill).toBeUndefined();
    expect(overlay.get('0:0:40')?.fill).toBe('#xirr');
    expect(overlay.get('0:0:41')?.fill).toBeUndefined();
    expect(overlay.get('0:0:42')?.fill).toBeUndefined();
    expect(overlay.get('0:0:43')?.fill).toBe('#irr');
    expect(overlay.get('0:0:44')?.fill).toBeUndefined();
    expect(overlay.get('0:0:45')?.fill).toBeUndefined();
    expect(overlay.get('0:0:46')?.fill).toBe('#dollarde-dollarfr');
    expect(overlay.get('0:0:47')?.fill).toBeUndefined();
    expect(overlay.get('0:0:48')?.fill).toBeUndefined();
    expect(overlay.get('0:0:49')?.fill).toBe('#disc-intrate');
    expect(overlay.get('0:0:50')?.fill).toBeUndefined();
    expect(overlay.get('0:0:51')?.fill).toBeUndefined();
    expect(overlay.get('0:0:52')?.fill).toBe('#pricedisc-received');
    expect(overlay.get('0:0:53')?.fill).toBeUndefined();
    expect(overlay.get('0:0:54')?.fill).toBeUndefined();
    expect(overlay.get('0:0:55')?.fill).toBe('#tbill');
    expect(overlay.get('0:0:56')?.fill).toBeUndefined();
    expect(overlay.get('0:0:57')?.fill).toBeUndefined();
    expect(overlay.get('0:0:58')?.fill).toBeUndefined();
    expect(overlay.get('0:0:59')?.fill).toBe('#accrintm');
    expect(overlay.get('0:0:60')?.fill).toBeUndefined();
    expect(overlay.get('0:0:61')?.fill).toBeUndefined();
  });

  it('formula rules evaluate financial functions over dynamic ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 0.09);
    s = seedNumber(s, 0, 1, 0.11);
    s = seedNumber(s, 0, 2, 0.1);
    s = seedNumber(s, 1, 0, -120000);
    s = seedNumber(s, 1, 1, 39000);
    s = seedNumber(s, 1, 2, 30000);
    s = seedNumber(s, 1, 3, 21000);
    s = seedNumber(s, 1, 4, 37000);
    s = seedNumber(s, 1, 5, 46000);
    s = seedNumber(s, 2, 0, -10000);
    s = seedNumber(s, 2, 1, 2750);
    s = seedNumber(s, 2, 2, 4250);
    s = seedNumber(s, 2, 3, 3250);
    s = seedNumber(s, 2, 4, 2750);
    s = seedNumber(s, 3, 0, dateSerial(2024, 1, 1));
    s = seedNumber(s, 3, 1, dateSerial(2024, 3, 1));
    s = seedNumber(s, 3, 2, dateSerial(2024, 10, 30));
    s = seedNumber(s, 3, 3, dateSerial(2025, 2, 15));
    s = seedNumber(s, 3, 4, dateSerial(2025, 4, 1));
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula:
              '=AND(FVSCHEDULE(1000,OFFSET(A1,0,0,1,3))=1330.89,ROUND(IRR(OFFSET(A2,0,0,1,6)),9)=0.130735539,ROUND(MIRR(INDIRECT("A2:F2"),0.1,0.12),6)=0.126094)',
            apply: { fill: '#financial-dynamic' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 6, r1: 1, c1: 6 },
            formula:
              '=AND(ROUND(XNPV(0.09,OFFSET(A3,0,0,1,5),INDIRECT("A4:E4")),6)=2086.647602,ROUND(XIRR(INDIRECT("A3:E3"),OFFSET(A4,0,0,1,5)),9)=0.373362534)',
            apply: { fill: '#financial-dynamic-x' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 6, r1: 2, c1: 6 },
            formula: '=XNPV(0.09,OFFSET(A3,0,0,1,5),INDIRECT("A4:D4"))=0',
            apply: { fill: '#financial-dynamic-mismatch' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:6')?.fill).toBe('#financial-dynamic');
    expect(overlay.get('0:1:6')?.fill).toBe('#financial-dynamic-x');
    expect(overlay.get('0:2:6')?.fill).toBeUndefined();
  });

  it('formula rules evaluate normal distribution numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            formula:
              '=AND(STANDARDIZE(42,40,2)=1,ROUND(NORM.S.DIST(1,TRUE()),6)=0.841345,ROUND(NORM.S.DIST(0,FALSE()),6)=0.398942,ROUND(PHI(0),6)=0.398942)',
            apply: { fill: '#standard-normal' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula:
              '=AND(ROUND(NORM.DIST(42,40,2,TRUE()),6)=0.841345,ROUND(NORM.DIST(42,40,2,FALSE()),6)=0.120985,ROUND(CONFIDENCE.NORM(0.05,2.5,50),6)=0.692952)',
            apply: { fill: '#normal' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 11, r1: 0, c1: 11 },
            formula: '=ROUND(CONFIDENCE.T(0.05,2.5,50),6)=ROUND(T.INV.2T(0.05,49)*2.5/SQRT(50),6)',
            apply: { fill: '#confidence-t' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula:
              '=AND(ROUND(NORMSDIST(1),6)=0.841345,ROUND(NORMDIST(42,40,2,FALSE()),6)=0.120985,ROUND(CONFIDENCE(0.05,2.5,50),6)=0.692952)',
            apply: { fill: '#legacy-normal' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula:
              '=AND(ROUND(FISHER(0.75),6)=0.972955,ROUND(FISHERINV(0.9729550745276566),6)=0.75)',
            apply: { fill: '#fisher' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=FISHER(1)=0',
            apply: { fill: '#fisher-out-of-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula:
              '=AND(ROUND(NORM.S.INV(0.841344746068543),6)=1,ROUND(NORM.INV(0.841344746068543,40,2),6)=42,ROUND(LOGNORM.INV(0.841344746068543,1,0.5),6)=4.481689)',
            apply: { fill: '#normal-inv' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula:
              '=AND(ROUND(NORMSINV(0.841344746068543),6)=1,ROUND(NORMINV(0.841344746068543,40,2),6)=42,ROUND(LOGINV(0.841344746068543,1,0.5),6)=4.481689)',
            apply: { fill: '#legacy-normal-inv' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=NORM.S.INV(0)=0',
            apply: { fill: '#normal-inv-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 10, r1: 0, c1: 10 },
            formula: '=CONFIDENCE.NORM(0,2.5,50)=0',
            apply: { fill: '#confidence-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 12, r1: 0, c1: 12 },
            formula: '=CONFIDENCE.T(0.05,2.5,1)=0',
            apply: { fill: '#confidence-t-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=STANDARDIZE(42,40,0)=0',
            apply: { fill: '#standardize-zero-deviation' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=NORM.DIST(42,40,0,TRUE())=0',
            apply: { fill: '#normal-zero-deviation' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#standard-normal');
    expect(overlay.get('0:0:1')?.fill).toBe('#normal');
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
    expect(overlay.get('0:0:4')?.fill).toBe('#legacy-normal');
    expect(overlay.get('0:0:5')?.fill).toBe('#fisher');
    expect(overlay.get('0:0:6')?.fill).toBeUndefined();
    expect(overlay.get('0:0:7')?.fill).toBe('#normal-inv');
    expect(overlay.get('0:0:8')?.fill).toBe('#legacy-normal-inv');
    expect(overlay.get('0:0:9')?.fill).toBeUndefined();
    expect(overlay.get('0:0:10')?.fill).toBeUndefined();
  });

  it('formula rules evaluate probability distribution numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            formula:
              '=AND(ROUND(BINOM.DIST(6,10,0.5,FALSE()),6)=0.205078,ROUND(BINOM.DIST(6,10,0.5,TRUE()),6)=0.828125,BINOM.INV(10,0.5,0.8)=6)',
            apply: { fill: '#binom-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula:
              '=AND(ROUND(POISSON.DIST(2,5,FALSE()),6)=0.084224,ROUND(POISSON.DIST(2,5,TRUE()),6)=0.124652)',
            apply: { fill: '#poisson-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula:
              '=AND(ROUND(EXPON.DIST(0.2,10,TRUE()),6)=0.864665,ROUND(EXPON.DIST(0.2,10,FALSE()),6)=1.353353)',
            apply: { fill: '#expon-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula:
              '=AND(ROUND(BINOMDIST(6,10,0.5,FALSE()),6)=0.205078,CRITBINOM(10,0.5,0.8)=6,ROUND(POISSON(2,5,TRUE()),6)=0.124652,ROUND(EXPONDIST(0.2,10,FALSE()),6)=1.353353)',
            apply: { fill: '#legacy-distributions' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula:
              '=AND(ROUND(LOGNORM.DIST(4,1,0.5,TRUE()),6)=0.780117,ROUND(LOGNORM.DIST(4,1,0.5,FALSE()),6)=0.148002)',
            apply: { fill: '#lognorm-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula:
              '=AND(ROUND(WEIBULL.DIST(3,2,4,TRUE()),6)=0.430217,ROUND(WEIBULL.DIST(3,2,4,FALSE()),6)=0.213669)',
            apply: { fill: '#weibull-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 10, r1: 0, c1: 10 },
            formula:
              '=AND(ROUND(LOGNORMDIST(4,1,0.5),6)=0.780117,ROUND(WEIBULL(3,2,4,FALSE()),6)=0.213669)',
            apply: { fill: '#legacy-lognorm-weibull' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 17, r1: 0, c1: 17 },
            formula:
              '=AND(ROUND(GAMMA.DIST(4,2,3,TRUE()),6)=0.38494,ROUND(GAMMA.DIST(4,2,3,FALSE()),6)=0.117154)',
            apply: { fill: '#gamma-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 18, r1: 0, c1: 18 },
            formula:
              '=AND(ROUND(GAMMADIST(2,3,2,TRUE()),6)=0.080301,ROUND(GAMMADIST(2,3,2,FALSE()),6)=0.09197)',
            apply: { fill: '#legacy-gamma-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 26, r1: 0, c1: 26 },
            formula:
              '=AND(ROUND(GAMMA.INV(0.384940011063304,2,3),6)=4,ROUND(GAMMAINV(0.080301397071394,3,2),6)=2)',
            apply: { fill: '#gamma-inv' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 28, r1: 0, c1: 28 },
            formula:
              '=AND(ROUND(BETA.DIST(0.4,2,3,TRUE()),6)=0.5248,ROUND(BETA.DIST(0.4,2,3,FALSE()),6)=1.728)',
            apply: { fill: '#beta-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 29, r1: 0, c1: 29 },
            formula:
              '=AND(ROUND(BETA.DIST(4,2,3,TRUE(),0,10),6)=0.5248,ROUND(BETA.DIST(4,2,3,FALSE(),0,10),6)=0.1728)',
            apply: { fill: '#beta-dist-scaled' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 30, r1: 0, c1: 30 },
            formula:
              '=AND(ROUND(BETADIST(0.4,2,3),6)=0.5248,ROUND(BETA.INV(0.5248,2,3,0,10),6)=4,ROUND(BETAINV(0.5248,2,3),6)=0.4)',
            apply: { fill: '#beta-inv-legacy' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 33, r1: 0, c1: 33 },
            formula:
              '=AND(ROUND(F.DIST(3,2,10,TRUE()),6)=0.904633,ROUND(F.DIST(3,2,10,FALSE()),6)=0.059605)',
            apply: { fill: '#f-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 34, r1: 0, c1: 34 },
            formula: '=AND(ROUND(F.DIST.RT(3,2,10),6)=0.095367,ROUND(FDIST(3,2,10),6)=0.095367)',
            apply: { fill: '#f-dist-rt' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 35, r1: 0, c1: 35 },
            formula:
              '=AND(ROUND(F.INV(0.904632568359375,2,10),6)=3,ROUND(F.INV.RT(0.095367431640625,2,10),6)=3,ROUND(FINV(0.095367431640625,2,10),6)=3)',
            apply: { fill: '#f-inv' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 38, r1: 0, c1: 38 },
            formula:
              '=AND(ROUND(T.DIST(1,10,TRUE()),6)=0.829553,ROUND(T.DIST(1,10,FALSE()),6)=0.230362)',
            apply: { fill: '#t-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 39, r1: 0, c1: 39 },
            formula:
              '=AND(ROUND(T.DIST.RT(1,10),6)=0.170447,ROUND(T.DIST.2T(1,10),6)=0.340893,ROUND(TDIST(1,10,2),6)=0.340893)',
            apply: { fill: '#t-dist-tails' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 40, r1: 0, c1: 40 },
            formula:
              '=AND(ROUND(T.INV(0.82955343384897,10),6)=1,ROUND(T.INV.2T(0.34089313230206,10),6)=1,ROUND(TINV(0.34089313230206,10),6)=1)',
            apply: { fill: '#t-inv' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 21, r1: 0, c1: 21 },
            formula:
              '=AND(ROUND(CHISQ.DIST(3,2,TRUE()),6)=0.77687,ROUND(CHISQ.DIST(3,2,FALSE()),6)=0.111565)',
            apply: { fill: '#chisq-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 22, r1: 0, c1: 22 },
            formula: '=AND(ROUND(CHISQ.DIST.RT(3,2),6)=0.22313,ROUND(CHIDIST(3,2),6)=0.22313)',
            apply: { fill: '#chisq-dist-rt' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 24, r1: 0, c1: 24 },
            formula:
              '=AND(ROUND(CHISQ.INV(0.77686983985157,2),6)=3,ROUND(CHISQ.INV.RT(0.22313016014843,2),6)=3,ROUND(CHIINV(0.22313016014843,2),6)=3)',
            apply: { fill: '#chisq-inv' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 11, r1: 0, c1: 11 },
            formula:
              '=AND(ROUND(NEGBINOM.DIST(10,5,0.25,FALSE()),6)=0.055049,ROUND(NEGBINOM.DIST(10,5,0.25,TRUE()),6)=0.313514)',
            apply: { fill: '#negbinom-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 12, r1: 0, c1: 12 },
            formula:
              '=AND(ROUND(HYPGEOM.DIST(1,4,8,20,FALSE()),6)=0.363261,ROUND(HYPGEOM.DIST(1,4,8,20,TRUE()),6)=0.465428)',
            apply: { fill: '#hypgeom-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 15, r1: 0, c1: 15 },
            formula:
              '=AND(ROUND(NEGBINOMDIST(10,5,0.25),6)=0.055049,ROUND(HYPGEOMDIST(1,4,8,20),6)=0.363261)',
            apply: { fill: '#legacy-discrete-distributions' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=BINOM.DIST(11,10,0.5,FALSE())=0',
            apply: { fill: '#binom-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 16, r1: 0, c1: 16 },
            formula: '=BINOM.INV(10,0.5,0)=0',
            apply: { fill: '#binom-inv-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 19, r1: 0, c1: 19 },
            formula: '=GAMMA.DIST(1,0,2,TRUE())=0',
            apply: { fill: '#gamma-dist-invalid-alpha' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 20, r1: 0, c1: 20 },
            formula: '=GAMMA.DIST(1,2,0,TRUE())=0',
            apply: { fill: '#gamma-dist-invalid-beta' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 27, r1: 0, c1: 27 },
            formula: '=GAMMA.INV(0,2,3)=0',
            apply: { fill: '#gamma-inv-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 31, r1: 0, c1: 31 },
            formula: '=BETA.DIST(0.4,0,3,TRUE())=0',
            apply: { fill: '#beta-dist-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 32, r1: 0, c1: 32 },
            formula: '=BETA.INV(0,2,3)=0',
            apply: { fill: '#beta-inv-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 36, r1: 0, c1: 36 },
            formula: '=F.DIST(1,0,10,TRUE())=0',
            apply: { fill: '#f-dist-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 37, r1: 0, c1: 37 },
            formula: '=F.INV(0,2,10)=0',
            apply: { fill: '#f-inv-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 41, r1: 0, c1: 41 },
            formula: '=T.DIST.RT(-1,10)=0',
            apply: { fill: '#t-dist-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 42, r1: 0, c1: 42 },
            formula: '=T.INV(0,10)=0',
            apply: { fill: '#t-inv-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 23, r1: 0, c1: 23 },
            formula: '=CHISQ.DIST(1,0,TRUE())=0',
            apply: { fill: '#chisq-dist-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 25, r1: 0, c1: 25 },
            formula: '=CHISQ.INV(0,2)=0',
            apply: { fill: '#chisq-inv-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=EXPON.DIST(1,0,TRUE())=0',
            apply: { fill: '#expon-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=LOGNORM.DIST(0,1,0.5,TRUE())=0',
            apply: { fill: '#lognorm-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=WEIBULL.DIST(3,0,4,TRUE())=0',
            apply: { fill: '#weibull-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 13, r1: 0, c1: 13 },
            formula: '=NEGBINOM.DIST(10,0,0.25,FALSE())=0',
            apply: { fill: '#negbinom-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 14, r1: 0, c1: 14 },
            formula: '=HYPGEOM.DIST(5,4,8,20,FALSE())=0',
            apply: { fill: '#hypgeom-invalid' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#binom-dist');
    expect(overlay.get('0:0:1')?.fill).toBe('#poisson-dist');
    expect(overlay.get('0:0:2')?.fill).toBe('#expon-dist');
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
    expect(overlay.get('0:0:4')?.fill).toBeUndefined();
    expect(overlay.get('0:0:5')?.fill).toBe('#legacy-distributions');
    expect(overlay.get('0:0:6')?.fill).toBe('#lognorm-dist');
    expect(overlay.get('0:0:7')?.fill).toBe('#weibull-dist');
    expect(overlay.get('0:0:8')?.fill).toBeUndefined();
    expect(overlay.get('0:0:9')?.fill).toBeUndefined();
    expect(overlay.get('0:0:10')?.fill).toBe('#legacy-lognorm-weibull');
    expect(overlay.get('0:0:11')?.fill).toBe('#negbinom-dist');
    expect(overlay.get('0:0:12')?.fill).toBe('#hypgeom-dist');
    expect(overlay.get('0:0:13')?.fill).toBeUndefined();
    expect(overlay.get('0:0:14')?.fill).toBeUndefined();
    expect(overlay.get('0:0:15')?.fill).toBe('#legacy-discrete-distributions');
    expect(overlay.get('0:0:16')?.fill).toBeUndefined();
    expect(overlay.get('0:0:17')?.fill).toBe('#gamma-dist');
    expect(overlay.get('0:0:18')?.fill).toBe('#legacy-gamma-dist');
    expect(overlay.get('0:0:19')?.fill).toBeUndefined();
    expect(overlay.get('0:0:20')?.fill).toBeUndefined();
    expect(overlay.get('0:0:21')?.fill).toBe('#chisq-dist');
    expect(overlay.get('0:0:22')?.fill).toBe('#chisq-dist-rt');
    expect(overlay.get('0:0:23')?.fill).toBeUndefined();
    expect(overlay.get('0:0:24')?.fill).toBe('#chisq-inv');
    expect(overlay.get('0:0:25')?.fill).toBeUndefined();
    expect(overlay.get('0:0:26')?.fill).toBe('#gamma-inv');
    expect(overlay.get('0:0:27')?.fill).toBeUndefined();
    expect(overlay.get('0:0:28')?.fill).toBe('#beta-dist');
    expect(overlay.get('0:0:29')?.fill).toBe('#beta-dist-scaled');
    expect(overlay.get('0:0:30')?.fill).toBe('#beta-inv-legacy');
    expect(overlay.get('0:0:31')?.fill).toBeUndefined();
    expect(overlay.get('0:0:32')?.fill).toBeUndefined();
    expect(overlay.get('0:0:33')?.fill).toBe('#f-dist');
    expect(overlay.get('0:0:34')?.fill).toBe('#f-dist-rt');
    expect(overlay.get('0:0:35')?.fill).toBe('#f-inv');
    expect(overlay.get('0:0:36')?.fill).toBeUndefined();
    expect(overlay.get('0:0:37')?.fill).toBeUndefined();
  });

  it('formula rules evaluate DATE/YEAR/MONTH/DAY operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, dateSerial(2026, 7, 12));
    s = seedNumber(s, 1, 0, dateSerial(2025, 2, 1));
    s = seedNumber(s, 2, 0, dateSerial(2024, 12, 25));
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(YEAR(A1)=2026,MONTH(A1)=7,DAY(A1)=12,DATE(2026,7,12)=A1)',
            apply: { fill: '#date-parts' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=DATE(2024,14,1)=A2',
            apply: { fill: '#date-overflow' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
            formula: '=AND(DATEVALUE("2024-12-25")=A3,DATEVALUE("12/25/2024")=A3)',
            apply: { fill: '#datevalue' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
            formula: '=DATEVALUE("2024-02-31")=DATE(2024,3,2)',
            apply: { fill: '#datevalue-invalid' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#date-parts');
    expect(overlay.get('0:1:1')?.fill).toBe('#date-overflow');
    expect(overlay.get('0:2:1')?.fill).toBe('#datevalue');
    expect(overlay.get('0:2:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate TIME/HOUR/MINUTE/SECOND operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, dateSerial(2026, 7, 12) + (13 * 3600 + 5 * 60 + 9) / 86_400);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula:
              '=AND(HOUR(A1)=13,MINUTE(A1)=5,SECOND(A1)=9,ROUND(TIME(13,5,9)*86400,0)=47109)',
            apply: { fill: '#time-parts' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=TIME(25,0,0)=1/24',
            apply: { fill: '#time-overflow' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=TIME(-1,0,0)=0',
            apply: { fill: '#time-negative' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula:
              '=AND(ROUND(TIMEVALUE("13:05:09")*86400,0)=47109,TIMEVALUE("1:05 PM")=TIME(13,5,0))',
            apply: { fill: '#timevalue' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=TIMEVALUE("25:00")=1/24',
            apply: { fill: '#timevalue-invalid' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#time-parts');
    expect(overlay.get('0:0:2')?.fill).toBe('#time-overflow');
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
    expect(overlay.get('0:0:4')?.fill).toBe('#timevalue');
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
  });

  it('formula rules evaluate EDATE/EOMONTH date operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, dateSerial(2024, 1, 31));
    s = seedNumber(s, 1, 0, dateSerial(2026, 7, 12));
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(EDATE(A1,1)=DATE(2024,2,29),EOMONTH(A1,1)=DATE(2024,2,29))',
            apply: { fill: '#edate-eomonth-leap' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=AND(EDATE(A2,-1)=DATE(2026,6,12),EOMONTH(A2,0)=DATE(2026,7,31))',
            apply: { fill: '#edate-eomonth' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#edate-eomonth-leap');
    expect(overlay.get('0:1:1')?.fill).toBe('#edate-eomonth');
  });

  it('formula rules evaluate DAYS date operand', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, dateSerial(2026, 7, 12));
    s = seedNumber(s, 0, 1, dateSerial(2026, 7, 1));
    s = seedCell(s, 1, 0, { kind: 'text', value: 'not-a-date' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=DAYS(A1,B1)=11',
            apply: { fill: '#days' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=DAYS(A2,B1)=0',
            apply: { fill: '#days-text' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#days');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
  });

  it('formula rules evaluate DAYS360 date operand', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, dateSerial(2024, 1, 1));
    s = seedNumber(s, 0, 1, dateSerial(2024, 2, 1));
    s = seedNumber(s, 1, 0, dateSerial(2024, 2, 29));
    s = seedNumber(s, 1, 1, dateSerial(2024, 3, 31));
    s = seedCell(s, 2, 0, { kind: 'text', value: 'not-a-date' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=AND(DAYS360(A1,B1)=30,DAYS360(A1,B1,)=30)',
            apply: { fill: '#days360-default' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=AND(DAYS360(A2,B2)=30,DAYS360(A2,B2,TRUE())=31)',
            apply: { fill: '#days360-method' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
            formula: '=DAYS360(A3,B1)=0',
            apply: { fill: '#days360-text' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#days360-default');
    expect(overlay.get('0:1:2')?.fill).toBe('#days360-method');
    expect(overlay.get('0:2:1')?.fill).toBeUndefined();
  });

  it('formula rules evaluate DATEDIF date operand', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, dateSerial(2020, 1, 15));
    s = seedNumber(s, 0, 1, dateSerial(2023, 3, 20));
    s = seedNumber(s, 1, 0, dateSerial(2020, 3, 31));
    s = seedNumber(s, 1, 1, dateSerial(2020, 5, 2));
    s = seedCell(s, 2, 0, { kind: 'text', value: 'not-a-date' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=AND(DATEDIF(A1,B1,"Y")=3,DATEDIF(A1,B1,"M")=38,DATEDIF(A1,B1,"D")=1160)',
            apply: { fill: '#datedif-primary' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=AND(DATEDIF(A1,B1,"YM")=2,DATEDIF(A1,B1,"YD")=64,DATEDIF(A2,B2,"MD")=2)',
            apply: { fill: '#datedif-units' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
            formula: '=DATEDIF(A3,B1,"D")=0',
            apply: { fill: '#datedif-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
            formula: '=DATEDIF(B1,A1,"D")=0',
            apply: { fill: '#datedif-reversed' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 3, r1: 2, c1: 3 },
            formula: '=DATEDIF(A1,B1,"BAD")=0',
            apply: { fill: '#datedif-invalid-unit' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#datedif-primary');
    expect(overlay.get('0:1:2')?.fill).toBe('#datedif-units');
    expect(overlay.get('0:2:1')?.fill).toBeUndefined();
    expect(overlay.get('0:2:2')?.fill).toBeUndefined();
    expect(overlay.get('0:2:3')?.fill).toBeUndefined();
  });

  it('formula rules evaluate YEARFRAC date operand', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, dateSerial(2024, 1, 1));
    s = seedNumber(s, 0, 1, dateSerial(2024, 7, 1));
    s = seedNumber(s, 1, 0, dateSerial(2024, 2, 29));
    s = seedNumber(s, 1, 1, dateSerial(2024, 3, 31));
    s = seedCell(s, 2, 0, { kind: 'text', value: 'not-a-date' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula:
              '=AND(YEARFRAC(A1,B1)=0.5,YEARFRAC(A1,B1,)=0.5,ROUND(YEARFRAC(A1,B1,1),6)=0.497268)',
            apply: { fill: '#yearfrac-primary' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula:
              '=AND(ROUND(YEARFRAC(A1,B1,2),6)=0.505556,ROUND(YEARFRAC(A1,B1,3),6)=0.49863,ROUND(YEARFRAC(A2,B2,4),6)=0.086111)',
            apply: { fill: '#yearfrac-basis' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
            formula: '=YEARFRAC(A3,B1)=0',
            apply: { fill: '#yearfrac-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
            formula: '=YEARFRAC(A1,B1,5)=0',
            apply: { fill: '#yearfrac-invalid-basis' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#yearfrac-primary');
    expect(overlay.get('0:1:2')?.fill).toBe('#yearfrac-basis');
    expect(overlay.get('0:2:1')?.fill).toBeUndefined();
    expect(overlay.get('0:2:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate NETWORKDAYS and WORKDAY date operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, dateSerial(2024, 12, 23));
    s = seedNumber(s, 0, 1, dateSerial(2024, 12, 27));
    s = seedCell(s, 0, 2, { kind: 'text', value: 'north' });
    s = seedNumber(s, 0, 8, dateSerial(2024, 12, 25));
    s = seedNumber(s, 0, 9, dateSerial(2024, 12, 26));
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=AND(NETWORKDAYS(A1,B1)=5,NETWORKDAYS(B1,A1)=-5)',
            apply: { fill: '#networkdays' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=AND(WORKDAY(A1,5)=DATE(2024,12,30),WORKDAY(B1,-5)=DATE(2024,12,20))',
            apply: { fill: '#workday' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=NETWORKDAYS(A1,C1)=0',
            apply: { fill: '#networkdays-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=NETWORKDAYS(A1,B1,$I$1:$J$1)=3',
            apply: { fill: '#networkdays-holidays' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=WORKDAY(A1,2,$I$1:$J$1)=DATE(2024,12,27)',
            apply: { fill: '#workday-holidays' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 10, r1: 0, c1: 10 },
            formula: '=NETWORKDAYS.INTL(A1,B1,7)=4',
            apply: { fill: '#networkdays-intl-code' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 11, r1: 0, c1: 11 },
            formula: '=NETWORKDAYS.INTL(A1,B1,"0011000")=3',
            apply: { fill: '#networkdays-intl-mask' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 12, r1: 0, c1: 12 },
            formula: '=NETWORKDAYS.INTL(A1,B1,,$I$1:$J$1)=3',
            apply: { fill: '#networkdays-intl-omitted-weekend' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 13, r1: 0, c1: 13 },
            formula: '=WORKDAY.INTL(A1,3,7,$I$1:$J$1)=DATE(2024,12,30)',
            apply: { fill: '#workday-intl' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 14, r1: 0, c1: 14 },
            formula: '=NETWORKDAYS.INTL(A1,B1,"1111111")=0',
            apply: { fill: '#networkdays-intl-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 15, r1: 0, c1: 15 },
            formula: '=NETWORKDAYS(A1,B1,OFFSET(I1,0,0,1,2))=3',
            apply: { fill: '#networkdays-dynamic-holidays' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 16, r1: 0, c1: 16 },
            formula: '=WORKDAY(A1,2,INDIRECT("I1:J1"))=DATE(2024,12,27)',
            apply: { fill: '#workday-dynamic-holidays' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 17, r1: 0, c1: 17 },
            formula: '=NETWORKDAYS.INTL(A1,B1,,OFFSET(I1,0,0,1,2))=3',
            apply: { fill: '#networkdays-intl-dynamic-holidays' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#networkdays');
    expect(overlay.get('0:0:4')?.fill).toBe('#workday');
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
    expect(overlay.get('0:0:6')?.fill).toBe('#networkdays-holidays');
    expect(overlay.get('0:0:7')?.fill).toBe('#workday-holidays');
    expect(overlay.get('0:0:10')?.fill).toBe('#networkdays-intl-code');
    expect(overlay.get('0:0:11')?.fill).toBe('#networkdays-intl-mask');
    expect(overlay.get('0:0:12')?.fill).toBe('#networkdays-intl-omitted-weekend');
    expect(overlay.get('0:0:13')?.fill).toBe('#workday-intl');
    expect(overlay.get('0:0:14')?.fill).toBeUndefined();
    expect(overlay.get('0:0:15')?.fill).toBe('#networkdays-dynamic-holidays');
    expect(overlay.get('0:0:16')?.fill).toBe('#workday-dynamic-holidays');
    expect(overlay.get('0:0:17')?.fill).toBe('#networkdays-intl-dynamic-holidays');
  });

  it('formula rules evaluate WEEKDAY return types', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, dateSerial(2026, 7, 12)); // Sunday.
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(WEEKDAY(A1)=1,WEEKDAY(A1,2)=7,WEEKDAY(A1,3)=6,WEEKDAY(A1,17)=1)',
            apply: { fill: '#weekday' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=WEEKDAY(A1,99)=1',
            apply: { fill: '#weekday-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=WEEKDAY(A1,)=1',
            apply: { fill: '#weekday-omitted-return-type' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#weekday');
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBe('#weekday-omitted-return-type');
  });

  it('formula rules evaluate WEEKNUM and ISOWEEKNUM date operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, dateSerial(2026, 1, 1));
    s = seedNumber(s, 1, 0, dateSerial(2026, 1, 4));
    s = seedNumber(s, 2, 0, dateSerial(2026, 1, 5));
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(WEEKNUM(A1)=1,WEEKNUM(A2)=2,WEEKNUM(A3,2)=2)',
            apply: { fill: '#weeknum' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=AND(ISOWEEKNUM(A1)=1,ISOWEEKNUM(A3)=2,WEEKNUM(A3,21)=2)',
            apply: { fill: '#isoweeknum' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
            formula: '=WEEKNUM(A1,99)=1',
            apply: { fill: '#weeknum-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=WEEKNUM(A1,)=1',
            apply: { fill: '#weeknum-omitted-return-type' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#weeknum');
    expect(overlay.get('0:1:1')?.fill).toBe('#isoweeknum');
    expect(overlay.get('0:2:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBe('#weeknum-omitted-return-type');
  });

  it('formula rules evaluate TODAY/NOW volatile date operands', () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date(Date.UTC(2026, 6, 12, 12, 0, 0)));
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, dateSerial(2026, 7, 11));
    s = seedNumber(s, 1, 0, dateSerial(2026, 7, 12));
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 1, c1: 1 },
            formula: '=A1<TODAY()',
            apply: { fill: '#before-today' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=AND(NOW()>TODAY(),NOW()<TODAY()+1)',
            apply: { fill: '#now-within-today' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#before-today');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBe('#now-within-today');
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
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=TRUE()',
            apply: { fill: '#true-function' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=IF(A1>5,,TRUE)',
            apply: { fill: '#if-omitted-boolean' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#if');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
    expect(overlay.get('0:2:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBe('#true-function');
    expect(overlay.get('0:0:4')?.fill).toBeUndefined();
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
