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
