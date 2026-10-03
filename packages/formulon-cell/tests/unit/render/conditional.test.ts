import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  _resetConditionalCache,
  evaluateConditional,
  iconSetSlotCount,
  iconSetSlotFor,
  topBottomThreshold,
} from '../../../src/render/conditional.js';
import type { ConditionalRule, State } from '../../../src/store/store.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import { cellValueRule, dateSerial, seedCell, seedNumber } from './conditional-fixtures.js';

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

  it('exposes data-bar appearance metadata and clears omitted borders', () => {
    const store = createSpreadsheetStore();
    let s = seedNumber(store.getState(), 0, 0, -10);
    s = seedNumber(s, 0, 1, 20);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'data-bar',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            color: '#70ad47',
          },
          {
            kind: 'data-bar',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
            color: '#0078d4',
            negativeColor: '#c00000',
            borderColor: '#1f1f1f',
            negativeBorderColor: '#7f0000',
            axisColor: '#404040',
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);
    expect(overlay.get('0:0:0')).toMatchObject({
      barColor: '#c00000',
      barBorderColor: '#7f0000',
      barAxisColor: '#404040',
      barAxisVisible: true,
    });
    // The higher-priority first rule has no border, so its explicit
    // undefined metadata must not inherit the lower rule's border.
    expect(overlay.get('0:0:1')).toMatchObject({
      barColor: '#70ad47',
      barBorderColor: undefined,
    });
  });

  it('uses whole-range engine lengths for middle and none axes', () => {
    const store = createSpreadsheetStore();
    let s = seedNumber(store.getState(), 0, 0, -10);
    s = seedNumber(s, 0, 1, -5);
    s = seedNumber(s, 0, 2, 20);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'data-bar',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 },
            color: '#0078d4',
            axisPosition: 'middle',
          },
        ],
      },
    };

    const middle = evaluateConditional(s);
    expect(middle.get('0:0:0')).toMatchObject({
      bar: 0,
      barAxis: 0.5,
      barDirection: 'left',
      barAxisVisible: true,
    });
    expect(middle.get('0:0:1')?.bar).toBeCloseTo(1 / 12);
    expect(middle.get('0:0:2')).toMatchObject({
      bar: 0.5,
      barAxis: 0.5,
      barDirection: 'right',
    });

    let positiveOnly = seedNumber(store.getState(), 0, 0, 10);
    positiveOnly = seedNumber(positiveOnly, 0, 1, 20);
    const positiveMiddle = evaluateConditional({
      ...positiveOnly,
      conditional: {
        rules: [
          {
            kind: 'data-bar',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
            color: '#0078d4',
            axisPosition: 'middle',
          },
        ],
      },
    });
    expect(positiveMiddle.get('0:0:1')?.barAxisVisible).toBe(true);

    let negativeOnly = seedNumber(store.getState(), 0, 0, -20);
    negativeOnly = seedNumber(negativeOnly, 0, 1, -10);
    const negativeMiddle = evaluateConditional({
      ...negativeOnly,
      conditional: {
        rules: [
          {
            kind: 'data-bar',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
            color: '#0078d4',
            axisPosition: 'middle',
          },
        ],
      },
    });
    expect(negativeMiddle.get('0:0:1')?.barAxisVisible).toBe(true);

    const none = evaluateConditional({
      ...s,
      conditional: {
        rules: [
          {
            kind: 'data-bar',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 },
            color: '#0078d4',
            axisPosition: 'none',
          },
        ],
      },
    });
    expect(none.get('0:0:0')).toMatchObject({
      bar: 0,
      barAxis: 0,
      barDirection: 'right',
      barAxisVisible: false,
    });
    expect(none.get('0:0:1')).toMatchObject({
      bar: 1 / 6,
      barAxis: 0,
      barDirection: 'right',
    });
    expect(none.get('0:0:2')).toMatchObject({
      bar: 1,
      barAxis: 0,
      barDirection: 'right',
    });
  });

  it('hides automatic axes without a mixed-sign population and mirrors none bars in RTL', () => {
    const store = createSpreadsheetStore();
    let s = seedNumber(store.getState(), 0, 0, -20);
    s = seedNumber(s, 0, 1, -10);
    const automatic = evaluateConditional({
      ...s,
      conditional: {
        rules: [
          {
            kind: 'data-bar',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
            color: '#0078d4',
          },
        ],
      },
    });
    expect(automatic.get('0:0:0')?.barAxisVisible).toBe(false);
    expect(automatic.get('0:0:1')?.barAxisVisible).toBe(false);

    const rtl = evaluateConditional({
      ...s,
      ui: { ...s.ui, rightToLeft: true },
      conditional: {
        rules: [
          {
            kind: 'data-bar',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
            color: '#0078d4',
            axisPosition: 'none',
          },
        ],
      },
    });
    expect(rtl.get('0:0:1')).toMatchObject({ barAxis: 1, barDirection: 'left' });
    expect(rtl.get('0:0:0')).toMatchObject({ barAxis: 1, barDirection: 'left' });
  });

  it('hides an automatic axis when explicit bounds cross zero without mixed-sign values', () => {
    const store = createSpreadsheetStore();
    let s = seedNumber(store.getState(), 0, 0, 10);
    s = seedNumber(s, 0, 1, 20);
    const overlay = evaluateConditional({
      ...s,
      conditional: {
        rules: [
          {
            kind: 'data-bar',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
            color: '#0078d4',
            min: { kind: 'number', value: -10 },
            max: { kind: 'number', value: 30 },
          },
        ],
      },
    });
    expect(overlay.get('0:0:0')).toMatchObject({
      barAxis: 0.25,
      barAxisVisible: false,
    });
    expect(overlay.get('0:0:1')?.barAxisVisible).toBe(false);
  });

  it('mirrors explicit data-bar direction and invalidates on sheet RTL changes', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, -10);
    s = seedNumber(s, 0, 1, 20);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'data-bar',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
            color: '#70ad47',
            direction: 'right-to-left',
          },
        ],
      },
    };

    const ltrOverlay = evaluateConditional(s);
    expect(ltrOverlay.get('0:0:0')?.barAxis).toBeCloseTo(2 / 3);
    expect(ltrOverlay.get('0:0:0')?.barDirection).toBe('right');
    expect(ltrOverlay.get('0:0:1')?.barAxis).toBeCloseTo(2 / 3);
    expect(ltrOverlay.get('0:0:1')?.barDirection).toBe('left');

    const rtlState: State = { ...s, ui: { ...s.ui, rightToLeft: true } };
    const rtlOverlay = evaluateConditional(rtlState);
    // Explicit right-to-left is independent of the sheet context, so it
    // remains mirrored after the sheet itself switches direction.
    expect(rtlOverlay.get('0:0:0')?.barAxis).toBeCloseTo(2 / 3);
    expect(rtlOverlay.get('0:0:0')?.barDirection).toBe('right');
    expect(rtlOverlay.get('0:0:1')?.barAxis).toBeCloseTo(2 / 3);
    expect(rtlOverlay.get('0:0:1')?.barDirection).toBe('left');
  });

  it('resolves and clamps numeric data-bar endpoints while honoring RTL direction', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    [0, 20, 50, 80, 100].forEach((value, col) => {
      s = seedNumber(s, 0, col, value);
    });
    s = {
      ...s,
      ui: { ...s.ui, rightToLeft: false },
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'data-bar',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 4 },
            color: '#70ad47',
            min: { kind: 'number', value: 20 },
            max: { kind: 'number', value: 80 },
            direction: 'right-to-left',
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')).toMatchObject({
      bar: 0,
      barAxis: 1,
      barDirection: 'left',
    });
    expect(overlay.get('0:0:1')?.bar).toBe(0);
    expect(overlay.get('0:0:2')?.bar).toBeCloseTo(0.5);
    expect(overlay.get('0:0:3')?.bar).toBe(1);
    expect(overlay.get('0:0:4')?.bar).toBe(1);
  });

  it('resolves percentile data-bar endpoints against the sorted range', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    [0, 20, 40, 60, 80].forEach((value, col) => {
      s = seedNumber(s, 0, col, value);
    });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'data-bar',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 4 },
            color: '#70ad47',
            min: { kind: 'percentile', value: 25 },
            max: { kind: 'percentile', value: 75 },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.bar).toBe(0);
    expect(overlay.get('0:0:1')?.bar).toBe(0);
    expect(overlay.get('0:0:2')?.bar).toBeCloseTo(0.5);
    expect(overlay.get('0:0:3')?.bar).toBe(1);
    expect(overlay.get('0:0:4')?.bar).toBe(1);
  });

  it.each([
    ['percent', 'percent', 'left-to-right', 0],
    ['percentile', 'percent', 'right-to-left', 0.2461538462],
    ['percent', 'percentile', 'right-to-left', 0],
    ['number', 'percent', 'context', 0.2032520325],
    ['number', 'number', 'right-to-left', 1 / 6],
    ['number', 'percentile', 'right-to-left', 0.3424657534],
    ['percent', 'number', 'context', 0],
    ['percentile', 'number', 'left-to-right', 0.2038216561],
    ['number', 'percentile', 'left-to-right', 0.3424657534],
    ['percentile', 'percent', 'context', 0.2461538462],
    ['percentile', 'percentile', 'context', 0.4],
  ] as const)(
    'covers data-bar endpoint kind and direction combinations (%s/%s/%s)',
    (minKind, maxKind, direction, expectedBar) => {
      const endpointValues = {
        number: { min: 20, max: 80 },
        percent: { min: 26, max: 74 },
        percentile: { min: 18, max: 66 },
      } as const;
      const store = createSpreadsheetStore();
      let s = store.getState();
      [10, 30, 90].forEach((value, col) => {
        s = seedNumber(s, 0, col, value);
      });
      s = {
        ...s,
        ui: { ...s.ui, rightToLeft: false },
        conditional: {
          ...s.conditional,
          rules: [
            {
              kind: 'data-bar',
              range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 },
              color: '#70ad47',
              min: { kind: minKind, value: endpointValues[minKind].min },
              max: { kind: maxKind, value: endpointValues[maxKind].max },
              direction,
            },
          ],
        },
      };

      const overlay = evaluateConditional(s);
      const mirrored = direction === 'right-to-left';
      expect(overlay.get('0:0:1')).toMatchObject({
        barAxis: mirrored ? 1 : 0,
        barDirection: mirrored ? 'left' : 'right',
      });
      expect(overlay.get('0:0:1')?.bar).toBeCloseTo(expectedBar, 8);
    },
  );

  it('keeps automatic data-bar zero baselines when endpoints are omitted', () => {
    const store = createSpreadsheetStore();
    let positive = store.getState();
    [10, 20].forEach((value, col) => {
      positive = seedNumber(positive, 0, col, value);
    });
    positive = {
      ...positive,
      conditional: {
        ...positive.conditional,
        rules: [
          {
            kind: 'data-bar',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
            color: '#70ad47',
          },
        ],
      },
    };
    const positiveOverlay = evaluateConditional(positive);
    expect(positiveOverlay.get('0:0:0')?.bar).toBeCloseTo(0.5);
    expect(positiveOverlay.get('0:0:1')?.bar).toBe(1);

    let negative = seedNumber(positive, 0, 0, -20);
    negative = seedNumber(negative, 0, 1, -10);
    const negativeOverlay = evaluateConditional(negative);
    expect(negativeOverlay.get('0:0:0')?.bar).toBe(1);
    expect(negativeOverlay.get('0:0:1')?.bar).toBeCloseTo(0.5);
  });

  it.each([10, -10, 0])('preserves automatic endpoints for equal-valued ranges (%s)', (value) => {
    let s = createSpreadsheetStore().getState();
    s = seedNumber(s, 0, 0, value);
    s = seedNumber(s, 0, 1, value);
    const rule: Extract<ConditionalRule, { kind: 'data-bar' }> = {
      kind: 'data-bar',
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      color: '#70ad47',
    };
    const implicit = evaluateConditional({ ...s, conditional: { rules: [rule] } });
    const explicit = evaluateConditional({
      ...s,
      conditional: { rules: [{ ...rule, min: { kind: 'min' }, max: { kind: 'max' } }] },
    });
    for (const overlay of [implicit, explicit]) {
      expect(overlay.get('0:0:0')?.bar).toBe(value === 0 ? 0 : 1);
      expect(overlay.get('0:0:1')?.bar).toBe(value === 0 ? 0 : 1);
    }
    expect(explicit).toEqual(implicit);
  });

  it('excludes non-finite values from data-bar bounds and overlays', () => {
    let s = createSpreadsheetStore().getState();
    [10, Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY, 20].forEach(
      (value, col) => {
        s = seedNumber(s, 0, col, value);
      },
    );
    const overlay = evaluateConditional({
      ...s,
      conditional: {
        rules: [
          { kind: 'data-bar', range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 4 }, color: '#70ad47' },
        ],
      },
    });
    expect(overlay.get('0:0:0')?.bar).toBe(0.5);
    expect(overlay.get('0:0:4')?.bar).toBe(1);
    expect(overlay.has('0:0:1')).toBe(false);
    expect(overlay.has('0:0:2')).toBe(false);
    expect(overlay.has('0:0:3')).toBe(false);
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

  it('suppresses icon-set output below an independent floor', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 5);
    s = seedNumber(s, 0, 1, 10);
    s = seedNumber(s, 0, 2, 20);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'icon-set',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 },
            icons: 'traffic3',
            floor: { kind: 'number', value: 10 },
            thresholds: [
              { kind: 'number', value: 15 },
              { kind: 'number', value: 25 },
            ],
          },
        ],
      },
    };
    const overlay = evaluateConditional(s);
    expect(overlay.get('0:0:0')).toBeUndefined();
    expect(overlay.get('0:0:1')?.iconSlot).toBe(0);
    expect(overlay.get('0:0:2')?.iconSlot).toBe(1);
  });

  it('honors strict floor and threshold comparisons for icon sets', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 14);
    s = seedNumber(s, 0, 1, 15);
    s = seedNumber(s, 0, 2, 20);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'icon-set',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 },
            icons: 'traffic3',
            floor: { kind: 'number', value: 15, gte: false },
            thresholds: [
              { kind: 'number', value: 20, gte: false },
              { kind: 'number', value: 40, gte: false },
            ],
          },
        ],
      },
    };
    const overlay = evaluateConditional(s);
    expect(overlay.get('0:0:0')).toBeUndefined();
    expect(overlay.get('0:0:1')).toBeUndefined();
    expect(overlay.get('0:0:2')?.iconSlot).toBe(0);
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
