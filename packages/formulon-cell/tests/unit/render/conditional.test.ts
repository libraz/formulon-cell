import { afterEach, describe, expect, it } from 'vitest';
import type { CellValue } from '../../../src/engine/types.js';
import {
  _resetConditionalCache,
  evaluateConditional,
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
    };
    s = { ...s, conditional: { ...s.conditional, rules: [rule] } };
    const overlay = evaluateConditional(s);
    expect(overlay.get('0:0:0')?.iconSlot).toBe(0);
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
