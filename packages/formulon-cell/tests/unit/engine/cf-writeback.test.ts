import { describe, expect, it, vi } from 'vitest';
import {
  conditionalRuleToEngineInput,
  syncConditionalRulesToEngine,
} from '../../../src/engine/cf-writeback.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import type { ConditionalRule } from '../../../src/store/types.js';

const range = { sheet: 0, r0: 1, c0: 2, r1: 3, c1: 4 };
const sqref = [{ firstRow: 1, firstCol: 2, lastRow: 3, lastCol: 4 }];

describe('conditionalRuleToEngineInput (H-13)', () => {
  it('translates a cell-value rule with operator + threshold', () => {
    const rule: ConditionalRule = {
      kind: 'cell-value',
      range,
      op: '>',
      a: 10,
      apply: { bold: true },
    };
    expect(conditionalRuleToEngineInput(rule)).toEqual({
      sqref,
      type: 1,
      op: 5,
      formula1: '10',
    });
  });

  it('emits formula2 for a between rule', () => {
    const rule: ConditionalRule = {
      kind: 'cell-value',
      range,
      op: 'between',
      a: 1,
      b: 9,
      apply: {},
    };
    const input = conditionalRuleToEngineInput(rule);
    expect(input).toMatchObject({ type: 1, op: 6, formula1: '1', formula2: '9' });
  });

  it('translates an =-prefixed formula rule to an expression rule', () => {
    const rule: ConditionalRule = {
      kind: 'formula',
      range,
      formula: '=A1>B1',
      apply: {},
    };
    expect(conditionalRuleToEngineInput(rule)).toEqual({
      sqref,
      type: 0,
      formula1: 'A1>B1',
    });
  });

  it('skips a comparator-prefix formula rule (not a standalone expression)', () => {
    const rule: ConditionalRule = { kind: 'formula', range, formula: '>10', apply: {} };
    expect(conditionalRuleToEngineInput(rule)).toBeNull();
  });

  it('translates text / blanks / errors / duplicate predicate rules', () => {
    expect(
      conditionalRuleToEngineInput({ kind: 'text-contains', range, text: 'hi', apply: {} }),
    ).toEqual({ sqref, type: 7, text: 'hi' });
    expect(conditionalRuleToEngineInput({ kind: 'blanks', range, apply: {} })).toEqual({
      sqref,
      type: 11,
    });
    expect(conditionalRuleToEngineInput({ kind: 'errors', range, apply: {} })).toEqual({
      sqref,
      type: 13,
    });
    expect(conditionalRuleToEngineInput({ kind: 'duplicates', range, apply: {} })).toEqual({
      sqref,
      type: 16,
    });
    expect(conditionalRuleToEngineInput({ kind: 'unique', range, apply: {} })).toEqual({
      sqref,
      type: 17,
    });
  });

  it('translates top-bottom and average rules', () => {
    expect(
      conditionalRuleToEngineInput({
        kind: 'top-bottom',
        range,
        mode: 'bottom',
        n: 5,
        percent: true,
        apply: {},
      }),
    ).toEqual({ sqref, type: 5, rank: 5, percent: true, bottom: true });
    expect(
      conditionalRuleToEngineInput({
        kind: 'average',
        range,
        mode: 'equal-or-above',
        apply: {},
      }),
    ).toEqual({ sqref, type: 6, aboveAverage: true, equalAverage: true });
  });

  it('returns null for visual and date-occurring rules the engine cannot author', () => {
    expect(
      conditionalRuleToEngineInput({ kind: 'color-scale', range, stops: ['#fff', '#000'] }),
    ).toBeNull();
    expect(conditionalRuleToEngineInput({ kind: 'data-bar', range, color: '#0078d4' })).toBeNull();
    expect(
      conditionalRuleToEngineInput({
        kind: 'date-occurring',
        range,
        period: 'today',
        apply: {},
      }),
    ).toBeNull();
  });
});

describe('syncConditionalRulesToEngine (H-13)', () => {
  const fakeWb = (mutate: boolean) => {
    const added: Array<{ sheet: number; type: number }> = [];
    const wb = {
      capabilities: { conditionalFormatMutate: mutate },
      addConditionalFormat: vi.fn((sheet: number, rule: { type: number }) => {
        added.push({ sheet, type: rule.type });
        return true;
      }),
    } as unknown as WorkbookHandle;
    return { wb, added };
  };

  it('writes representable rules and counts skips for the rest', () => {
    const { wb, added } = fakeWb(true);
    const rules: ConditionalRule[] = [
      { kind: 'cell-value', range, op: '>', a: 10, apply: {} },
      { kind: 'data-bar', range, color: '#0078d4' }, // visual → skipped
      { kind: 'text-contains', range, text: 'x', apply: {} },
      { kind: 'cell-value', range: { ...range, sheet: 1 }, op: '<', a: 0, apply: {} }, // other sheet
    ];
    const result = syncConditionalRulesToEngine(wb, rules, 0);
    expect(result).toEqual({ written: 2, skipped: 1 });
    expect(added).toEqual([
      { sheet: 0, type: 1 },
      { sheet: 0, type: 7 },
    ]);
  });

  it('is a no-op when the engine cannot author conditional formats', () => {
    const { wb, added } = fakeWb(false);
    const result = syncConditionalRulesToEngine(
      wb,
      [{ kind: 'cell-value', range, op: '>', a: 10, apply: {} }],
      0,
    );
    expect(result).toEqual({ written: 0, skipped: 0 });
    expect(added).toHaveLength(0);
  });
});
