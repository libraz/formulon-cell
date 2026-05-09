import { describe, expect, it } from 'vitest';
import {
  addConditionalRule,
  clearConditionalRules,
  clearConditionalRulesInRange,
  conditionalRulesForRange,
  listConditionalRules,
  removeConditionalRuleAt,
} from '../../../src/commands/conditional-format.js';
import { type ConditionalRule, createSpreadsheetStore } from '../../../src/store/store.js';

const range = (r0: number, c0: number, r1: number, c1: number) =>
  ({ sheet: 0, r0, c0, r1, c1 }) as const;

const cellRule = (id: number, r = range(0, 0, 1, 1)): ConditionalRule => ({
  kind: 'cell-value',
  range: r,
  op: '>',
  a: id,
  apply: { fill: '#fff2cc' },
});

describe('conditional-format commands', () => {
  it('adds, lists, and removes session conditional rules', () => {
    const store = createSpreadsheetStore();
    const a = cellRule(1);
    const b = cellRule(2, range(4, 0, 5, 0));

    addConditionalRule(store, a);
    addConditionalRule(store, b);

    expect(listConditionalRules(store.getState())).toEqual([a, b]);

    removeConditionalRuleAt(store, 0);
    expect(listConditionalRules(store.getState())).toEqual([b]);
  });

  it('filters and clears rules by intersecting range', () => {
    const store = createSpreadsheetStore();
    const a = cellRule(1, range(0, 0, 1, 1));
    const b = cellRule(2, range(5, 5, 6, 6));

    addConditionalRule(store, a);
    addConditionalRule(store, b);

    expect(conditionalRulesForRange(store.getState(), range(1, 1, 2, 2))).toEqual([a]);

    clearConditionalRulesInRange(store, range(0, 0, 1, 1));
    expect(listConditionalRules(store.getState())).toEqual([b]);

    clearConditionalRules(store);
    expect(listConditionalRules(store.getState())).toEqual([]);
  });
});
