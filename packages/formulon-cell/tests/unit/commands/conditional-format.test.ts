import { describe, expect, it } from 'vitest';
import {
  addConditionalRule,
  applyConditionalPresetAction,
  clearConditionalRules,
  clearConditionalRulesInRange,
  clearConditionalRulesOnSheet,
  conditionalRulesForRange,
  listConditionalRules,
  removeConditionalRuleAt,
} from '../../../src/commands/conditional-format.js';
import { type ConditionalRule, createSpreadsheetStore } from '../../../src/store/store.js';

const range = (r0: number, c0: number, r1: number, c1: number) =>
  ({ sheet: 0, r0, c0, r1, c1 }) as const;
const rangeOnSheet = (
  sheet: number,
  r0: number,
  c0: number,
  r1: number,
  c1: number,
): ConditionalRule['range'] => ({ sheet, r0, c0, r1, c1 });

const cellRule = (
  id: number,
  r: ConditionalRule['range'] = range(0, 0, 1, 1),
): ConditionalRule => ({
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

  it('clears rules from one sheet without deleting rules on other sheets', () => {
    const store = createSpreadsheetStore();
    const sheet0 = cellRule(1, rangeOnSheet(0, 0, 0, 0, 0));
    const sheet1 = cellRule(2, rangeOnSheet(1, 0, 0, 0, 0));

    addConditionalRule(store, sheet0);
    addConditionalRule(store, sheet1);

    clearConditionalRulesOnSheet(store, 0);
    expect(listConditionalRules(store.getState())).toEqual([sheet1]);

    expect(applyConditionalPresetAction(store, 'clear-sheet', rangeOnSheet(1, 2, 2, 3, 3))).toBe(
      true,
    );
    expect(listConditionalRules(store.getState())).toEqual([]);
  });

  it('applies Excel-style conditional-format preset actions to the active range', () => {
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      selection: {
        ...s.selection,
        range: range(2, 1, 5, 3),
      },
    }));

    expect(applyConditionalPresetAction(store, 'data-green')).toBe(true);
    expect(applyConditionalPresetAction(store, 'data-solid-green')).toBe(true);
    expect(applyConditionalPresetAction(store, 'scale-ryg')).toBe(true);
    expect(applyConditionalPresetAction(store, 'icons-arrows5')).toBe(true);
    expect(applyConditionalPresetAction(store, 'icons-symbols3')).toBe(true);
    expect(applyConditionalPresetAction(store, 'icons-bars5')).toBe(true);
    expect(applyConditionalPresetAction(store, 'top10-percent')).toBe(true);
    expect(applyConditionalPresetAction(store, 'duplicates')).toBe(true);

    const rules = listConditionalRules(store.getState());
    expect(rules.map((rule) => rule.kind)).toEqual([
      'data-bar',
      'data-bar',
      'color-scale',
      'icon-set',
      'icon-set',
      'icon-set',
      'top-bottom',
      'duplicates',
    ]);
    expect(rules[0]).toMatchObject({
      kind: 'data-bar',
      range: range(2, 1, 5, 3),
      color: '#63a95c',
      gradient: true,
      showValue: true,
    });
    expect(rules[1]).toMatchObject({
      kind: 'data-bar',
      range: range(2, 1, 5, 3),
      color: '#70ad47',
      gradient: false,
      showValue: true,
    });
    expect(rules[2]).toMatchObject({
      kind: 'color-scale',
      stops: ['#f8696b', '#ffeb84', '#63be7b'],
    });
    expect(rules[3]).toMatchObject({
      kind: 'icon-set',
      icons: 'arrows5',
      showValue: true,
      thresholds: [
        { kind: 'percent', value: 20 },
        { kind: 'percent', value: 40 },
        { kind: 'percent', value: 60 },
        { kind: 'percent', value: 80 },
      ],
    });
    expect(rules[4]).toMatchObject({
      kind: 'icon-set',
      icons: 'symbols3',
      showValue: true,
      thresholds: [
        { kind: 'percent', value: 100 / 3 },
        { kind: 'percent', value: 200 / 3 },
      ],
    });
    expect(rules[5]).toMatchObject({ kind: 'icon-set', icons: 'bars5', showValue: true });
    expect(rules[6]).toMatchObject({ kind: 'top-bottom', mode: 'top', n: 10, percent: true });
    expect(rules[7]).toMatchObject({ kind: 'duplicates', apply: { fill: '#ffc7ce' } });

    expect(applyConditionalPresetAction(store, 'clear-selection')).toBe(true);
    expect(listConditionalRules(store.getState())).toEqual([]);
  });
});
