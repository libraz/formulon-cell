import type { ConditionalRule, SpreadsheetStore, State } from '../store/store.js';
import { mutators } from '../store/store.js';

export function listConditionalRules(state: State): readonly ConditionalRule[] {
  return state.conditional.rules;
}

export function addConditionalRule(store: SpreadsheetStore, rule: ConditionalRule): void {
  mutators.addConditionalRule(store, rule);
}

export function removeConditionalRuleAt(store: SpreadsheetStore, index: number): void {
  mutators.removeConditionalRuleAt(store, index);
}

export function clearConditionalRules(store: SpreadsheetStore): void {
  mutators.clearConditionalRules(store);
}

export function clearConditionalRulesInRange(
  store: SpreadsheetStore,
  range: ConditionalRule['range'],
): void {
  mutators.clearConditionalRulesInRange(store, range);
}

export function conditionalRulesForRange(
  state: State,
  range: ConditionalRule['range'],
): readonly ConditionalRule[] {
  return state.conditional.rules.filter((rule) => rangesIntersect(rule.range, range));
}

const rangesIntersect = (a: ConditionalRule['range'], b: ConditionalRule['range']): boolean =>
  a.sheet === b.sheet && !(a.r1 < b.r0 || a.r0 > b.r1 || a.c1 < b.c0 || a.c0 > b.c1);
