import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import {
  addConditionalRule,
  conditionalRulesForRange,
  removeConditionalRuleAt,
} from '../../src/commands/conditional-format.js';
import { type MountedStubSheet, mountStubSheet } from '../test-utils/index.js';

describe('integration: conditional format rules', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
  });

  afterEach(() => sheet.dispose());

  it('addConditionalRule stores the rule against the right range', () => {
    const { instance } = sheet;
    addConditionalRule(instance.store, {
      kind: 'cell-value',
      range: { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 },
      op: '>',
      a: 10,
      apply: { fill: '#ffcccc' },
    });

    const all = instance.store.getState().conditional.rules;
    expect(all).toHaveLength(1);
    expect(all.at(0)?.kind).toBe('cell-value');
  });

  it('conditionalRulesForRange filters to rules intersecting the query range', () => {
    const { instance } = sheet;
    // Rule on A1:A5 only.
    addConditionalRule(instance.store, {
      kind: 'cell-value',
      range: { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 },
      op: '>',
      a: 5,
      apply: { color: '#f00' },
    });
    // Rule on B1:B5 only.
    addConditionalRule(instance.store, {
      kind: 'color-scale',
      range: { sheet: 0, r0: 0, c0: 1, r1: 4, c1: 1 },
      stops: ['#ffffff', '#0078d4'],
    });

    const aRules = conditionalRulesForRange(instance.store.getState(), {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 4,
      c1: 0,
    });
    expect(aRules).toHaveLength(1);
    expect(aRules.at(0)?.kind).toBe('cell-value');

    const overlap = conditionalRulesForRange(instance.store.getState(), {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 4,
      c1: 1,
    });
    expect(overlap).toHaveLength(2);
  });

  it('removeConditionalRuleAt removes by index', () => {
    const { instance } = sheet;
    addConditionalRule(instance.store, {
      kind: 'cell-value',
      range: { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 },
      op: '>',
      a: 1,
      apply: {},
    });
    addConditionalRule(instance.store, {
      kind: 'cell-value',
      range: { sheet: 0, r0: 5, c0: 0, r1: 9, c1: 0 },
      op: '<',
      a: 5,
      apply: {},
    });

    removeConditionalRuleAt(instance.store, 0);
    const all = instance.store.getState().conditional.rules;
    expect(all).toHaveLength(1);
    // The remaining rule is the second one we added.
    expect((all[0] as { op: string }).op).toBe('<');
  });

  it('removeConditionalRuleAt on an out-of-bounds index is a no-op', () => {
    const { instance } = sheet;
    addConditionalRule(instance.store, {
      kind: 'cell-value',
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      op: '=',
      a: 1,
      apply: {},
    });
    expect(() => removeConditionalRuleAt(instance.store, 99)).not.toThrow();
    expect(instance.store.getState().conditional.rules).toHaveLength(1);
  });
});
