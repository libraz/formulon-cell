import { describe, expect, it } from 'vitest';
import { evaluateConditional } from '../../../src/render/conditional.js';
import type { ConditionalRule, State } from '../../../src/store/store.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import { withConditionalRules } from './conditional-fixtures.js';

const range = { sheet: 0, r0: 0, c0: 0, r1: 199, c1: 0 };
const apply = { fill: '#abcdef' };

const rules: Record<string, ConditionalRule> = {
  'color-scale': { kind: 'color-scale', range, stops: ['#ff0000', '#00ff00'] },
  'data-bar': { kind: 'data-bar', range, color: '#0000ff' },
  'icon-set': { kind: 'icon-set', range, icons: '3Arrows' as never },
  'top-bottom': { kind: 'top-bottom', range, mode: 'top', n: 50, apply },
  average: { kind: 'average', range, mode: 'above', apply },
  duplicates: { kind: 'duplicates', range, apply },
};

function fixture(rule: ConditionalRule): { state: State; reads: () => number } {
  const base = createSpreadsheetStore().getState();
  const cells = new Map(base.data.cells);
  for (let r = 0; r <= 199; r += 1) {
    cells.set(`0:${r}:0`, { value: { kind: 'number', value: r % 37 }, formula: null });
  }
  let reads = 0;
  const read = cells.get.bind(cells);
  cells.get = (key: string) => {
    reads += 1;
    return read(key);
  };
  const state = withConditionalRules({ ...base, data: { ...base.data, cells } }, [rule]);
  return { state, reads: () => reads };
}

const window = (r0: number) => ({ r0, c0: 0, r1: r0 + 9, c1: 0 });

describe('conditional rule-wide stats', () => {
  it.each(Object.keys(rules))('%s: scrolling reuses the rule stats', (name) => {
    const { state, reads } = fixture(rules[name] as ConditionalRule);
    const first = evaluateConditional(state, null, '', window(0));
    const afterFirst = reads();
    const second = evaluateConditional(state, null, '', window(10));
    const scrollReads = reads() - afterFirst;

    // The second window touches its own 10 cells, not the whole 200-cell rule.
    expect(scrollReads).toBeLessThanOrEqual(30);
    const full = evaluateConditional(state);
    for (const [key, overlay] of second) expect(overlay).toEqual(full.get(key));
    for (const [key, overlay] of first) expect(overlay).toEqual(full.get(key));
    expect([...second.keys()].every((k) => Number(k.split(':')[1]) >= 10)).toBe(true);
  });
});
