import { describe, expect, it, vi } from 'vitest';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('sheet object mutators no-ops', () => {
  it('clearing an absent sparkline keeps state identity and does not notify', () => {
    const store = createSpreadsheetStore();
    const before = store.getState();
    const fn = vi.fn();
    store.subscribe(fn);
    mutators.clearSparkline(store, { sheet: 0, row: 1, col: 1 });
    mutators.setSparkline(store, { sheet: 0, row: 1, col: 1 }, null);
    expect(store.getState()).toBe(before);
    expect(fn).not.toHaveBeenCalled();
  });

  it('removing an out-of-range conditional rule index keeps state identity', () => {
    const store = createSpreadsheetStore();
    const before = store.getState();
    mutators.removeConditionalRuleAt(store, 3);
    expect(store.getState()).toBe(before);
  });
});
