import { describe, expect, it, vi } from 'vitest';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe.each([
  ['setFillPreview', 'fillPreview'],
  ['setMovePreview', 'movePreview'],
] as const)('%s', (name, key) => {
  it('does not notify when an equal range is set again', () => {
    const store = createSpreadsheetStore();
    mutators[name](store, { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 });
    const spy = vi.fn();
    store.subscribe(spy);
    mutators[name](store, { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 });
    expect(spy).not.toHaveBeenCalled();
    mutators[name](store, { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 2 });
    expect(spy).toHaveBeenCalledTimes(1);
    expect(store.getState().ui[key]?.r1).toBe(3);
    mutators[name](store, null);
    mutators[name](store, null);
    expect(spy).toHaveBeenCalledTimes(2);
  });
});
