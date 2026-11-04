import { describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { recordWatchesChange, watchCell } from '../../../src/commands/watch.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('recordWatchesChange', () => {
  it('does not snapshot the watch list when there is no history to record into', () => {
    const store = createSpreadsheetStore();
    const getState = vi.spyOn(store, 'getState');
    expect(recordWatchesChange(null, store, () => 7)).toBe(7);
    expect(getState).not.toHaveBeenCalled();
  });

  it('records an undo step when history is present', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    recordWatchesChange(history, store, () => watchCell(store, { sheet: 0, row: 0, col: 0 }));
    expect(store.getState().watch.watches).toHaveLength(1);
    expect(history.undo()).toBe(true);
    expect(store.getState().watch.watches).toHaveLength(0);
  });
});
