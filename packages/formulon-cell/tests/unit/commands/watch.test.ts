import { describe, expect, it } from 'vitest';
import {
  clearWatchedCells,
  isWatched,
  setWatchWindowOpen,
  toggleWatchCell,
  unwatchCell,
  watchCell,
} from '../../../src/commands/watch.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('watch commands', () => {
  it('adds a watched cell and ignores duplicates through the store mutator', () => {
    const store = createSpreadsheetStore();
    const addr = { sheet: 0, row: 1, col: 2 };

    watchCell(store, addr);
    watchCell(store, addr);

    expect(isWatched(store, addr)).toBe(true);
    expect(store.getState().watch.watches).toEqual([addr]);
  });

  it('removes a watched cell', () => {
    const store = createSpreadsheetStore();
    const a1 = { sheet: 0, row: 0, col: 0 };
    const b2 = { sheet: 0, row: 1, col: 1 };

    watchCell(store, a1);
    watchCell(store, b2);
    unwatchCell(store, a1);

    expect(isWatched(store, a1)).toBe(false);
    expect(isWatched(store, b2)).toBe(true);
    expect(store.getState().watch.watches).toEqual([b2]);
  });

  it('toggles and clears watched cells', () => {
    const store = createSpreadsheetStore();
    const addr = { sheet: 1, row: 2, col: 3 };

    expect(toggleWatchCell(store, addr)).toBe(true);
    expect(isWatched(store, addr)).toBe(true);

    expect(toggleWatchCell(store, addr)).toBe(false);
    expect(isWatched(store, addr)).toBe(false);

    watchCell(store, addr);
    clearWatchedCells(store);
    expect(store.getState().watch.watches).toEqual([]);
  });

  it('opens and closes the watch window flag', () => {
    const store = createSpreadsheetStore();

    setWatchWindowOpen(store, true);
    expect(store.getState().ui.watchPanelOpen).toBe(true);

    setWatchWindowOpen(store, false);
    expect(store.getState().ui.watchPanelOpen).toBe(false);
  });
});
