import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  clearWatchedCells,
  isWatched,
  recordWatchesChange,
  setWatchWindowOpen,
  toggleWatchCell,
  unwatchCell,
  watchCell,
  watchRange,
  watchRanges,
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

  it('adds all cells in a range without duplicating existing watches', () => {
    const store = createSpreadsheetStore();

    watchCell(store, { sheet: 0, row: 0, col: 0 });
    watchRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });

    expect(store.getState().watch.watches).toEqual([
      { sheet: 0, row: 0, col: 0 },
      { sheet: 0, row: 0, col: 1 },
      { sheet: 0, row: 1, col: 0 },
      { sheet: 0, row: 1, col: 1 },
    ]);
  });

  it('adds multiple selected ranges in order', () => {
    const store = createSpreadsheetStore();

    watchRanges(store, [
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      { sheet: 1, r0: 2, c0: 2, r1: 2, c1: 2 },
    ]);

    expect(store.getState().watch.watches).toEqual([
      { sheet: 0, row: 0, col: 0 },
      { sheet: 0, row: 0, col: 1 },
      { sheet: 1, row: 2, col: 2 },
    ]);
  });

  it('ignores huge ranges instead of materializing millions of watches', () => {
    const store = createSpreadsheetStore();

    watchRange(store, { sheet: 0, r0: 0, c0: 2, r1: 1048575, c1: 2 });

    expect(store.getState().watch.watches).toEqual([]);
  });

  it('does not record history when a huge watch range is ignored', () => {
    const store = createSpreadsheetStore();
    const history = new History();

    recordWatchesChange(history, store, () => {
      watchRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1048575, c1: 0 });
    });

    expect(store.getState().watch.watches).toEqual([]);
    expect(history.canUndo()).toBe(false);
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

  it('records watch list changes as undoable actions', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 };

    recordWatchesChange(history, store, () => {
      watchRange(store, range);
    });

    expect(store.getState().watch.watches).toEqual([
      { sheet: 0, row: 0, col: 0 },
      { sheet: 0, row: 0, col: 1 },
    ]);
    expect(history.canUndo()).toBe(true);

    history.undo();
    expect(store.getState().watch.watches).toEqual([]);

    history.redo();
    expect(store.getState().watch.watches).toEqual([
      { sheet: 0, row: 0, col: 0 },
      { sheet: 0, row: 0, col: 1 },
    ]);

    recordWatchesChange(history, store, () => {
      clearWatchedCells(store);
    });
    expect(store.getState().watch.watches).toEqual([]);

    history.undo();
    expect(store.getState().watch.watches).toEqual([
      { sheet: 0, row: 0, col: 0 },
      { sheet: 0, row: 0, col: 1 },
    ]);
  });

  it('opens and closes the watch window flag', () => {
    const store = createSpreadsheetStore();

    setWatchWindowOpen(store, true);
    expect(store.getState().ui.watchPanelOpen).toBe(true);

    setWatchWindowOpen(store, false);
    expect(store.getState().ui.watchPanelOpen).toBe(false);
  });
});
