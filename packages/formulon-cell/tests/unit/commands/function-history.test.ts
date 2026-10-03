import { describe, expect, it } from 'vitest';
import { allFunctionNames } from '../../../src/commands/function-categories.js';
import {
  getRecentFunctions,
  recordRecentFunction,
  subscribeRecentFunctions,
} from '../../../src/commands/function-history.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('function history', () => {
  it('keeps a per-store MRU list and notifies listeners only when its order changes', () => {
    const store = createSpreadsheetStore();
    const otherStore = createSpreadsheetStore();
    const notifications: string[][] = [];
    let latestSnapshot: readonly string[] | undefined;
    const unsubscribe = subscribeRecentFunctions(store, (names) => {
      notifications.push([...names]);
      latestSnapshot = names;
    });

    expect(recordRecentFunction(store, 'SUM')).toBe(true);
    expect(recordRecentFunction(store, 'SUM')).toBe(false);
    expect(recordRecentFunction(store, 'AVERAGE')).toBe(true);
    expect(recordRecentFunction(store, 'SUM')).toBe(true);
    expect(getRecentFunctions(store)).toEqual(['SUM', 'AVERAGE']);
    expect(Object.isFrozen(getRecentFunctions(store))).toBe(true);
    expect(Object.isFrozen(latestSnapshot)).toBe(true);
    expect(getRecentFunctions(otherStore)).toEqual([]);
    expect(notifications).toEqual([['SUM'], ['AVERAGE', 'SUM'], ['SUM', 'AVERAGE']]);

    unsubscribe();
    expect(recordRecentFunction(store, 'COUNT')).toBe(true);
    expect(notifications).toHaveLength(3);
  });

  it('keeps a broken subscriber from interrupting other listeners or recording', () => {
    const store = createSpreadsheetStore();
    const received: string[][] = [];
    subscribeRecentFunctions(store, () => {
      throw new Error('subscriber failed');
    });
    subscribeRecentFunctions(store, (names) => received.push([...names]));

    expect(() => recordRecentFunction(store, 'SUM')).not.toThrow();
    expect(getRecentFunctions(store)).toEqual(['SUM']);
    expect(received).toEqual([['SUM']]);
  });

  it('rejects unknown and non-canonical function names', () => {
    const store = createSpreadsheetStore();
    expect(recordRecentFunction(store, 'sum')).toBe(false);
    expect(recordRecentFunction(store, 'NOT_A_FUNCTION')).toBe(false);
    expect(getRecentFunctions(store)).toEqual([]);
  });

  it('keeps only the twelve most recent known functions', () => {
    const store = createSpreadsheetStore();
    const names = allFunctionNames().slice(0, 13);
    for (const name of names) recordRecentFunction(store, name);

    expect(getRecentFunctions(store)).toEqual(names.slice(1).reverse());
    expect(getRecentFunctions(store)).toHaveLength(12);
  });

  it('records engine-only functions when the current catalog authorizes them', () => {
    const store = createSpreadsheetStore();
    const liveNames = new Set(['LIVE_ONLY_FN']);

    expect(recordRecentFunction(store, 'LIVE_ONLY_FN')).toBe(false);
    expect(recordRecentFunction(store, 'LIVE_ONLY_FN', liveNames)).toBe(true);
    expect(getRecentFunctions(store, liveNames)).toEqual(['LIVE_ONLY_FN']);
    expect(getRecentFunctions(store, new Set(['SUM']))).toEqual([]);
    expect(getRecentFunctions(store, liveNames)).toEqual(['LIVE_ONLY_FN']);
  });

  it('keeps stale MRU entries for a later workbook catalog', () => {
    const store = createSpreadsheetStore();
    expect(recordRecentFunction(store, 'LIVE_ONLY_FN', new Set(['LIVE_ONLY_FN']))).toBe(true);
    expect(recordRecentFunction(store, 'SUM', new Set(['SUM']))).toBe(true);
    expect(getRecentFunctions(store, new Set(['LIVE_ONLY_FN', 'SUM']))).toEqual([
      'SUM',
      'LIVE_ONLY_FN',
    ]);
    expect(getRecentFunctions(store, new Set(['LIVE_ONLY_FN']))).toEqual(['LIVE_ONLY_FN']);
    expect(getRecentFunctions(store, new Set(['SUM']))).toEqual(['SUM']);
  });

  it('keeps engine-only entries out of the static fallback view', () => {
    const store = createSpreadsheetStore();
    expect(recordRecentFunction(store, 'LIVE_ONLY_FN', new Set(['LIVE_ONLY_FN']))).toBe(true);
    expect(recordRecentFunction(store, 'SUM')).toBe(true);
    expect(getRecentFunctions(store)).toEqual(['SUM']);
    expect(getRecentFunctions(store, new Set(['LIVE_ONLY_FN']))).toEqual(['LIVE_ONLY_FN']);
  });
});
