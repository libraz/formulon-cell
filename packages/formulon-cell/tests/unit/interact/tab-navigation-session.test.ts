import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import {
  attachTabNavigationSession,
  clearTabReturnOrigin,
  resetTabNavigationSession,
  setTabNavigationMemory,
  tabNavigationMemoryFor,
  withTabNavigationSelectionMutation,
} from '../../../src/interact/tab-navigation-session.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const addr = (row: number, col: number) => ({ sheet: 0, row, col });
const range = (r0: number, c0: number, r1: number, c1: number) => ({
  sheet: 0,
  r0,
  c0,
  r1,
  c1,
});

describe('Tab navigation session', () => {
  let store: SpreadsheetStore;
  let detach: () => void;

  beforeEach(() => {
    store = createSpreadsheetStore();
    mutators.setActive(store, addr(0, 0));
    detach = attachTabNavigationSession(store);
    setTabNavigationMemory(store, { laneRow: 0, returnOrigin: addr(0, 0) });
  });

  afterEach(() => detach());

  it('preserves memory when a programmatic reselection is semantically unchanged', () => {
    mutators.setActive(store, addr(0, 0));

    expect(tabNavigationMemoryFor(store)).toEqual({ laneRow: 0, returnOrigin: addr(0, 0) });
  });

  it('clears memory after a genuinely different selection', () => {
    mutators.setActive(store, addr(2, 3));

    expect(tabNavigationMemoryFor(store)).toEqual({ laneRow: null, returnOrigin: null });
  });

  it('guards a session-owned selection move from its own reset listener', () => {
    withTabNavigationSelectionMutation(store, () => mutators.setActive(store, addr(0, 1)));

    expect(tabNavigationMemoryFor(store)).toEqual({ laneRow: 0, returnOrigin: addr(0, 0) });
  });

  it('preserves memory through cell-data writes', () => {
    store.setState((state) => {
      const cells = new Map(state.data.cells);
      cells.set('0:0:0', { value: { kind: 'text', value: 'x' }, formula: null });
      return { ...state, data: { ...state.data, cells } };
    });

    expect(tabNavigationMemoryFor(store)).toEqual({ laneRow: 0, returnOrigin: addr(0, 0) });
  });

  it('resets on sheet, row visibility, column visibility, and merge changes', () => {
    store.setState((state) => ({ ...state, data: { ...state.data, sheetIndex: 1 } }));
    expect(tabNavigationMemoryFor(store)).toEqual({ laneRow: null, returnOrigin: null });
    setTabNavigationMemory(store, { laneRow: 0, returnOrigin: addr(0, 0) });

    store.setState((state) => ({
      ...state,
      layout: { ...state.layout, hiddenRows: new Set([1]) },
    }));
    expect(tabNavigationMemoryFor(store)).toEqual({ laneRow: null, returnOrigin: null });
    setTabNavigationMemory(store, { laneRow: 0, returnOrigin: addr(0, 0) });

    store.setState((state) => ({
      ...state,
      layout: { ...state.layout, hiddenCols: new Set([1]) },
    }));
    expect(tabNavigationMemoryFor(store)).toEqual({ laneRow: null, returnOrigin: null });
    setTabNavigationMemory(store, { laneRow: 0, returnOrigin: addr(0, 0) });

    mutators.mergeRange(store, range(0, 1, 1, 1));
    expect(tabNavigationMemoryFor(store)).toEqual({ laneRow: null, returnOrigin: null });
  });

  it('clears only the Return origin after reverse Tab and supports an explicit reset', () => {
    clearTabReturnOrigin(store);
    expect(tabNavigationMemoryFor(store)).toEqual({ laneRow: 0, returnOrigin: null });

    resetTabNavigationSession(store);
    expect(tabNavigationMemoryFor(store)).toEqual({ laneRow: null, returnOrigin: null });
  });

  it('disposes its listener and registry entry', () => {
    detach();
    detach = () => {};

    expect(tabNavigationMemoryFor(store)).toBeNull();
    mutators.setActive(store, addr(1, 1));
    expect(tabNavigationMemoryFor(store)).toBeNull();
  });
});
