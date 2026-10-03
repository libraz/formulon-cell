import { beforeEach, describe, expect, it } from 'vitest';
import {
  bumpDecimals,
  cycleCurrency,
  cyclePercent,
  setNumFmt,
} from '../../../../src/commands/format.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../../src/store/store.js';
import { effectiveFmtAt, setRange } from './fixtures.js';

describe('setNumFmt / cycleCurrency / cyclePercent', () => {
  let store: SpreadsheetStore;

  beforeEach(() => {
    store = createSpreadsheetStore();
    setRange(store, 0, 0, 0, 0);
  });

  it('setNumFmt installs the supplied format', () => {
    setNumFmt(store.getState(), store, { kind: 'fixed', decimals: 3 });
    expect(effectiveFmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'fixed', decimals: 3 });
  });

  it('cycleCurrency turns currency on when none is set', () => {
    cycleCurrency(store.getState(), store);
    expect(effectiveFmtAt(store, 0, 0)?.numFmt).toEqual({
      kind: 'currency',
      decimals: 2,
      symbol: '$',
    });
  });

  it('cycleCurrency uses the active locale currency symbol', () => {
    cycleCurrency(store.getState(), store, 'ja');
    expect(effectiveFmtAt(store, 0, 0)?.numFmt).toEqual({
      kind: 'currency',
      decimals: 2,
      symbol: '¥',
    });
  });

  it('cycleCurrency clears back to general when at least one cell is currency', () => {
    cycleCurrency(store.getState(), store); // on
    cycleCurrency(store.getState(), store); // off
    expect(effectiveFmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'general' });
  });

  it('cyclePercent toggles percent on / off', () => {
    cyclePercent(store.getState(), store);
    expect(effectiveFmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'percent', decimals: 0 });
    cyclePercent(store.getState(), store);
    expect(effectiveFmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'general' });
  });
});

describe('bumpDecimals', () => {
  let store: SpreadsheetStore;

  beforeEach(() => {
    store = createSpreadsheetStore();
    setRange(store, 0, 0, 0, 0);
  });

  it('promotes a general cell to fixed:2 on +1', () => {
    bumpDecimals(store.getState(), store, 1);
    expect(effectiveFmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'fixed', decimals: 2 });
  });

  it('does nothing on -1 when the cell is general', () => {
    bumpDecimals(store.getState(), store, -1);
    expect(effectiveFmtAt(store, 0, 0)?.numFmt).toBeUndefined();
  });

  it('walks fixed decimals up and down with clamping', () => {
    setNumFmt(store.getState(), store, { kind: 'fixed', decimals: 0 });
    bumpDecimals(store.getState(), store, -1);
    expect(effectiveFmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'fixed', decimals: 0 });
    for (let i = 0; i < 12; i += 1) bumpDecimals(store.getState(), store, 1);
    expect(effectiveFmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'fixed', decimals: 10 });
  });

  it('preserves currency symbol while bumping decimals', () => {
    setNumFmt(store.getState(), store, { kind: 'currency', decimals: 2, symbol: '€' });
    bumpDecimals(store.getState(), store, 1);
    expect(effectiveFmtAt(store, 0, 0)?.numFmt).toEqual({
      kind: 'currency',
      decimals: 3,
      symbol: '€',
    });
  });

  it('walks percent decimals', () => {
    setNumFmt(store.getState(), store, { kind: 'percent', decimals: 0 });
    bumpDecimals(store.getState(), store, 1);
    expect(effectiveFmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'percent', decimals: 1 });
  });
});
