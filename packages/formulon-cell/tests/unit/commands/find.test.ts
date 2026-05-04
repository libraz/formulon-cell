import { beforeEach, describe, expect, it } from 'vitest';
import {
  applySubstitution,
  findAll,
  findNext,
  replaceAll,
  replaceOne,
} from '../../../src/commands/find.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const seed = (
  store: SpreadsheetStore,
  cells: Array<{ row: number; col: number; value: string | number }>,
): void => {
  store.setState((s) => {
    const map = new Map(s.data.cells);
    for (const c of cells) {
      const addr = { sheet: 0, row: c.row, col: c.col };
      const value =
        typeof c.value === 'number'
          ? ({ kind: 'number', value: c.value } as const)
          : ({ kind: 'text', value: c.value } as const);
      map.set(addrKey(addr), { value, formula: null });
    }
    return { ...s, data: { ...s.data, cells: map } };
  });
};

describe('findAll', () => {
  let store: SpreadsheetStore;

  beforeEach(() => {
    store = createSpreadsheetStore();
  });

  it('returns empty when query is empty', () => {
    seed(store, [{ row: 0, col: 0, value: 'apple' }]);
    expect(findAll(store.getState(), { query: '' })).toEqual([]);
  });

  it('matches by substring (case-insensitive by default)', () => {
    seed(store, [
      { row: 0, col: 0, value: 'Apple' },
      { row: 1, col: 0, value: 'pineapple' },
      { row: 2, col: 0, value: 'banana' },
    ]);
    const got = findAll(store.getState(), { query: 'apple' });
    expect(got.map((m) => m.addr.row).sort()).toEqual([0, 1]);
  });

  it('respects caseSensitive', () => {
    seed(store, [
      { row: 0, col: 0, value: 'Apple' },
      { row: 1, col: 0, value: 'apple' },
    ]);
    const got = findAll(store.getState(), { query: 'apple', caseSensitive: true });
    expect(got).toHaveLength(1);
    expect(got[0]?.addr.row).toBe(1);
  });

  it('respects matchWhole — must equal entire cell', () => {
    seed(store, [
      { row: 0, col: 0, value: 'apple' },
      { row: 1, col: 0, value: 'pineapple' },
    ]);
    const got = findAll(store.getState(), { query: 'apple', matchWhole: true });
    expect(got).toHaveLength(1);
    expect(got[0]?.addr.row).toBe(0);
  });

  it('searches numeric values via formatted display', () => {
    seed(store, [{ row: 0, col: 0, value: 42 }]);
    expect(findAll(store.getState(), { query: '42' })).toHaveLength(1);
  });

  it('returns row-major-ordered results', () => {
    seed(store, [
      { row: 5, col: 0, value: 'x' },
      { row: 1, col: 3, value: 'x' },
      { row: 1, col: 1, value: 'x' },
    ]);
    const got = findAll(store.getState(), { query: 'x' });
    expect(got.map((m) => `${m.addr.row}:${m.addr.col}`)).toEqual(['1:1', '1:3', '5:0']);
  });

  it('ignores cells from other sheets', () => {
    store.setState((s) => {
      const map = new Map(s.data.cells);
      map.set(addrKey({ sheet: 1, row: 0, col: 0 }), {
        value: { kind: 'text', value: 'hit' },
        formula: null,
      });
      map.set(addrKey({ sheet: 0, row: 0, col: 0 }), {
        value: { kind: 'text', value: 'hit' },
        formula: null,
      });
      return { ...s, data: { ...s.data, cells: map } };
    });
    const got = findAll(store.getState(), { query: 'hit' });
    expect(got).toHaveLength(1);
    expect(got[0]?.addr.sheet).toBe(0);
  });
});

describe('findNext', () => {
  let store: SpreadsheetStore;

  beforeEach(() => {
    store = createSpreadsheetStore();
    seed(store, [
      { row: 0, col: 0, value: 'x' },
      { row: 2, col: 1, value: 'x' },
      { row: 5, col: 0, value: 'x' },
    ]);
  });

  it('returns null when there are no matches', () => {
    expect(findNext(store.getState(), { query: 'zz' }, null, 'next')).toBeNull();
  });

  it('returns the first match when called with no cursor (next)', () => {
    const got = findNext(store.getState(), { query: 'x' }, null, 'next');
    expect(got?.addr).toEqual({ sheet: 0, row: 0, col: 0 });
  });

  it('returns the last match when called with no cursor (prev)', () => {
    const got = findNext(store.getState(), { query: 'x' }, null, 'prev');
    expect(got?.addr).toEqual({ sheet: 0, row: 5, col: 0 });
  });

  it('walks forward past the cursor', () => {
    const from = { sheet: 0, row: 0, col: 0 };
    const got = findNext(store.getState(), { query: 'x' }, from, 'next');
    expect(got?.addr).toEqual({ sheet: 0, row: 2, col: 1 });
  });

  it('wraps around when no match remains in the chosen direction', () => {
    const lastMatch = { sheet: 0, row: 5, col: 0 };
    const got = findNext(store.getState(), { query: 'x' }, lastMatch, 'next');
    expect(got?.addr).toEqual({ sheet: 0, row: 0, col: 0 });
  });

  it('walks backward past the cursor', () => {
    const from = { sheet: 0, row: 5, col: 0 };
    const got = findNext(store.getState(), { query: 'x' }, from, 'prev');
    expect(got?.addr).toEqual({ sheet: 0, row: 2, col: 1 });
  });
});

describe('applySubstitution', () => {
  it('replaces all occurrences (case-insensitive default)', () => {
    expect(applySubstitution('Apple apple', { query: 'apple' }, 'banana')).toBe('banana banana');
  });

  it('replaces only matching case when caseSensitive', () => {
    expect(
      applySubstitution('Apple apple', { query: 'apple', caseSensitive: true }, 'banana'),
    ).toBe('Apple banana');
  });

  it('matchWhole replaces only when string equals query', () => {
    expect(applySubstitution('apple', { query: 'apple', matchWhole: true }, 'banana')).toBe(
      'banana',
    );
    expect(applySubstitution('apples', { query: 'apple', matchWhole: true }, 'banana')).toBe(
      'apples',
    );
  });

  it('returns input unchanged when query is empty', () => {
    expect(applySubstitution('hello', { query: '' }, 'x')).toBe('hello');
  });

  it('preserves non-matching surrounding text under case-insensitive replace', () => {
    expect(applySubstitution('FooBarFOObar', { query: 'foo' }, 'X')).toBe('XBarXbar');
  });
});

describe('replaceOne / replaceAll', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  const sync = (cells: Array<{ row: number; col: number; raw: string }>): void => {
    store.setState((s) => {
      const map = new Map(s.data.cells);
      for (const c of cells) {
        wb.setText({ sheet: 0, row: c.row, col: c.col }, c.raw);
        map.set(addrKey({ sheet: 0, row: c.row, col: c.col }), {
          value: { kind: 'text', value: c.raw },
          formula: null,
        });
      }
      return { ...s, data: { ...s.data, cells: map } };
    });
    wb.recalc();
  };

  it('replaceOne writes the replacement verbatim through writeInput', () => {
    sync([{ row: 0, col: 0, raw: 'apple' }]);
    replaceOne(wb, { addr: { sheet: 0, row: 0, col: 0 } }, 'banana');
    wb.recalc();
    const v = wb.getValue({ sheet: 0, row: 0, col: 0 });
    expect(v.kind === 'text' && v.value).toBe('banana');
  });

  it('replaceOne refuses to overwrite formula cells', () => {
    wb.setFormula({ sheet: 0, row: 0, col: 0 }, '=1+1');
    wb.recalc();
    replaceOne(wb, { addr: { sheet: 0, row: 0, col: 0 } }, 'overwrite');
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=1+1');
  });

  it('replaceAll counts substituted cells and skips formula cells', () => {
    sync([
      { row: 0, col: 0, raw: 'foo bar' },
      { row: 1, col: 0, raw: 'foofoo' },
      { row: 2, col: 0, raw: 'baz' },
    ]);
    wb.setFormula({ sheet: 0, row: 3, col: 0 }, '=1');
    store.setState((s) => {
      const map = new Map(s.data.cells);
      map.set(addrKey({ sheet: 0, row: 3, col: 0 }), {
        value: { kind: 'number', value: 1 },
        formula: '=1',
      });
      return { ...s, data: { ...s.data, cells: map } };
    });

    const count = replaceAll(store.getState(), wb, { query: 'foo' }, 'X');
    expect(count).toBe(2);
    wb.recalc();
    const a = wb.getValue({ sheet: 0, row: 0, col: 0 });
    const b = wb.getValue({ sheet: 0, row: 1, col: 0 });
    expect(a.kind === 'text' && a.value).toBe('X bar');
    expect(b.kind === 'text' && b.value).toBe('XX');
  });

  it('replaceAll returns 0 for empty query', () => {
    sync([{ row: 0, col: 0, raw: 'hi' }]);
    expect(replaceAll(store.getState(), wb, { query: '' }, 'x')).toBe(0);
  });
});
