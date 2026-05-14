import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import {
  applyFilter,
  clearFilter,
  distinctValues,
  setAutoFilter,
} from '../../src/commands/filter.js';
import { sortRange } from '../../src/commands/sort.js';
import { mutators } from '../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../test-utils/index.js';

describe('integration: sort range', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
    const { workbook, instance } = sheet;
    // Header + 4 rows.
    workbook.setText({ sheet: 0, row: 0, col: 0 }, 'item');
    workbook.setText({ sheet: 0, row: 0, col: 1 }, 'qty');
    const rows = [
      ['banana', 7],
      ['apple', 12],
      ['cherry', 3],
      ['date', 25],
    ] as const;
    rows.forEach(([item, qty], i) => {
      workbook.setText({ sheet: 0, row: i + 1, col: 0 }, item);
      workbook.setNumber({ sheet: 0, row: i + 1, col: 1 }, qty);
    });
    workbook.recalc();
    mutators.replaceCells(instance.store, workbook.cells(0));
  });

  afterEach(() => sheet.dispose());

  it('sorts the range ascending by column 0 (text), preserving the header', () => {
    const { instance, workbook } = sheet;
    const ok = sortRange(
      instance.store.getState(),
      instance.store,
      workbook,
      { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 1 },
      { byCol: 0, direction: 'asc', hasHeader: true },
    );
    expect(ok).toBe(true);
    // After sort, row 1 → apple (qty 12), row 4 → date (qty 25)
    expect(workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'text',
      value: 'apple',
    });
    expect(workbook.getValue({ sheet: 0, row: 4, col: 0 })).toEqual({
      kind: 'text',
      value: 'date',
    });
    // Header untouched.
    expect(workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'item',
    });
  });

  it('sorts descending by a numeric column', () => {
    const { instance, workbook } = sheet;
    sortRange(
      instance.store.getState(),
      instance.store,
      workbook,
      { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 1 },
      { byCol: 1, direction: 'desc', hasHeader: true },
    );
    // qty desc: 25, 12, 7, 3 → row 1 = date(25)
    expect(workbook.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({
      kind: 'number',
      value: 25,
    });
    expect(workbook.getValue({ sheet: 0, row: 4, col: 1 })).toEqual({
      kind: 'number',
      value: 3,
    });
  });

  it('refuses to sort when the by-column lies outside the range', () => {
    const { instance, workbook } = sheet;
    const ok = sortRange(
      instance.store.getState(),
      instance.store,
      workbook,
      { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 1 },
      { byCol: 5, direction: 'asc' },
    );
    expect(ok).toBe(false);
  });
});

describe('integration: filter range', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
    const { workbook, instance } = sheet;
    workbook.setText({ sheet: 0, row: 0, col: 0 }, 'fruit');
    workbook.setText({ sheet: 0, row: 1, col: 0 }, 'apple');
    workbook.setText({ sheet: 0, row: 2, col: 0 }, 'banana');
    workbook.setText({ sheet: 0, row: 3, col: 0 }, 'apple');
    workbook.setText({ sheet: 0, row: 4, col: 0 }, 'cherry');
    workbook.recalc();
    mutators.replaceCells(instance.store, workbook.cells(0));
  });

  afterEach(() => sheet.dispose());

  it('applyFilter hides rows where the predicate returns false', () => {
    const { instance } = sheet;
    const range = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 };
    const hidden = applyFilter(instance.store.getState(), instance.store, range, 0, (cell) => {
      const v = cell?.value as { kind: string; value?: string } | undefined;
      return v?.kind === 'text' && v.value === 'apple';
    });
    expect(hidden).toBe(2); // banana + cherry hidden
    const hr = instance.store.getState().layout.hiddenRows;
    expect(hr.has(2)).toBe(true); // banana
    expect(hr.has(4)).toBe(true); // cherry
    expect(hr.has(1)).toBe(false); // apple visible
  });

  it('applyFilter stamps ui.filterRange so headers paint the chevron', () => {
    const { instance } = sheet;
    const range = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 };
    applyFilter(instance.store.getState(), instance.store, range, 0, () => true);
    expect(instance.store.getState().ui.filterRange).toEqual(range);
  });

  it('clearFilter without a range reveals every hidden row', () => {
    const { instance } = sheet;
    const range = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 };
    applyFilter(instance.store.getState(), instance.store, range, 0, () => false);
    expect(instance.store.getState().layout.hiddenRows.size).toBe(4);
    clearFilter(instance.store.getState(), instance.store);
    expect(instance.store.getState().layout.hiddenRows.size).toBe(0);
    expect(instance.store.getState().ui.filterRange).toBeNull();
  });

  it('setAutoFilter writes ui.filterRange without changing hiddenRows', () => {
    const { instance } = sheet;
    setAutoFilter(instance.store, { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 });
    expect(instance.store.getState().ui.filterRange).not.toBeNull();
    expect(instance.store.getState().layout.hiddenRows.size).toBe(0);
  });

  it('distinctValues returns unique cell values from the by-column', () => {
    const { instance } = sheet;
    const range = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 };
    expect(distinctValues(instance.store.getState(), range, 0).sort()).toEqual([
      'apple',
      'banana',
      'cherry',
    ]);
  });
});
