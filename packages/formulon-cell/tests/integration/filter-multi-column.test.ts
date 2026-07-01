import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import {
  applyConditionFilter,
  applyValueFilter,
  clearFilter,
  reapplyFilters,
} from '../../src/commands/filter.js';
import { mutators } from '../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../test-utils/index.js';

/** Region / Amount table: 4 data rows under a header. */
const seed = (sheet: MountedStubSheet): void => {
  const { workbook, instance } = sheet;
  workbook.setText({ sheet: 0, row: 0, col: 0 }, 'Region');
  workbook.setText({ sheet: 0, row: 0, col: 1 }, 'Amount');
  const rows = [
    ['West', 50],
    ['East', 200],
    ['West', 150],
    ['East', 80],
  ] as const;
  rows.forEach(([region, amount], i) => {
    workbook.setText({ sheet: 0, row: i + 1, col: 0 }, region);
    workbook.setNumber({ sheet: 0, row: i + 1, col: 1 }, amount);
  });
  workbook.recalc();
  mutators.replaceCells(instance.store, workbook.cells(0));
};

const RANGE = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 1 } as const;

describe('integration: multi-column AutoFilter ANDs across columns (C-3)', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
    seed(sheet);
  });

  afterEach(() => sheet.dispose());

  const hidden = (): number[] =>
    Array.from(sheet.instance.store.getState().layout.hiddenRows).sort((a, b) => a - b);

  it('keeps the first column filter when a second column filter is applied', () => {
    const { instance } = sheet;
    // Column 0: show only West → hide the two East rows (2, 4).
    applyValueFilter(instance.store.getState(), instance.store, RANGE, 0, ['East']);
    expect(hidden()).toEqual([2, 4]);

    // Column 1: Amount > 100. Excel ANDs — the West restriction must survive.
    applyConditionFilter(instance.store.getState(), instance.store, RANGE, 1, {
      op: 'greaterThan',
      value: '100',
    });
    // Only row 3 (West, 150) passes BOTH criteria. Rows 1 (West,50), 2 & 4 (East) hidden.
    expect(hidden()).toEqual([1, 2, 4]);
  });

  it('re-ANDs after a value filter follows a condition filter', () => {
    const { instance } = sheet;
    applyConditionFilter(instance.store.getState(), instance.store, RANGE, 1, {
      op: 'greaterThan',
      value: '100',
    });
    // Amount>100 hides rows 1 (50) and 4 (80).
    expect(hidden()).toEqual([1, 4]);
    applyValueFilter(instance.store.getState(), instance.store, RANGE, 0, ['East']);
    // AND with West-only: only row 3 survives.
    expect(hidden()).toEqual([1, 2, 4]);
  });
});

describe('integration: condition filters persist for Reapply (H-15)', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
    seed(sheet);
  });

  afterEach(() => sheet.dispose());

  it('re-applies a persisted condition filter after data changes', () => {
    const { instance, workbook } = sheet;
    applyConditionFilter(instance.store.getState(), instance.store, RANGE, 1, {
      op: 'greaterThan',
      value: '100',
    });
    expect(Array.from(instance.store.getState().layout.hiddenRows).sort()).toEqual([1, 4]);

    // The criterion is persisted (not lost like the old code).
    expect(instance.store.getState().ui.filterCriteria).toHaveLength(1);
    expect(instance.store.getState().ui.filterCriteria[0]?.condition).toEqual({
      op: 'greaterThan',
      value: '100',
    });

    // Change row 1 Amount to 300 (now passes), then Reapply.
    workbook.setNumber({ sheet: 0, row: 1, col: 1 }, 300);
    workbook.recalc();
    mutators.replaceCells(instance.store, workbook.cells(0));
    reapplyFilters(instance.store.getState(), instance.store);
    // Now only row 4 (80) stays hidden.
    expect(Array.from(instance.store.getState().layout.hiddenRows).sort()).toEqual([4]);
  });

  it('clearFilter drops both value and condition criteria', () => {
    const { instance } = sheet;
    applyConditionFilter(instance.store.getState(), instance.store, RANGE, 1, {
      op: 'greaterThan',
      value: '100',
    });
    clearFilter(instance.store.getState(), instance.store, { ...RANGE });
    expect(instance.store.getState().layout.hiddenRows.size).toBe(0);
    expect(instance.store.getState().ui.filterCriteria).toHaveLength(0);
  });
});
