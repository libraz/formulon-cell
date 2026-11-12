import { ComparisonOperator, DynamicFilterType, FilterKind } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { applyConditionFilter, type ConditionFilterOptions } from '../../../src/commands/filter.js';
import { addrKey } from '../../../src/engine/address.js';
import {
  autoFilterXmlFromState,
  hydrateAutoFilterFromEngine,
  syncAutoFilterToEngine,
} from '../../../src/engine/auto-filter-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const range = { sheet: 0, r0: 0, c0: 0, r1: 5, c1: 0 };
const VALUES: (string | number)[] = ['N', 1, 4, 6, 9, 'apple'];

const seeded = async () => {
  const wb = await WorkbookHandle.createDefault();
  const store = createSpreadsheetStore();
  store.setState((state) => {
    const cells = new Map(state.data.cells);
    VALUES.forEach((value, row) => {
      const addr = { sheet: 0, row, col: 0 };
      if (typeof value === 'number') wb.setNumber(addr, value);
      else wb.setText(addr, value);
      cells.set(addrKey(addr), {
        value: typeof value === 'number' ? { kind: 'number', value } : { kind: 'text', value },
        formula: null,
      });
    });
    return { ...state, data: { ...state.data, cells } };
  });
  return { wb, store };
};

/** Rows 1-5 the filter keeps, in the JS store and in the engine's own evaluation. */
const kept = (
  wb: WorkbookHandle,
  store: ReturnType<typeof createSpreadsheetStore>,
): { js: number[]; engine: number[] } => {
  const hidden = store.getState().layout.hiddenRows;
  const match = wb.evaluateAutoFilter(0)?.match ?? [];
  return {
    js: [1, 2, 3, 4, 5].filter((row) => !hidden.has(row)),
    engine: [1, 2, 3, 4, 5].filter((row) => match[row - 1] === true),
  };
};

const roundTrip = async (wb: WorkbookHandle) => {
  const reopened = await WorkbookHandle.loadBytes(wb.save());
  const reloaded = createSpreadsheetStore();
  hydrateAutoFilterFromEngine(reopened, reloaded, 0);
  const condition = reloaded.getState().ui.filterCriteria[0]?.condition;
  reopened.dispose();
  return condition;
};

describe('AutoFilter two-condition custom filters', () => {
  const cases: {
    condition: ConditionFilterOptions;
    native: object;
    xml: string;
    rows: number[];
  }[] = [
    {
      condition: {
        op: 'lessThan',
        value: '2',
        second: { op: 'greaterThan', value: '8' },
        join: 'or',
      },
      native: {
        kind: FilterKind.Custom,
        customAnd: false,
        conditions: [
          { operator: ComparisonOperator.LessThan, value: '2' },
          { operator: ComparisonOperator.GreaterThan, value: '8' },
        ],
      },
      xml: '<customFilters><customFilter operator="lessThan" val="2"/><customFilter operator="greaterThan" val="8"/></customFilters>',
      rows: [1, 4],
    },
    {
      condition: {
        op: 'greaterThan',
        value: '2',
        second: { op: 'notEquals', value: '6' },
        join: 'and',
      },
      native: {
        kind: FilterKind.Custom,
        customAnd: true,
        conditions: [
          { operator: ComparisonOperator.GreaterThan, value: '2' },
          { operator: ComparisonOperator.NotEqual, value: '6' },
        ],
      },
      xml: '<customFilters and="1"><customFilter operator="greaterThan" val="2"/><customFilter operator="notEqual" val="6"/></customFilters>',
      rows: [2, 4],
    },
    {
      condition: {
        op: 'beginsWith',
        value: 'ap',
        second: { op: 'equals', value: '1' },
        join: 'or',
      },
      native: {
        kind: FilterKind.Custom,
        customAnd: false,
        conditions: [
          { operator: ComparisonOperator.Equal, value: 'ap*' },
          { operator: ComparisonOperator.Equal, value: '1' },
        ],
      },
      xml: '<customFilters><customFilter operator="equal" val="ap*"/><customFilter operator="equal" val="1"/></customFilters>',
      rows: [1, 5],
    },
  ];

  it('writes, evaluates and reads back an And or Or pair', async () => {
    for (const { condition, native, xml, rows } of cases) {
      const { wb, store } = await seeded();
      try {
        applyConditionFilter(store.getState(), store, range, 0, condition);
        expect(syncAutoFilterToEngine(wb, store.getState(), 0, store)).toBe(true);
        expect(wb.getAutoFilter(0)?.columns[0]).toMatchObject(native);
        expect(autoFilterXmlFromState(store.getState(), range)).toContain(xml);
        expect(kept(wb, store)).toEqual({ js: rows, engine: rows });
        expect(await roundTrip(wb)).toEqual(condition);
      } finally {
        wb.dispose();
      }
    }
  });

  it('keeps the greater-or-equal and less-or-equal And pair as Between', async () => {
    const { wb, store } = await seeded();
    try {
      applyConditionFilter(store.getState(), store, range, 0, {
        op: 'greaterThanOrEqual',
        value: '4',
        second: { op: 'lessThanOrEqual', value: '6' },
        join: 'and',
      });
      syncAutoFilterToEngine(wb, store.getState(), 0, store);
      expect(await roundTrip(wb)).toEqual({ op: 'between', value: '4', value2: '6' });
    } finally {
      wb.dispose();
    }
  });

  it('writes Below Average as a dynamic filter and reads it back after a save', async () => {
    const { wb, store } = await seeded();
    try {
      applyConditionFilter(store.getState(), store, range, 0, { op: 'belowAverage', value: '' });
      syncAutoFilterToEngine(wb, store.getState(), 0, store);
      expect(wb.getAutoFilter(0)?.columns[0]).toMatchObject({
        kind: FilterKind.Dynamic,
        dynamicType: DynamicFilterType.BelowAverage,
      });
      expect(kept(wb, store).js).toEqual([1, 2]);
      expect(await roundTrip(wb)).toEqual({ op: 'belowAverage', value: '' });
    } finally {
      wb.dispose();
    }
  });
});
