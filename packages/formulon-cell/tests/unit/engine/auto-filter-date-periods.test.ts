import { DynamicFilterType, FilterKind } from '@libraz/formulon';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { DATE_PERIODS } from '../../../src/commands/date-periods.js';
import { applyConditionFilter } from '../../../src/commands/filter.js';
import { recordFilterChange } from '../../../src/commands/filter-history.js';
import { History } from '../../../src/commands/history.js';
import { addrKey } from '../../../src/engine/address.js';
import {
  autoFilterXmlFromState,
  hydrateAutoFilterFromEngine,
  syncAutoFilterToEngine,
} from '../../../src/engine/auto-filter-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };

// 2026-11-11, 2026-11-12 and 2026-12-01 in the 1900 date system.
const NOV_11 = 46337;
const NOV_12 = 46338;
const DEC_1 = 46357;

const seeded = async () => {
  const wb = await WorkbookHandle.createDefault();
  const store = createSpreadsheetStore();
  const values: (string | number)[] = ['Date', NOV_11, NOV_12, DEC_1];
  store.setState((state) => {
    const cells = new Map(state.data.cells);
    values.forEach((value, row) => {
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

afterEach(() => {
  vi.useRealTimers();
});

describe('AutoFilter relative date periods', () => {
  it('stores every period as its dynamic filter type and reads it back after a save', async () => {
    expect(DATE_PERIODS).toHaveLength(32);
    expect(new Set(DATE_PERIODS.map(([, type]) => type)).size).toBe(32);
    for (const [period, type, xmlType] of DATE_PERIODS) {
      const { wb, store } = await seeded();
      try {
        applyConditionFilter(store.getState(), store, range, 0, {
          op: 'datePeriod',
          value: '',
          period,
        });
        expect(syncAutoFilterToEngine(wb, store.getState(), 0, store), period).toBe(true);
        expect(wb.getAutoFilter(0)?.columns[0], period).toMatchObject({
          kind: FilterKind.Dynamic,
          dynamicType: type,
        });
        expect(autoFilterXmlFromState(store.getState(), range)).toContain(
          `<dynamicFilter type="${xmlType}"/>`,
        );
        const reopened = await WorkbookHandle.loadBytes(wb.save());
        try {
          const reloaded = createSpreadsheetStore();
          hydrateAutoFilterFromEngine(reopened, reloaded, 0);
          expect(reloaded.getState().ui.filterCriteria[0]?.condition, period).toEqual({
            op: 'datePeriod',
            value: '',
            period,
          });
        } finally {
          reopened.dispose();
        }
      } finally {
        wb.dispose();
      }
    }
  });

  it('keeps the rows in the period, judged against today', async () => {
    vi.useFakeTimers({ toFake: ['Date'] });
    vi.setSystemTime(new Date(2026, 10, 12, 9, 0));
    const { wb, store } = await seeded();
    try {
      const shown = (period: (typeof DATE_PERIODS)[number][0]): number[] => {
        applyConditionFilter(store.getState(), store, range, 0, {
          op: 'datePeriod',
          value: '',
          period,
        });
        const hidden = store.getState().layout.hiddenRows;
        return [1, 2, 3].filter((row) => !hidden.has(row));
      };
      expect(shown('today')).toEqual([2]);
      expect(shown('yesterday')).toEqual([1]);
      expect(shown('thisMonth')).toEqual([1, 2]);
      expect(shown('nextMonth')).toEqual([3]);
      expect(shown('december')).toEqual([3]);
      expect(shown('quarter4')).toEqual([1, 2, 3]);
      expect(shown('yearToDate')).toEqual([1, 2]);
    } finally {
      wb.dispose();
    }
  });

  it('reads a non-date dynamic type as no period', async () => {
    const { wb, store } = await seeded();
    try {
      applyConditionFilter(store.getState(), store, range, 0, { op: 'aboveAverage', value: '' });
      syncAutoFilterToEngine(wb, store.getState(), 0, store);
      expect(wb.getAutoFilter(0)?.columns[0]?.dynamicType).toBe(DynamicFilterType.AboveAverage);
      const reloaded = createSpreadsheetStore();
      hydrateAutoFilterFromEngine(wb, reloaded, 0);
      expect(reloaded.getState().ui.filterCriteria[0]?.condition?.op).toBe('aboveAverage');
    } finally {
      wb.dispose();
    }
  });
});

describe('AutoFilter period changes in the undo history', () => {
  it('records a change from one period to another as its own step', async () => {
    vi.useFakeTimers({ toFake: ['Date'] });
    vi.setSystemTime(new Date(2026, 10, 12, 9, 0));
    const { wb, store } = await seeded();
    try {
      const history = new History();
      const apply = (period: 'lastYear' | 'nextYear') =>
        recordFilterChange(history, store, () =>
          applyConditionFilter(store.getState(), store, range, 0, {
            op: 'datePeriod',
            value: '',
            period,
          }),
        );
      // Both periods hide every row here, so only the condition tells the two steps apart.
      apply('lastYear');
      apply('nextYear');
      expect(store.getState().ui.filterCriteria[0]?.condition?.period).toBe('nextYear');
      history.undo();
      expect(store.getState().ui.filterCriteria[0]?.condition?.period).toBe('lastYear');
    } finally {
      wb.dispose();
    }
  });
});
