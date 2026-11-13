import { afterEach, describe, expect, it, vi } from 'vitest';
import { DATE_PERIODS } from '../../../src/commands/date-periods.js';
import { addrKey } from '../../../src/engine/address.js';
import { en, ja } from '../../../src/i18n/strings.js';
import { attachFilterDropdown } from '../../../src/interact/filter-dropdown.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

// 2026-11-11, 2026-11-12 and 2026-12-01 in the 1900 date system.
const DATES = [46337, 46338, 46357];
const RANGE = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };

const openMac = (dated: boolean, strings = ja): SpreadsheetStore => {
  const store = createSpreadsheetStore();
  store.setState((state) => {
    const cells = new Map(state.data.cells);
    const formats = new Map(state.format.formats);
    cells.set(addrKey({ sheet: 0, row: 0, col: 0 }), {
      value: { kind: 'text', value: 'Date' },
      formula: null,
    });
    DATES.forEach((value, i) => {
      const key = addrKey({ sheet: 0, row: i + 1, col: 0 });
      cells.set(key, { value: { kind: 'number', value }, formula: null });
      if (dated) formats.set(key, { numFmt: { kind: 'date', pattern: 'yyyy-mm-dd' } });
    });
    return {
      ...state,
      data: { ...state.data, cells },
      format: { ...state.format, formats },
    };
  });
  const host = document.createElement('div');
  host.className = 'fc-host';
  host.dataset.fcPlatform = 'mac';
  document.body.appendChild(host);
  attachFilterDropdown({ store, host, strings, locale: strings === ja ? 'ja-JP' : 'en-US' }).open(
    RANGE,
    0,
    { x: 10, y: 20, h: 24 },
  );
  return store;
};

const panel = () => document.querySelector<HTMLElement>('.fc-filter-dropdown') as HTMLElement;
const op = () =>
  panel().querySelector<HTMLSelectElement>(
    '.fc-filter-dropdown__condition-op',
  ) as HTMLSelectElement;
const input = (name: string) =>
  panel().querySelector<HTMLInputElement>(`.fc-filter-dropdown__${name}`) as HTMLInputElement;
const type = (name: string, text: string) => {
  const field = input(name);
  field.value = text;
  field.dispatchEvent(new Event('change'));
};

/** Closes the panel and opens it again on the same column. */
const reopen = (store: SpreadsheetStore, strings = ja): void => {
  document.body.replaceChildren();
  const host = document.createElement('div');
  host.className = 'fc-host';
  host.dataset.fcPlatform = 'mac';
  document.body.appendChild(host);
  attachFilterDropdown({ store, host, strings, locale: 'en-US' }).open(RANGE, 0, {
    x: 10,
    y: 20,
    h: 24,
  });
};

afterEach(() => {
  vi.useRealTimers();
  document.body.replaceChildren();
});

describe('Mac column filter on a date column', () => {
  it('lists the date conditions in the Mac order, ending at All Dates in the Period', () => {
    openMac(true);
    const m = ja.mac.filterPanel;
    const items = [...op().options].map((o) => (o.disabled ? '-' : o.textContent));
    const p = m.periods;
    expect(items).toEqual([
      m.chooseOne,
      '-',
      m.equals,
      '-',
      m.before,
      m.after,
      m.between,
      '-',
      p.tomorrow,
      p.today,
      p.yesterday,
      '-',
      p.nextWeek,
      p.thisWeek,
      p.lastWeek,
      '-',
      p.nextMonth,
      p.thisMonth,
      p.lastMonth,
      '-',
      p.nextQuarter,
      p.thisQuarter,
      p.lastQuarter,
      '-',
      p.nextYear,
      p.thisYear,
      p.lastYear,
      '-',
      p.yearToDate,
      '-',
      m.allDatesInPeriod,
    ]);
    expect(m.allDatesInPeriod).toBe('期間内の全日付...');
  });

  it('takes All Dates in the Period from a quarter and month pop-up and reopens on it', () => {
    const store = openMac(true);
    const m = ja.mac.filterPanel;
    const period = () =>
      panel().querySelector<HTMLSelectElement>(
        '.fc-filter-dropdown__condition-period',
      ) as HTMLSelectElement;
    expect(period().hidden).toBe(true);
    op().value = 'allDatesInPeriod';
    op().dispatchEvent(new Event('change'));
    expect([period().hidden, input('condition-value').hidden]).toEqual([false, true]);
    expect([...period().options].map((o) => o.textContent)).toEqual(
      DATE_PERIODS.slice(DATE_PERIODS.findIndex(([key]) => key === 'quarter1')).map(
        ([key]) => m.periods[key],
      ),
    );
    period().value = 'november';
    period().dispatchEvent(new Event('change'));
    expect(store.getState().ui.filterCriteria[0]?.condition).toEqual({
      op: 'datePeriod',
      value: '',
      period: 'november',
    });
    expect([...store.getState().layout.hiddenRows]).toEqual([3]);
    reopen(store);
    expect([op().value, period().value, period().hidden]).toEqual([
      'allDatesInPeriod',
      'november',
      false,
    ]);
  });

  it('shows a stored operator the menu does not list as an item after it', () => {
    const store = openMac(true, en);
    store.setState((state) => ({
      ...state,
      ui: {
        ...state.ui,
        filterCriteria: [
          {
            range: RANGE,
            byCol: 0,
            hiddenValues: [],
            condition: { op: 'lessThanOrEqual', value: '46338' },
          },
        ],
      },
    }));
    reopen(store, en);
    const last = [...op().options].at(-1);
    expect([op().value, last?.textContent]).toEqual([
      'lessThanOrEqual',
      en.mac.filterPanel.beforeOrEqual,
    ]);
    expect(input('condition-value').value).toBe('2026-11-12');
  });

  it('keeps the numeric list for a column of plain numbers', () => {
    openMac(false);
    const values = [...op().options].map((o) => o.value);
    expect(values).toContain('top10');
    expect(values.some((v) => v.startsWith('datePeriod:'))).toBe(false);
  });

  it('reads a typed date in the locale and stores Before as its serial', () => {
    const store = openMac(true, en);
    op().value = 'lessThan';
    op().dispatchEvent(new Event('change'));
    type('condition-value', '11/12/2026');
    expect(store.getState().ui.filterCriteria[0]?.condition).toEqual({
      op: 'lessThan',
      value: '46338',
    });
    expect([...store.getState().layout.hiddenRows].sort()).toEqual([2, 3]);
  });

  it('applies a relative period with no operand and shows it again on reopening', () => {
    vi.useFakeTimers({ toFake: ['Date'] });
    vi.setSystemTime(new Date(2026, 10, 12, 9, 0));
    const store = openMac(true);
    op().value = 'datePeriod:thisMonth';
    op().dispatchEvent(new Event('change'));
    expect(input('condition-value').hidden).toBe(true);
    expect(store.getState().ui.filterCriteria[0]?.condition).toEqual({
      op: 'datePeriod',
      value: '',
      period: 'thisMonth',
    });
    expect([...store.getState().layout.hiddenRows]).toEqual([3]);
  });

  it('shows a stored date bound as its date on reopening', () => {
    const store = openMac(true, en);
    op().value = 'between';
    op().dispatchEvent(new Event('change'));
    input('condition-value').value = '11/1/2026';
    type('condition-value2', '11/30/2026');
    expect(store.getState().ui.filterCriteria[0]?.condition).toMatchObject({
      op: 'between',
      value: '46327',
      value2: '46356',
    });
    document.body.replaceChildren();
    const host = document.createElement('div');
    host.className = 'fc-host';
    host.dataset.fcPlatform = 'mac';
    document.body.appendChild(host);
    attachFilterDropdown({ store, host, strings: en, locale: 'en-US' }).open(RANGE, 0, {
      x: 10,
      y: 20,
      h: 24,
    });
    expect(op().value).toBe('between');
    expect([input('condition-value').value, input('condition-value2').value]).toEqual([
      '2026-11-01',
      '2026-11-30',
    ]);
  });
});
