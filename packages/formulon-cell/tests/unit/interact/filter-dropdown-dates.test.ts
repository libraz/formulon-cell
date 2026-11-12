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

afterEach(() => {
  vi.useRealTimers();
  document.body.replaceChildren();
});

describe('Mac column filter on a date column', () => {
  // The Mac pop-up's order for a date column is provisional: it follows the desktop Date Filters
  // until measured on the Mac app.
  it('lists the date comparisons, then the relative periods, in the Mac wording', () => {
    openMac(true);
    const m = ja.mac.filterPanel;
    const enabled = [...op().options].filter((o) => !o.disabled).map((o) => o.textContent);
    expect(enabled).toEqual([
      m.chooseOne,
      m.equals,
      m.notEquals,
      m.before,
      m.beforeOrEqual,
      m.after,
      m.afterOrEqual,
      m.between,
      ...DATE_PERIODS.map(([key]) => m.periods[key]),
    ]);
    expect([...op().options].filter((o) => o.disabled)).toHaveLength(8);
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
