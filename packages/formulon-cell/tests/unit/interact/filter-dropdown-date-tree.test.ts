import { afterEach, describe, expect, it } from 'vitest';
import { addrKey } from '../../../src/engine/address.js';
import { en } from '../../../src/i18n/strings.js';
import { attachFilterDropdown } from '../../../src/interact/filter-dropdown.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

// 2025-12-31, 2026-05-08, 2026-05-18, 2026-06-07 in the 1900 date system.
const DATES = [46022, 46150, 46160, 46180];
const RANGE = { sheet: 0, r0: 0, c0: 0, r1: 5, c1: 0 };

const open = (): SpreadsheetStore => {
  const store = createSpreadsheetStore();
  store.setState((state) => {
    const cells = new Map(state.data.cells);
    const formats = new Map(state.format.formats);
    const put = (
      row: number,
      value: { kind: 'text'; value: string } | { kind: 'number'; value: number },
    ) => cells.set(addrKey({ sheet: 0, row, col: 0 }), { value, formula: null });
    put(0, { kind: 'text', value: 'Date' });
    DATES.forEach((value, i) => {
      put(i + 1, { kind: 'number', value });
      formats.set(addrKey({ sheet: 0, row: i + 1, col: 0 }), {
        numFmt: { kind: 'date', pattern: 'yyyy-mm-dd' },
      });
    });
    put(5, { kind: 'text', value: 'n/a' });
    return { ...state, data: { ...state.data, cells }, format: { ...state.format, formats } };
  });
  const host = document.createElement('div');
  document.body.appendChild(host);
  attachFilterDropdown({ store, host, strings: en, locale: 'en-US' }).open(RANGE, 0, {
    x: 10,
    y: 20,
    h: 24,
  });
  return store;
};

const rows = (): string[] =>
  [...document.querySelectorAll<HTMLElement>('.fc-filter-dropdown__list .fc-filter-dropdown__row')]
    .slice(1)
    .map((row) => `${row.dataset.fcLevel ?? '0'}:${row.textContent}`);
const box = (value: string): HTMLInputElement =>
  [...document.querySelectorAll<HTMLInputElement>('.fc-filter-dropdown__list input')].find(
    (input) => input.value === value,
  ) as HTMLInputElement;
const expand = (path: string): void =>
  document
    .querySelector<HTMLButtonElement>(
      `.fc-filter-dropdown__disclosure[data-fc-date-path="${path}"]`,
    )
    ?.click();
const toggle = (value: string): void => {
  const input = box(value);
  input.checked = !input.checked;
  input.dispatchEvent(new Event('change'));
};

afterEach(() => {
  document.body.replaceChildren();
});

describe('column filter checklist on a date column', () => {
  it('groups the dates by year, then month, then day, collapsed until opened', () => {
    open();
    expect(rows()).toEqual(['0:2025', '0:2026', '0:n/a']);
    expand('2026');
    expect(rows()).toEqual(['0:2025', '0:2026', '1:May', '1:June', '0:n/a']);
    expand('2026-5');
    expect(rows()).toEqual(['0:2025', '0:2026', '1:May', '2:8', '2:18', '1:June', '0:n/a']);
  });

  it('hides a whole month from its checkbox and shows the year as mixed', () => {
    const store = open();
    expand('2026');
    toggle('group:2026-5');
    expect([box('group:2026').checked, box('group:2026').indeterminate]).toEqual([false, true]);
    document.querySelector<HTMLButtonElement>('.fc-filter-dropdown__apply')?.click();
    expect(store.getState().ui.filterCriteria[0]?.hiddenValues.sort()).toEqual(['46150', '46160']);
    expect([...store.getState().layout.hiddenRows].sort()).toEqual([2, 3]);
  });

  it('marks the month mixed when one of its days is cleared', () => {
    open();
    expand('2026');
    expand('2026-5');
    toggle('46160');
    expect([box('group:2026-5').checked, box('group:2026-5').indeterminate]).toEqual([false, true]);
  });

  it('lists the dates flat while searching', () => {
    open();
    const search = document.querySelector<HTMLInputElement>('.fc-filter-dropdown__search');
    if (!search) throw new Error('missing search');
    search.value = '2026-05';
    search.dispatchEvent(new Event('input'));
    expect(rows()).toEqual(['0:2026-05-08', '0:2026-05-18']);
  });
});
