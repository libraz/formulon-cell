import { afterEach, describe, expect, it } from 'vitest';
import { addrKey } from '../../../src/engine/address.js';
import { en } from '../../../src/i18n/strings.js';
import { attachFilterDropdown } from '../../../src/interact/filter-dropdown.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const RANGE = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 };

const open = (store: SpreadsheetStore): void => {
  const host = document.createElement('div');
  host.className = 'fc-host';
  host.dataset.fcPlatform = 'mac';
  document.body.appendChild(host);
  attachFilterDropdown({ store, host, strings: en, locale: 'en-US' }).open(RANGE, 0, {
    x: 10,
    y: 20,
    h: 24,
  });
};

const seeded = (): SpreadsheetStore => {
  const store = createSpreadsheetStore();
  store.setState((state) => {
    const cells = new Map(state.data.cells);
    cells.set(addrKey({ sheet: 0, row: 0, col: 0 }), {
      value: { kind: 'text', value: 'N' },
      formula: null,
    });
    [1, 4, 6, 9].forEach((value, i) => {
      cells.set(addrKey({ sheet: 0, row: i + 1, col: 0 }), {
        value: { kind: 'number', value },
        formula: null,
      });
    });
    return { ...state, data: { ...state.data, cells } };
  });
  return store;
};

const q = <T extends HTMLElement>(selector: string): T =>
  document.querySelector<T>(`.fc-filter-dropdown ${selector}`) as T;
const pick = (selector: string, value: string): void => {
  const el = q<HTMLSelectElement | HTMLInputElement>(selector);
  el.value = value;
  el.dispatchEvent(new Event('change'));
};

afterEach(() => {
  document.body.replaceChildren();
});

describe('Mac column filter with two conditions', () => {
  it('offers the And/Or row and a second comparison only after a one-operand comparison', () => {
    open(seeded());
    const join = q<HTMLElement>('.fc-filter-dropdown__condition-join');
    const second = q<HTMLSelectElement>('.fc-filter-dropdown__condition-op2');
    expect([join.hidden, second.hidden]).toEqual([true, true]);
    pick('.fc-filter-dropdown__condition-op', 'lessThan');
    expect([join.hidden, second.hidden]).toEqual([false, false]);
    expect([...join.querySelectorAll('label')].map((l) => l.textContent)).toEqual([
      en.mac.filterPanel.and,
      en.mac.filterPanel.or,
    ]);
    expect([...second.options].map((o) => o.value)).not.toContain('top10');
    expect(q<HTMLInputElement>('.fc-filter-dropdown__condition-second').hidden).toBe(true);
    pick('.fc-filter-dropdown__condition-op', 'top10');
    expect([join.hidden, second.hidden]).toEqual([true, true]);
  });

  it('applies an Or pair and shows it again on reopening', () => {
    const store = seeded();
    open(store);
    pick('.fc-filter-dropdown__condition-op', 'lessThan');
    q<HTMLInputElement>('.fc-filter-dropdown__condition-value').value = '2';
    const or = q<HTMLInputElement>('.fc-filter-dropdown__condition-join input[value="or"]');
    or.checked = true;
    pick('.fc-filter-dropdown__condition-op2', 'greaterThan');
    pick('.fc-filter-dropdown__condition-second', '8');
    expect(store.getState().ui.filterCriteria[0]?.condition).toEqual({
      op: 'lessThan',
      value: '2',
      second: { op: 'greaterThan', value: '8' },
      join: 'or',
    });
    expect([...store.getState().layout.hiddenRows].sort()).toEqual([2, 3]);

    document.body.replaceChildren();
    open(store);
    expect(q<HTMLSelectElement>('.fc-filter-dropdown__condition-op2').value).toBe('greaterThan');
    expect(q<HTMLInputElement>('.fc-filter-dropdown__condition-second').value).toBe('8');
    expect(
      q<HTMLInputElement>('.fc-filter-dropdown__condition-join input[value="or"]').checked,
    ).toBe(true);
  });
});
