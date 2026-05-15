import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { attachConditionalDialog } from '../../../src/interact/conditional-dialog.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const setRange = (
  store: SpreadsheetStore,
  r0: number,
  c0: number,
  r1: number,
  c1: number,
): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      active: { sheet: 0, row: r0, col: c0 },
      anchor: { sheet: 0, row: r0, col: c0 },
      range: { sheet: 0, r0, c0, r1, c1 },
    },
  }));
};

describe('attachConditionalDialog', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;

  beforeEach(() => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    store = createSpreadsheetStore();
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('mounts a hidden overlay and pre-fills selection range on open', () => {
    setRange(store, 0, 0, 4, 2);
    const handle = attachConditionalDialog({ host, store });
    const overlay = document.querySelector<HTMLElement>('.fc-conddlg');
    expect(overlay?.hidden).toBe(true);

    handle.open();
    expect(overlay?.hidden).toBe(false);
    const rangeInput = document.querySelector<HTMLInputElement>(
      '.fc-conddlg__form input[type="text"]',
    );
    expect(rangeInput?.value).toBe('A1:C5');
    handle.detach();
  });

  it('adds a cell-value rule then removes it', () => {
    setRange(store, 0, 0, 1, 0);
    const handle = attachConditionalDialog({ host, store });
    handle.open();

    const valueA = document.querySelector<HTMLInputElement>(
      '.fc-conddlg__sub input[type="number"]',
    ) as HTMLInputElement;
    valueA.value = '50';
    valueA.dispatchEvent(new Event('input', { bubbles: true }));

    const addBtn = document.querySelector<HTMLButtonElement>(
      '.fc-conddlg__addrow .fc-fmtdlg__btn--primary',
    );
    addBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const rules = store.getState().conditional.rules;
    expect(rules).toHaveLength(1);
    expect(rules[0]?.kind).toBe('cell-value');
    if (rules[0]?.kind === 'cell-value') {
      expect(rules[0].a).toBe(50);
      expect(rules[0].op).toBe('>');
    }

    const removeBtn = document.querySelector<HTMLButtonElement>(
      '.fc-conddlg__item .fc-fmtdlg__btn',
    );
    removeBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(store.getState().conditional.rules).toHaveLength(0);

    handle.detach();
  });

  it('clear-all removes every rule', () => {
    mutators.addConditionalRule(store, {
      kind: 'data-bar',
      range: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 },
      color: '#638ec6',
      showValue: true,
    });
    mutators.addConditionalRule(store, {
      kind: 'color-scale',
      range: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 },
      stops: ['#ff0000', '#00ff00'],
    });

    const handle = attachConditionalDialog({ host, store });
    handle.open();

    const buttons = Array.from(document.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__btn'));
    const clearAll = buttons.find((b) => b.textContent === 'すべて削除') as HTMLButtonElement;
    clearAll.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(store.getState().conditional.rules).toHaveLength(0);
    handle.detach();
  });

  it('switches subforms when kind changes', () => {
    const handle = attachConditionalDialog({ host, store });
    handle.open();

    const subs = document.querySelectorAll<HTMLDivElement>('.fc-conddlg__sub');
    expect(subs[0]?.hidden).toBe(false); // cell-value visible by default
    expect(subs[1]?.hidden).toBe(true);
    expect(subs[2]?.hidden).toBe(true);

    const kindSelect = document.querySelector<HTMLSelectElement>(
      '.fc-conddlg__form select',
    ) as HTMLSelectElement;
    kindSelect.value = 'data-bar';
    kindSelect.dispatchEvent(new Event('change', { bubbles: true }));
    expect(subs[0]?.hidden).toBe(true);
    expect(subs[2]?.hidden).toBe(false);

    handle.detach();
  });

  it('labels apply-format color controls and treats Enter as Add Rule', () => {
    setRange(store, 0, 0, 1, 0);
    const handle = attachConditionalDialog({ host, store });
    handle.open();

    const colorInputs = Array.from(
      document.querySelectorAll<HTMLInputElement>('.fc-conddlg__form input[type="color"]'),
    );
    expect(colorInputs.length).toBeGreaterThan(0);
    for (const input of colorInputs) expect(input.getAttribute('aria-label')).toBeTruthy();

    const toggleInputs = Array.from(
      document.querySelectorAll<HTMLInputElement>(
        '.fc-conddlg__form .fc-fmtdlg__row > input[type="checkbox"]',
      ),
    );
    expect(toggleInputs.length).toBeGreaterThan(0);
    for (const input of toggleInputs) expect(input.getAttribute('aria-label')).toBeTruthy();

    document
      .querySelector<HTMLElement>('.fc-conddlg')
      ?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(store.getState().conditional.rules).toHaveLength(1);
    handle.detach();
  });

  it('Escape closes the overlay', () => {
    const handle = attachConditionalDialog({ host, store });
    handle.open();
    const overlay = document.querySelector<HTMLElement>('.fc-conddlg') as HTMLElement;
    overlay.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(overlay.hidden).toBe(true);
    handle.detach();
  });
});
