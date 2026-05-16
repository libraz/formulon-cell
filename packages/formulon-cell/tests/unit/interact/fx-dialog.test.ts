import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { ja } from '../../../src/i18n/strings.js';
import { attachFxDialog, FUNCTION_DESCRIPTIONS } from '../../../src/interact/fx-dialog.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('attachFxDialog', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('mounts a hidden overlay until open() is called', () => {
    const handle = attachFxDialog({
      host,
      store: createSpreadsheetStore(),
      onInsert: () => {},
    });
    const overlay = document.querySelector<HTMLElement>('.fc-fxdialog');
    expect(overlay?.hidden).toBe(true);
    handle.open();
    expect(overlay?.hidden).toBe(false);
    handle.detach();
  });

  it('renders the function picker step on open without a seed', () => {
    const handle = attachFxDialog({
      host,
      store: createSpreadsheetStore(),
      onInsert: () => {},
    });
    handle.open();
    const picker = document.querySelector<HTMLElement>('.fc-fxdialog__picker');
    const args = document.querySelector<HTMLElement>('.fc-fxdialog__args');
    expect(picker?.hidden).toBe(false);
    expect(args?.hidden).toBe(true);
    expect(document.querySelectorAll('.fc-fxdialog__item').length).toBeGreaterThan(0);
    handle.detach();
  });

  it('jumps straight to argument entry when open() is given a known function', () => {
    const handle = attachFxDialog({
      host,
      store: createSpreadsheetStore(),
      onInsert: () => {},
    });
    handle.open('SUM');
    const picker = document.querySelector<HTMLElement>('.fc-fxdialog__picker');
    const args = document.querySelector<HTMLElement>('.fc-fxdialog__args');
    expect(picker?.hidden).toBe(true);
    expect(args?.hidden).toBe(false);
    const argName = document.querySelector<HTMLElement>('.fc-fxdialog__args-name');
    expect(argName?.textContent).toMatch(/^SUM\(/);
    handle.detach();
  });

  it('prefills seeded function arguments from the spreadsheet context', () => {
    const inserted: string[] = [];
    const handle = attachFxDialog({
      host,
      store: createSpreadsheetStore(),
      getInitialArguments: (name) => (name === 'SUM' ? ['A1:A5'] : null),
      onInsert: (formula) => inserted.push(formula),
    });
    handle.open('SUM');

    const input = document.querySelector<HTMLInputElement>('.fc-fxdialog__arg-input');
    expect(input?.value).toBe('A1:A5');
    expect(document.querySelector<HTMLElement>('.fc-fxdialog__preview')?.textContent).toBe(
      '=SUM(A1:A5)',
    );

    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    expect(inserted).toEqual(['=SUM(A1:A5)']);
    handle.detach();
  });

  it('filters the picker list by case-insensitive prefix', () => {
    const handle = attachFxDialog({
      host,
      store: createSpreadsheetStore(),
      onInsert: () => {},
    });
    handle.open();
    const search = document.querySelector<HTMLInputElement>('.fc-fxdialog__search');
    expect(search).toBeTruthy();
    if (!search) return;
    search.value = 'vlo';
    search.dispatchEvent(new Event('input'));
    const items = document.querySelectorAll<HTMLElement>('.fc-fxdialog__item-name');
    const names = Array.from(items).map((i) => i.textContent ?? '');
    expect(names.every((n) => n.includes('VLO'))).toBe(true);
    expect(names).toContain('VLOOKUP');
    handle.detach();
  });

  it('filters the function picker by category', () => {
    const handle = attachFxDialog({
      host,
      store: createSpreadsheetStore(),
      onInsert: () => {},
    });
    handle.open();
    const category = document.querySelector<HTMLSelectElement>('.fc-fxdialog__category');
    expect(category).toBeTruthy();
    if (!category) return;

    category.value = 'text';
    category.dispatchEvent(new Event('change'));

    const names = Array.from(document.querySelectorAll<HTMLElement>('.fc-fxdialog__item-name')).map(
      (item) => item.textContent ?? '',
    );
    expect(names).toContain('CONCAT');
    expect(names).toContain('TEXT');
    expect(names).not.toContain('SUM');
    handle.detach();
  });

  it('lists recently used functions alphabetically', () => {
    const handle = attachFxDialog({
      host,
      store: createSpreadsheetStore(),
      onInsert: () => {},
    });

    handle.open('VLOOKUP');
    handle.close();
    handle.open('IF');
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn')?.click();

    const category = document.querySelector<HTMLSelectElement>('.fc-fxdialog__category');
    expect(category).toBeTruthy();
    if (!category) return;
    category.value = 'recent';
    category.dispatchEvent(new Event('change'));

    const names = Array.from(document.querySelectorAll<HTMLElement>('.fc-fxdialog__item-name')).map(
      (item) => item.textContent ?? '',
    );
    expect(names).toEqual(['IF', 'VLOOKUP']);
    handle.detach();
  });

  it('wires the function search box to the active listbox option', () => {
    const handle = attachFxDialog({
      host,
      store: createSpreadsheetStore(),
      onInsert: () => {},
    });
    handle.open();
    const search = document.querySelector<HTMLInputElement>('.fc-fxdialog__search');
    const list = document.querySelector<HTMLElement>('.fc-fxdialog__list');
    if (!search || !list) throw new Error('expected function picker controls');
    expect(search.getAttribute('role')).toBe('combobox');
    expect(search.getAttribute('aria-controls')).toBe(list.id);
    expect(search.getAttribute('aria-label')).toBeTruthy();
    expect(list.getAttribute('role')).toBe('listbox');
    expect(list.getAttribute('aria-label')).toBeTruthy();

    const firstActive = search.getAttribute('aria-activedescendant');
    expect(firstActive).toBeTruthy();
    expect(document.getElementById(firstActive ?? '')?.getAttribute('aria-selected')).toBe('true');
    const firstName = document.querySelector<HTMLElement>('.fc-fxdialog__summary-name')?.textContent;
    expect(firstName).toContain('(');

    search.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    const secondActive = search.getAttribute('aria-activedescendant');
    expect(secondActive).toBeTruthy();
    expect(secondActive).not.toBe(firstActive);
    expect(document.querySelector<HTMLElement>('.fc-fxdialog__summary-name')?.textContent).not.toBe(
      firstName,
    );

    search.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', bubbles: true }));
    expect(search.getAttribute('aria-activedescendant')).toBe(firstActive);

    search.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true }));
    const lastActive = search.getAttribute('aria-activedescendant');
    expect(lastActive).toBeTruthy();
    expect(lastActive).not.toBe(firstActive);
    handle.detach();
  });

  it('assembles the formula and fires onInsert on confirm', () => {
    const inserted: string[] = [];
    const handle = attachFxDialog({
      host,
      store: createSpreadsheetStore(),
      onInsert: (f) => inserted.push(f),
    });
    handle.open('SUM');
    const inputs = document.querySelectorAll<HTMLInputElement>('.fc-fxdialog__arg-input');
    expect(inputs.length).toBeGreaterThan(0);
    const input = inputs[0];
    expect(input).toBeDefined();
    if (!input) throw new Error('expected function argument input');
    input.value = 'A1:A5';
    input.dispatchEvent(new Event('input'));
    const insertBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    insertBtn?.click();
    expect(inserted).toEqual(['=SUM(A1:A5)']);
    const overlay = document.querySelector<HTMLElement>('.fc-fxdialog');
    expect(overlay?.hidden).toBe(true);
    handle.detach();
  });

  it('assembles multi-argument and zero-argument formulas', () => {
    const inserted: string[] = [];
    const handle = attachFxDialog({
      host,
      store: createSpreadsheetStore(),
      onInsert: (f) => inserted.push(f),
    });

    handle.open('IF');
    const ifInputs = document.querySelectorAll<HTMLInputElement>('.fc-fxdialog__arg-input');
    expect(ifInputs).toHaveLength(3);
    const [logicalTest, valueIfTrue, valueIfFalse] = Array.from(ifInputs);
    if (!logicalTest || !valueIfTrue || !valueIfFalse) throw new Error('expected IF argument inputs');
    logicalTest.value = 'A1>5';
    logicalTest.dispatchEvent(new Event('input'));
    valueIfTrue.value = '"yes"';
    valueIfTrue.dispatchEvent(new Event('input'));
    valueIfFalse.value = '"no"';
    valueIfFalse.dispatchEvent(new Event('input'));
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    expect(inserted.at(-1)).toBe('=IF(A1>5, "yes", "no")');

    handle.open('ROUND');
    const roundInputs = document.querySelectorAll<HTMLInputElement>('.fc-fxdialog__arg-input');
    expect(roundInputs).toHaveLength(2);
    const [numberInput, digitsInput] = Array.from(roundInputs);
    if (!numberInput || !digitsInput) throw new Error('expected ROUND argument inputs');
    numberInput.value = '1.234';
    numberInput.dispatchEvent(new Event('input'));
    digitsInput.value = '2';
    digitsInput.dispatchEvent(new Event('input'));
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    expect(inserted.at(-1)).toBe('=ROUND(1.234, 2)');

    handle.open('TODAY');
    expect(document.querySelectorAll<HTMLInputElement>('.fc-fxdialog__arg-input')).toHaveLength(0);
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    expect(inserted.at(-1)).toBe('=TODAY()');

    handle.detach();
  });

  it('exposes spreadsheet-style descriptions for the common functions', () => {
    expect(FUNCTION_DESCRIPTIONS.SUM?.en).toMatch(/sum|add/i);
    expect(FUNCTION_DESCRIPTIONS.IF?.en).toMatch(/condition|true|false/i);
    expect(FUNCTION_DESCRIPTIONS.VLOOKUP?.en).toMatch(/lookup|column/i);
  });

  it('renders Japanese picker and argument labels from the i18n dictionary', () => {
    const handle = attachFxDialog({
      host,
      store: createSpreadsheetStore(),
      strings: ja,
      onInsert: () => {},
    });

    handle.open();
    expect(document.querySelector<HTMLElement>('.fc-fmtdlg__header')?.textContent).toBe(
      '関数の引数',
    );
    expect(document.querySelector<HTMLElement>('.fc-fxdialog__category-row')?.textContent).toContain(
      'カテゴリを選択',
    );
    expect(document.querySelector<HTMLInputElement>('.fc-fxdialog__search')?.placeholder).toBe(
      '関数を検索…',
    );
    const categoryLabels = Array.from(
      document.querySelectorAll<HTMLOptionElement>('.fc-fxdialog__category option'),
    ).map((option) => option.textContent ?? '');
    expect(categoryLabels).toEqual(
      expect.arrayContaining(['すべて', '最近使用した関数', '論理', '検索/行列']),
    );

    handle.open('IF');
    expect(document.querySelector<HTMLElement>('.fc-fxdialog__args-desc')?.textContent).toContain(
      '条件が真',
    );
    expect(document.querySelector<HTMLElement>('.fc-fxdialog__preview-label')?.textContent).toBe(
      '数式の結果',
    );
    expect(document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.textContent).toBe(
      '挿入',
    );

    handle.detach();
  });

  it('clicks on rendered picker items via event delegation (no per-item listeners)', () => {
    // Pre-refactor regression check: each render of the picker used to attach
    // a fresh `click` listener to every item, leaving 9 add / 7 remove pairs
    // in detach(). The delegated handler should fire whether the click hits
    // the item element directly or any of its children (name span / desc span).
    const handle = attachFxDialog({
      host,
      store: createSpreadsheetStore(),
      onInsert: () => {},
    });
    handle.open();
    const sumItem = Array.from(document.querySelectorAll<HTMLElement>('.fc-fxdialog__item')).find(
      (el) => el.dataset.fxName === 'SUM',
    );
    expect(sumItem).toBeTruthy();
    if (!sumItem) return;

    // Click on the name span (a child) — delegation must still resolve back
    // to the parent item via closest().
    const nameSpan = sumItem.querySelector<HTMLElement>('.fc-fxdialog__item-name');
    nameSpan?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const args = document.querySelector<HTMLElement>('.fc-fxdialog__args');
    expect(args?.hidden).toBe(false);
    const argName = document.querySelector<HTMLElement>('.fc-fxdialog__args-name');
    expect(argName?.textContent).toMatch(/^SUM\(/);
    handle.detach();
  });

  it('does not leak listeners across many search-filter rerenders', () => {
    // The picker rebuilds its item list on every keystroke. With the old
    // per-item listener approach, 200 keystrokes × ~600 items would
    // accumulate hundreds of thousands of listeners. With delegation, only
    // the single shell-tracked listener on `list` exists. Functional check:
    // after lots of rerenders the click still works exactly once.
    let inserted = 0;
    const handle = attachFxDialog({
      host,
      store: createSpreadsheetStore(),
      onInsert: () => {
        inserted += 1;
      },
    });
    handle.open();
    const search = document.querySelector<HTMLInputElement>('.fc-fxdialog__search');
    if (!search) throw new Error('expected search input');
    // End with the empty filter so SUM is back in the rendered list.
    for (const q of ['', 'S', 'SU', 'SUM', '', 'V', 'VL', 'VLO', '', 'I', 'IF', '']) {
      search.value = q;
      search.dispatchEvent(new Event('input'));
    }
    const sumItem = Array.from(document.querySelectorAll<HTMLElement>('.fc-fxdialog__item')).find(
      (el) => el.dataset.fxName === 'SUM',
    );
    expect(sumItem).toBeTruthy();
    sumItem?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    const insertBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    insertBtn?.click();
    expect(inserted).toBe(1);
    handle.detach();
  });

  it('detach removes the overlay and disables further listener firing', () => {
    let inserted = 0;
    const handle = attachFxDialog({
      host,
      store: createSpreadsheetStore(),
      onInsert: () => {
        inserted += 1;
      },
    });
    handle.open('SUM');
    const insertBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    expect(insertBtn).toBeTruthy();
    handle.detach();
    expect(document.querySelector('.fc-fxdialog')).toBeNull();
    // Stale reference should be inert.
    insertBtn?.click();
    expect(inserted).toBe(0);
  });
});
