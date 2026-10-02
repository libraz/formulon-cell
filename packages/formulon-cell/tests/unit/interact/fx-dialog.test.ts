import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { getRecentFunctions } from '../../../src/commands/function-history.js';
import { en, ja } from '../../../src/i18n/strings.js';
import { attachFxDialog, FUNCTION_DESCRIPTIONS } from '../../../src/interact/fx-dialog.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

const liveFunctionReader = (names: readonly string[]) => ({
  functionNames: () => names,
  functionMetadata: (name: string, locale: number) => {
    if (name === 'ACOS') {
      return {
        name,
        minArity: 1,
        maxArity: 1,
        localizedName: locale === 1 ? '逆余弦' : 'Arccosine',
        signatureTemplate: locale === 1 ? '逆余弦(number)' : 'ACOS(number)',
        description: locale === 1 ? '逆余弦を返します。' : 'Returns the arccosine.',
      };
    }
    if (name === 'ACCRINT') return { name, minArity: 0, maxArity: null };
    if (name === 'ZERO') return { name, minArity: 0, maxArity: 0 };
    return { name, minArity: 1, maxArity: null };
  },
});

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

  it('projects and clears the Insert disabled reason across picker and args steps', () => {
    const handle = attachFxDialog({
      host,
      store: createSpreadsheetStore(),
      strings: en,
      onInsert: () => {},
    });
    handle.open();
    const insertBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    expect(insertBtn?.disabled).toBe(true);
    expect(insertBtn?.dataset.disabledReason).toBe('Select a function before inserting it.');
    expect(insertBtn?.getAttribute('aria-description')).toBe(
      'Select a function before inserting it.',
    );

    const sumItem = Array.from(document.querySelectorAll<HTMLElement>('.fc-fxdialog__item')).find(
      (el) => el.dataset.fxName === 'SUM',
    );
    sumItem?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(insertBtn?.disabled).toBe(false);
    expect(insertBtn?.dataset.disabledReason).toBeUndefined();
    expect(insertBtn?.getAttribute('aria-description')).toBeNull();

    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn')?.click();
    expect(insertBtn?.disabled).toBe(true);
    expect(insertBtn?.dataset.disabledReason).toBe('Select a function before inserting it.');
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
      onInsert: (formula) => {
        inserted.push(formula);
      },
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
    expect(Array.from(category.options, (option) => option.value)).toEqual([
      'all',
      'recent',
      'financial',
      'logical',
      'text',
      'datetime',
      'lookup',
      'math',
      'statistical',
      'compatibility',
      'dynamicArray',
    ]);

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

  it('opens the statistical function family when requested', () => {
    const handle = attachFxDialog({
      host,
      store: createSpreadsheetStore(),
      onInsert: () => {},
    });
    handle.open(undefined, { category: 'statistical' });
    const names = Array.from(document.querySelectorAll<HTMLElement>('.fc-fxdialog__item-name')).map(
      (item) => item.textContent ?? '',
    );
    expect(names).toContain('AVERAGE');
    expect(names).toContain('COUNTIFS');
    expect(names).not.toContain('IF');
    handle.detach();
  });

  it('records a function after insertion and lists shared recents in MRU order', () => {
    const store = createSpreadsheetStore();
    const handle = attachFxDialog({
      host,
      store,
      onInsert: () => {},
    });

    handle.open('VLOOKUP');
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    handle.open('IF');
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    expect(getRecentFunctions(store)).toEqual(['IF', 'VLOOKUP']);

    handle.open(undefined, { category: 'recent' });
    const category = document.querySelector<HTMLSelectElement>('.fc-fxdialog__category');
    expect(category).toBeTruthy();
    if (!category) return;

    const names = Array.from(document.querySelectorAll<HTMLElement>('.fc-fxdialog__item-name')).map(
      (item) => item.textContent ?? '',
    );
    expect(names).toEqual(['IF', 'VLOOKUP']);
    handle.detach();
  });

  it('does not record a chosen function when the dialog is canceled before insertion', () => {
    const store = createSpreadsheetStore();
    const handle = attachFxDialog({
      host,
      store,
      onInsert: () => {},
    });

    handle.open();
    const lookup = Array.from(document.querySelectorAll<HTMLElement>('.fc-fxdialog__item')).find(
      (item) => item.dataset.fxName === 'VLOOKUP',
    );
    expect(lookup).toBeTruthy();
    lookup?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    handle.close();

    handle.open(undefined, { category: 'recent' });
    expect(getRecentFunctions(store)).toEqual([]);
    expect(document.querySelector('.fc-fxdialog__empty')?.textContent).toBeTruthy();
    handle.detach();
  });

  it('keeps the dialog open and does not record when insertion is rejected', () => {
    const store = createSpreadsheetStore();
    const handle = attachFxDialog({
      host,
      store,
      onInsert: () => false,
    });
    handle.open('SUM');
    const argInput = document.querySelector<HTMLInputElement>('.fc-fxdialog__arg-input');
    if (!argInput) throw new Error('expected SUM argument input');
    argInput.value = 'A1:A3';
    argInput.dispatchEvent(new Event('input'));
    const insertBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');

    insertBtn?.click();

    expect(document.querySelector<HTMLElement>('.fc-fxdialog')?.hidden).toBe(false);
    expect(document.querySelector<HTMLElement>('.fc-fxdialog__args')?.hidden).toBe(false);
    expect(argInput.value).toBe('A1:A3');
    expect(document.activeElement).toBe(insertBtn);
    expect(getRecentFunctions(store)).toEqual([]);
    handle.detach();
  });

  it('focuses seeded arguments immediately without stealing a later field focus', async () => {
    const handle = attachFxDialog({
      host,
      store: createSpreadsheetStore(),
      onInsert: () => {},
    });

    handle.open('IF');
    const inputs = document.querySelectorAll<HTMLInputElement>('.fc-fxdialog__arg-input');
    const first = inputs[0];
    const second = inputs[1];
    if (!first || !second) throw new Error('expected IF argument inputs');
    expect(document.activeElement).toBe(first);
    second.focus();
    await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));

    expect(document.activeElement).toBe(second);
    handle.detach();
  });

  it('keeps the dialog open and does not record when insertion throws', () => {
    const store = createSpreadsheetStore();
    const handle = attachFxDialog({
      host,
      store,
      onInsert: () => {
        throw new Error('write failed');
      },
    });
    handle.open('SUM');
    const insertBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');

    expect(() => insertBtn?.click()).not.toThrow();
    expect(document.querySelector<HTMLElement>('.fc-fxdialog')?.hidden).toBe(false);
    expect(document.activeElement).toBe(insertBtn);
    expect(getRecentFunctions(store)).toEqual([]);
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
    const firstName = document.querySelector<HTMLElement>(
      '.fc-fxdialog__summary-name',
    )?.textContent;
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
      onInsert: (f) => {
        inserted.push(f);
      },
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
      onInsert: (f) => {
        inserted.push(f);
      },
    });

    handle.open('IF');
    const ifInputs = document.querySelectorAll<HTMLInputElement>('.fc-fxdialog__arg-input');
    expect(ifInputs).toHaveLength(3);
    const [logicalTest, valueIfTrue, valueIfFalse] = Array.from(ifInputs);
    if (!logicalTest || !valueIfTrue || !valueIfFalse)
      throw new Error('expected IF argument inputs');
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

  it('uses the live catalog for engine-only names, metadata display, and canonical insertion', () => {
    const store = createSpreadsheetStore();
    const inserted: string[] = [];
    const liveNames = ['ACOS', 'SUM'] as const;
    const handle = attachFxDialog({
      host,
      store,
      getWb: () => liveFunctionReader(liveNames),
      getLocale: () => 'en-US',
      onInsert: (formula) => {
        inserted.push(formula);
      },
    });

    handle.open();
    const search = document.querySelector<HTMLInputElement>('.fc-fxdialog__search');
    if (!search) throw new Error('expected function search');
    search.value = 'arccos';
    search.dispatchEvent(new Event('input'));
    const item = document.querySelector<HTMLElement>('[data-fx-name="ACOS"]');
    expect(item?.textContent).toContain('Arccosine');
    item?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(document.querySelectorAll<HTMLInputElement>('.fc-fxdialog__arg-input')).toHaveLength(1);
    const input = document.querySelector<HTMLInputElement>('.fc-fxdialog__arg-input');
    if (!input) throw new Error('expected ACOS argument input');
    input.value = '0';
    input.dispatchEvent(new Event('input'));
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

    expect(inserted).toEqual(['=ACOS(0)']);
    expect(getRecentFunctions(store, new Set(liveNames))).toEqual(['ACOS']);
    handle.detach();
  });

  it('uses structural arity for exact zero and unbounded progressive arguments', () => {
    const handle = attachFxDialog({
      host,
      store: createSpreadsheetStore(),
      getWb: () => liveFunctionReader(['ACCRINT', 'ZERO']),
      onInsert: () => {},
    });

    handle.open('ACCRINT');
    expect(document.querySelectorAll<HTMLInputElement>('.fc-fxdialog__arg-input')).toHaveLength(1);
    expect(
      document.querySelector<HTMLButtonElement>('[data-fx-action="add-argument"]'),
    ).toBeTruthy();
    expect(
      document.querySelector<HTMLButtonElement>('[data-fx-action="remove-argument"]'),
    ).toBeTruthy();
    for (let i = 0; i < 2; i += 1) {
      document.querySelector<HTMLButtonElement>('[data-fx-action="add-argument"]')?.click();
    }
    expect(document.querySelectorAll<HTMLInputElement>('.fc-fxdialog__arg-input')).toHaveLength(3);
    for (let i = 0; i < 3; i += 1) {
      document.querySelector<HTMLButtonElement>('[data-fx-action="remove-argument"]')?.click();
    }
    expect(document.querySelectorAll<HTMLInputElement>('.fc-fxdialog__arg-input')).toHaveLength(0);

    handle.open('ZERO');
    expect(document.querySelectorAll<HTMLInputElement>('.fc-fxdialog__arg-input')).toHaveLength(0);
    handle.detach();
  });

  it('adds and removes finite optional arguments within the live arity bounds', () => {
    const handle = attachFxDialog({
      host,
      store: createSpreadsheetStore(),
      getWb: () => ({
        functionNames: () => ['FINITE'],
        functionMetadata: () => ({ name: 'FINITE', minArity: 1, maxArity: 3 }),
      }),
      onInsert: () => {},
    });
    const inputs = () => document.querySelectorAll<HTMLInputElement>('.fc-fxdialog__arg-input');
    const action = (name: string) =>
      document.querySelector<HTMLButtonElement>(`[data-fx-action="${name}-argument"]`);

    handle.open('FINITE');
    expect(inputs()).toHaveLength(1);
    expect(action('remove')).toBeNull();
    const requiredInput = inputs()[0];
    if (!requiredInput) throw new Error('expected the required argument');
    requiredInput.value = 'A1';
    action('add')?.click();
    expect(inputs()).toHaveLength(2);
    expect(inputs()[0]?.value).toBe('A1');
    action('add')?.click();
    expect(inputs()).toHaveLength(3);
    expect(action('add')).toBeNull();
    action('remove')?.click();
    expect(inputs()).toHaveLength(2);
    action('remove')?.click();
    expect(inputs()).toHaveLength(1);
    expect(action('remove')).toBeNull();
    expect(inputs()[0]?.value).toBe('A1');
    handle.detach();
  });

  it('preserves interior argument blanks while omitting trailing blanks', () => {
    const inserted: string[] = [];
    const handle = attachFxDialog({
      host,
      store: createSpreadsheetStore(),
      getWb: () => liveFunctionReader(['ACCRINT']),
      onInsert: (formula) => {
        inserted.push(formula);
      },
    });
    handle.open('ACCRINT');
    for (let index = 0; index < 3; index += 1) {
      document.querySelector<HTMLButtonElement>('[data-fx-action="add-argument"]')?.click();
    }
    const inputs = document.querySelectorAll<HTMLInputElement>('.fc-fxdialog__arg-input');
    expect(inputs).toHaveLength(4);
    const [first, , third, trailing] = Array.from(inputs);
    if (!first || !third || !trailing) throw new Error('expected four argument fields');
    first.value = '7';
    third.value = '3';
    trailing.value = '';
    third.dispatchEvent(new Event('input'));
    expect(document.querySelector('.fc-fxdialog__preview')?.textContent).toBe('=ACCRINT(7, , 3)');
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    expect(inserted).toEqual(['=ACCRINT(7, , 3)']);
    handle.detach();
  });

  it('rebuilds the live catalog on every open and restores stale Recent entries by workbook', () => {
    const store = createSpreadsheetStore();
    let reader = liveFunctionReader(['ACOS', 'SUM']);
    const handle = attachFxDialog({
      host,
      store,
      getWb: () => reader,
      onInsert: () => {},
    });

    handle.open('ACOS');
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    reader = liveFunctionReader(['SUM']);
    handle.open('ACOS');
    expect(document.querySelector<HTMLElement>('.fc-fxdialog__picker')?.hidden).toBe(false);
    expect(document.querySelector<HTMLElement>('[data-fx-name="ACOS"]')).toBeNull();
    handle.open(undefined, { category: 'recent' });
    expect(document.querySelector<HTMLElement>('[data-fx-name="ACOS"]')).toBeNull();

    reader = liveFunctionReader(['ACOS', 'SUM']);
    handle.open(undefined, { category: 'recent' });
    expect(document.querySelector<HTMLElement>('[data-fx-name="ACOS"]')).toBeTruthy();
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
    expect(
      document.querySelector<HTMLElement>('.fc-fxdialog__category-row')?.textContent,
    ).toContain('カテゴリを選択');
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

  it('keeps Function Arguments on compact desktop dialog geometry', () => {
    const css = readFileSync(
      join(root, 'src/styles/core/app/dialog-modules/fx-dialog.css'),
      'utf8',
    );

    expect(css).toMatch(/\.fc-fxdialog__item--active\s*\{[\s\S]*?background: var\(--fc-bg-hover/);
    expect(css).toMatch(
      /\.fc-fxdialog__item--active\s*\{[\s\S]*?box-shadow: inset 0 0 0 1px var\(--fc-fmtdlg-list-focus-border/,
    );
    expect(css).toMatch(/\.fc-fxdialog__args-header\s*\{[\s\S]*?border-radius: 2px 2px 0 0;/);
    expect(css).toMatch(/\.fc-fxdialog__args-fields\s*\{[\s\S]*?border-radius: 0 0 2px 2px;/);
    expect(css).toMatch(/\.fc-fxdialog__preview-label\s*\{[\s\S]*?letter-spacing: 0;/);
    expect(css).toMatch(/\.fc-fxdialog__preview\s*\{[\s\S]*?border-radius: 2px;/);
    expect(css).not.toContain('background: #e7f6ed');
    expect(css).not.toContain('letter-spacing: 0.04em');
  });
});
