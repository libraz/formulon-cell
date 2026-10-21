import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { defaultStrings } from '../../../src/i18n/strings.js';
import { attachGoToDialog } from '../../../src/interact/goto-dialog.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const overlay = (): HTMLElement | null => document.querySelector<HTMLElement>('.fc-goto');
const kindRadios = (): HTMLInputElement[] =>
  Array.from(
    document.querySelectorAll<HTMLInputElement>(
      '.fc-goto__kinds input[type="radio"]:not([data-sub])',
    ),
  );
const subRadios = (sub: 'levels' | 'rule'): HTMLInputElement[] =>
  Array.from(document.querySelectorAll<HTMLInputElement>(`input[data-sub="${sub}"]`));
const valueFilterChecks = (): HTMLInputElement[] =>
  Array.from(
    document.querySelectorAll<HTMLInputElement>('.fc-goto__value-filters input[type="checkbox"]'),
  );
const referenceInput = (): HTMLInputElement | null =>
  document.querySelector<HTMLInputElement>('.fc-goto__reference input');
const okBtn = (): HTMLButtonElement | null =>
  document.querySelector<HTMLButtonElement>('.fc-goto .fc-fmtdlg__btn--primary');
const specialBtn = (): HTMLButtonElement | undefined =>
  Array.from(document.querySelectorAll<HTMLButtonElement>('.fc-goto .fc-fmtdlg__btn')).find(
    (b) => b.textContent === defaultStrings.goToDialog.special,
  );
const typeReference = (text: string): void => {
  const input = referenceInput();
  if (!input) throw new Error('missing reference input');
  input.value = text;
};
const status = (): HTMLElement | null => document.querySelector<HTMLElement>('.fc-goto__status');
const targetItems = (): string[] =>
  Array.from(document.querySelectorAll<HTMLElement>('.fc-goto__targets [role="option"]')).map(
    (el) => el.textContent ?? '',
  );
const alertText = (): string =>
  document.querySelector<HTMLElement>('.fc-valdlg:not([hidden]) .fc-tb__dlg__message')
    ?.textContent ?? '';

const defineNames = (wb: WorkbookHandle, names: Record<string, string>): void => {
  wb.definedNames = function* () {
    for (const [name, formula] of Object.entries(names)) yield { name, formula, localSheetId: -1 };
  };
};

const sync = (store: SpreadsheetStore, wb: WorkbookHandle): void => {
  store.setState((s) => {
    const cells = new Map<
      string,
      { value: ReturnType<WorkbookHandle['getValue']>; formula: string | null }
    >();
    for (const e of wb.cells(0)) {
      cells.set(addrKey(e.addr), { value: e.value, formula: e.formula });
    }
    return { ...s, data: { ...s.data, cells } };
  });
};

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
      ...s.selection,
      active: { sheet: 0, row: r0, col: c0 },
      anchor: { sheet: 0, row: r0, col: c0 },
      range: { sheet: 0, r0, c0, r1, c1 },
    },
  }));
};

describe('attachGoToDialog', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    host = document.createElement('div');
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  afterEach(() => {
    while (document.body.firstChild) document.body.removeChild(document.body.firstChild);
  });

  it('renders the options in the two-column order with constants pre-checked', () => {
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open();
    expect(kindRadios().map((r) => r.value)).toEqual([
      'comments',
      'constants',
      'formulas',
      'blanks',
      'current-region',
      'current-array',
      'row-differences',
      'column-differences',
      'precedents',
      'dependents',
      'last-cell',
      'visible-cells',
      'conditional-format',
      'data-validation',
    ]);
    expect(kindRadios().find((r) => r.checked)?.value).toBe('constants');
    expect(valueFilterChecks().map((c) => c.value)).toEqual([
      'numbers',
      'text',
      'logical',
      'errors',
    ]);
    expect(valueFilterChecks().every((c) => c.checked && !c.disabled)).toBe(true);
    expect(subRadios('levels').every((r) => r.disabled)).toBe(true);
    expect(subRadios('rule').every((r) => r.disabled)).toBe(true);
    handle.detach();
  });

  it('enables the sub-options only for the options they belong to', () => {
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open();
    const pick = (kind: string): void => {
      const radio = kindRadios().find((r) => r.value === kind) as HTMLInputElement;
      radio.checked = true;
      radio.dispatchEvent(new Event('change', { bubbles: true }));
    };
    pick('blanks');
    expect(valueFilterChecks().every((c) => c.disabled)).toBe(true);
    pick('dependents');
    expect(subRadios('levels').every((r) => !r.disabled)).toBe(true);
    pick('data-validation');
    expect(subRadios('rule').every((r) => !r.disabled)).toBe(true);
    expect(subRadios('levels').every((r) => r.disabled)).toBe(true);
    handle.detach();
  });

  it('OK with no matches closes the dialog and raises the no-cells alert', () => {
    sync(store, wb);
    setRange(store, 3, 3, 3, 3);
    const before = store.getState().selection;
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open();
    const formulaRadio = kindRadios().find((r) => r.value === 'formulas');
    if (formulaRadio) formulaRadio.checked = true;
    okBtn()?.click();
    expect(overlay()?.hidden).toBe(true);
    expect(alertText()).toBe(defaultStrings.goToDialog.noResults);
    expect(store.getState().selection).toBe(before);
    handle.detach();
  });

  it('OK with a too-large result raises the too-large alert', () => {
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    wb.recalc();
    sync(store, wb);
    setRange(store, 0, 0, 1048575, 0);
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open();
    const blanks = kindRadios().find((r) => r.value === 'blanks');
    if (blanks) blanks.checked = true;
    okBtn()?.click();
    expect(alertText()).toBe(defaultStrings.goToDialog.selectionTooLarge);
    handle.detach();
  });

  it('OK with matches selects merged rectangles and activates the first top-left', () => {
    wb.setNumber({ sheet: 0, row: 2, col: 1 }, 11);
    wb.setNumber({ sheet: 0, row: 2, col: 2 }, 12);
    wb.setNumber({ sheet: 0, row: 3, col: 1 }, 13);
    wb.setNumber({ sheet: 0, row: 3, col: 2 }, 14);
    wb.setNumber({ sheet: 0, row: 5, col: 4 }, 22);
    wb.recalc();
    sync(store, wb);
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open();
    okBtn()?.click();
    expect(overlay()?.hidden).toBe(true);
    const sel = store.getState().selection;
    expect(sel.active).toEqual({ sheet: 0, row: 2, col: 1 });
    expect(sel.range).toEqual({ sheet: 0, r0: 2, c0: 1, r1: 3, c1: 2 });
    expect(sel.extraRanges).toEqual([{ sheet: 0, r0: 5, c0: 4, r1: 5, c1: 4 }]);
    handle.detach();
  });

  it('a multi-cell selection confines the search to itself', () => {
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    wb.setNumber({ sheet: 0, row: 8, col: 8 }, 2);
    wb.recalc();
    sync(store, wb);
    setRange(store, 0, 0, 2, 2);
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open();
    okBtn()?.click();
    const sel = store.getState().selection;
    expect(sel.range).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    expect(sel.extraRanges).toEqual([]);
    handle.detach();
  });

  it('OK applies constants value-kind filters before selecting matches', () => {
    wb.setNumber({ sheet: 0, row: 1, col: 1 }, 11);
    wb.setText({ sheet: 0, row: 2, col: 2 }, 'text');
    wb.setBool({ sheet: 0, row: 3, col: 3 }, true);
    wb.recalc();
    sync(store, wb);
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open();
    for (const check of valueFilterChecks()) {
      check.checked = check.value === 'logical';
    }
    okBtn()?.click();
    const sel = store.getState().selection;
    expect(sel.active).toEqual({ sheet: 0, row: 3, col: 3 });
    expect(sel.extraRanges).toEqual([]);
    handle.detach();
  });

  it('Go To lists range names then recent references and fills the reference on click', () => {
    defineNames(wb, { Zeta: '=Sheet1!$B$2:$C$3', alpha: '=Sheet1!$A$1' });
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open('go-to');
    expect(targetItems()).toEqual(['alpha', 'Zeta']);
    typeReference('B2');
    okBtn()?.click();
    handle.open('go-to');
    expect(targetItems()).toEqual(['alpha', 'Zeta', '$B$2']);
    document.querySelectorAll<HTMLElement>('.fc-goto__targets [role="option"]')[1]?.click();
    expect(referenceInput()?.value).toBe('Zeta');
    handle.detach();
  });

  it('keeps only the last four recent references, newest first', () => {
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    for (const ref of ['A1', 'B1', 'C1', 'D1', 'E1', 'C1']) {
      handle.open('go-to');
      typeReference(ref);
      okBtn()?.click();
    }
    handle.open('go-to');
    expect(targetItems()).toEqual(['$C$1', '$E$1', '$D$1', '$B$1']);
    handle.detach();
  });

  it('Go To accepts defined names, sheet-qualified and multi-area references', () => {
    defineNames(wb, { Tax: '=Sheet1!$B$2:$C$3' });
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open('go-to');
    typeReference('Tax');
    okBtn()?.click();
    expect(store.getState().selection.range).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 });

    handle.open('go-to');
    typeReference('A1,C3:D4');
    okBtn()?.click();
    const sel = store.getState().selection;
    expect(sel.active).toEqual({ sheet: 0, row: 0, col: 0 });
    expect(sel.range).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    expect(sel.extraRanges).toEqual([{ sheet: 0, r0: 2, c0: 2, r1: 3, c1: 3 }]);
    handle.detach();
  });

  it('Go To switches sheet for a sheet-qualified reference', () => {
    wb.addSheet('My Sheet');
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open('go-to');
    typeReference("'My Sheet'!B4");
    okBtn()?.click();
    expect(store.getState().data.sheetIndex).toBe(1);
    expect(store.getState().selection.active).toEqual({ sheet: 1, row: 3, col: 1 });
    store.setState((s) => ({ ...s, data: { ...s.data, sheetIndex: 0 } }));
    handle.open('go-to');
    expect(targetItems()).toEqual(["'My Sheet'!$B$4"]);
    handle.detach();
  });

  it('the Special button opens Go To Special', () => {
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open('go-to');
    expect(document.body.textContent).toContain('参照先');
    expect(kindRadios()[0]?.closest('.fc-goto__kinds')?.hasAttribute('hidden')).toBe(true);
    specialBtn()?.click();
    expect(overlay()?.hidden).toBe(false);
    expect(overlay()?.querySelector('.fc-fmtdlg__header')?.textContent).toBe(
      defaultStrings.goToDialog.title,
    );
    expect(referenceInput()?.closest('label')?.hidden).toBe(true);
    handle.detach();
  });

  it('normal Go To jumps to a typed reference range', () => {
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open('go-to');
    const input = referenceInput();
    if (!input) throw new Error('missing reference input');
    input.value = 'B2:D4';
    okBtn()?.click();
    expect(overlay()?.hidden).toBe(true);
    const sel = store.getState().selection;
    expect(sel.active).toEqual({ sheet: 0, row: 1, col: 1 });
    expect(sel.range).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 3, c1: 3 });
    handle.detach();
  });

  it('normal Go To uses the shared range picker for the reference input', () => {
    setRange(store, 1, 1, 3, 3);
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open('go-to');

    const picker = document.querySelector<HTMLButtonElement>(
      '[data-range-picker="go-to-reference"]',
    );
    const input = referenceInput();
    expect(picker?.getAttribute('aria-label')).toBe('範囲の選択');
    picker?.click();
    expect(input?.value).toBe('B2:D4');
    expect(picker?.getAttribute('aria-pressed')).toBe('true');
    expect(overlay()?.classList.contains('fc-fmtdlg--range-picking')).toBe(true);

    setRange(store, 4, 2, 6, 4);
    expect(input?.value).toBe('C5:E7');
    handle.close();
    expect(picker?.getAttribute('aria-pressed')).toBe('false');
    expect(overlay()?.classList.contains('fc-fmtdlg--range-picking')).toBe(false);
    handle.detach();
  });

  it('normal Go To keeps the dialog open on an invalid reference', () => {
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open('go-to');
    const input = referenceInput();
    if (input) input.value = 'not a ref';
    okBtn()?.click();
    expect(status()?.textContent).toBe(defaultStrings.goToDialog.invalidReference);
    expect(overlay()?.hidden).toBe(false);
    handle.detach();
  });

  it('Escape closes the dialog without mutating the selection', () => {
    setRange(store, 1, 1, 1, 1);
    const before = { ...store.getState().selection.active };
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open();
    overlay()?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(overlay()?.hidden).toBe(true);
    expect(store.getState().selection.active).toEqual(before);
    handle.detach();
  });

  it('detach removes the overlay node', () => {
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.detach();
    expect(overlay()).toBeNull();
  });

  it('keeps Go To Special controls on compact desktop dialog geometry', () => {
    const css = readFileSync(
      join(root, 'src/styles/core/app/dialog-modules/goto-special.css'),
      'utf8',
    );

    expect(css).toMatch(
      /\.fc-goto__reference input\s*\{[\s\S]*?min-height: 24px;[\s\S]*?border-radius: 2px;/,
    );
    expect(css).toMatch(
      /\.fc-goto__kinds\s*\{[\s\S]*?border-radius: 2px;[\s\S]*?padding: 6px 8px;/,
    );
    expect(css).toMatch(/\.fc-goto__radio\s*\{[\s\S]*?padding: 2px 4px;[\s\S]*?border-radius: 0;/);
    expect(css).toMatch(/\.fc-goto__radio:hover\s*\{[\s\S]*?background: var\(--fc-bg-hover/);
    expect(css).toMatch(/\.fc-goto__value-filters\s*\{[\s\S]*?border-radius: 2px;/);
    expect(css).toMatch(/\.fc-goto__check\s*\{[\s\S]*?padding: 2px 4px;[\s\S]*?border-radius: 0;/);
    expect(css).toMatch(/\.fc-goto__check:hover\s*\{[\s\S]*?background: var\(--fc-bg-hover/);
    expect(css).not.toContain('background: var(--fc-accent-soft');
  });
});
