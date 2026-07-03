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
  Array.from(document.querySelectorAll<HTMLInputElement>('.fc-goto__kinds input[type="radio"]'));
const scopeRadios = (): HTMLInputElement[] =>
  Array.from(document.querySelectorAll<HTMLInputElement>('.fc-goto__scope input[type="radio"]'));
const valueFilterChecks = (): HTMLInputElement[] =>
  Array.from(
    document.querySelectorAll<HTMLInputElement>('.fc-goto__value-filters input[type="checkbox"]'),
  );
const referenceInput = (): HTMLInputElement | null =>
  document.querySelector<HTMLInputElement>('.fc-goto__reference input');
const okBtn = (): HTMLButtonElement | null =>
  document.querySelector<HTMLButtonElement>('.fc-goto .fc-fmtdlg__btn--primary');
const status = (): HTMLElement | null => document.querySelector<HTMLElement>('.fc-goto__status');

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

  it('renders kind radios in the canonical order with constants pre-checked', () => {
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open();
    const radios = kindRadios();
    expect(radios.map((r) => r.value)).toEqual([
      'blanks',
      'non-blanks',
      'formulas',
      'constants',
      'numbers',
      'text',
      'errors',
      'data-validation',
      'conditional-format',
    ]);
    const checked = radios.find((r) => r.checked);
    expect(checked?.value).toBe('constants');
    expect(document.querySelector<HTMLElement>('.fc-goto__value-filters')?.hidden).toBe(false);
    expect(valueFilterChecks().map((c) => c.value)).toEqual([
      'numbers',
      'text',
      'logical',
      'errors',
    ]);
    expect(valueFilterChecks().every((c) => c.checked)).toBe(true);
    handle.detach();
  });

  it('disables the selection-scope radio when the current selection is a single cell', () => {
    setRange(store, 0, 0, 0, 0);
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open();
    const [sheetRadio, selectionRadio] = scopeRadios();
    const reason = defaultStrings.goToDialog.scopeSelectionRequiresMultiCell;
    expect(sheetRadio?.checked).toBe(true);
    expect(selectionRadio?.disabled).toBe(true);
    expect(selectionRadio?.dataset.disabledReason).toBe(reason);
    expect(selectionRadio?.getAttribute('aria-description')).toBe(reason);
    expect(selectionRadio?.title).toBe(reason);
    expect(selectionRadio?.closest('label')?.title).toBe(
      `${defaultStrings.goToDialog.scopeSelection}\n${reason}`,
    );
    handle.detach();
  });

  it('enables the selection-scope radio when the selection covers more than one cell', () => {
    setRange(store, 0, 0, 4, 4);
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open();
    const [, selectionRadio] = scopeRadios();
    expect(selectionRadio?.disabled).toBe(false);
    expect(selectionRadio?.dataset.disabledReason).toBeUndefined();
    expect(selectionRadio?.hasAttribute('aria-description')).toBe(false);
    expect(selectionRadio?.title).toBe('');
    expect(selectionRadio?.closest('label')?.title).toBe(defaultStrings.goToDialog.scopeSelection);
    handle.detach();
  });

  it('OK with no matches keeps the dialog open and sets the no-results status', () => {
    // Sheet has no formula cells.
    sync(store, wb);
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open();
    const radios = kindRadios();
    const formulaRadio = radios.find((r) => r.value === 'formulas');
    if (formulaRadio) {
      formulaRadio.checked = true;
    }
    okBtn()?.click();
    expect(status()?.textContent).not.toBe('');
    expect(overlay()?.hidden).toBe(false);
    handle.detach();
  });

  it('OK with 1+ matches closes the dialog and selects the exact matched cells', () => {
    wb.setNumber({ sheet: 0, row: 2, col: 1 }, 11);
    wb.setNumber({ sheet: 0, row: 5, col: 4 }, 22);
    wb.recalc();
    sync(store, wb);
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open();
    const numRadio = kindRadios().find((r) => r.value === 'numbers');
    if (numRadio) numRadio.checked = true;
    okBtn()?.click();
    expect(overlay()?.hidden).toBe(true);
    const sel = store.getState().selection;
    expect(sel.active).toEqual({ sheet: 0, row: 2, col: 1 });
    expect(sel.range).toEqual({ sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 });
    expect(sel.extraRanges).toEqual([{ sheet: 0, r0: 5, c0: 4, r1: 5, c1: 4 }]);
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

  it('normal Go To jumps to a typed reference range', () => {
    const handle = attachGoToDialog({ host, store, getWb: () => wb });
    handle.open('go-to');
    expect(document.body.textContent).toContain('参照先');
    expect(referenceInput()?.hidden).toBe(false);
    expect(kindRadios()[0]?.closest('.fc-goto__kinds')?.hasAttribute('hidden')).toBe(true);
    expect(document.querySelector<HTMLElement>('.fc-goto__value-filters')?.hidden).toBe(true);
    const input = referenceInput();
    expect(input).toBeTruthy();
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
    expect(status()?.textContent).toContain('有効なセル参照');
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
