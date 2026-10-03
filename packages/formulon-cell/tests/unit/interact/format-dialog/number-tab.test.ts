import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../../src/commands/history.js';
import { addrKey } from '../../../../src/engine/workbook-handle.js';
import { attachFormatDialog } from '../../../../src/interact/format-dialog.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../../src/store/store.js';
import { root, setActive } from './fixtures.js';

describe('attachFormatDialog', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;

  beforeEach(() => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    setActive(store, 0, 0);
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('clicking number category updates draft and visibility', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const decimalsRow = document.querySelector<HTMLLabelElement>(
      '.fc-fmtdlg__cat-controls .fc-fmtdlg__row',
    );
    expect(decimalsRow?.hidden).toBe(true); // general → hidden

    const fixedBtn = document.querySelector<HTMLButtonElement>('button[data-fc-cat="fixed"]');
    fixedBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(fixedBtn?.getAttribute('aria-selected')).toBe('true');
    expect(decimalsRow?.hidden).toBe(false);
    const thousands = document.querySelector<HTMLInputElement>('input[data-fc-check="thousands"]');
    expect(thousands?.closest('label')?.hidden).toBe(false);

    const symbolRow = document.querySelectorAll<HTMLLabelElement>(
      '.fc-fmtdlg__cat-controls .fc-fmtdlg__row',
    )[1];
    expect(symbolRow?.hidden).toBe(true);

    const currencyBtn = document.querySelector<HTMLButtonElement>('button[data-fc-cat="currency"]');
    currencyBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(symbolRow?.hidden).toBe(false);

    const percentBtn = document.querySelector<HTMLButtonElement>('button[data-fc-cat="percent"]');
    percentBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(symbolRow?.hidden).toBe(true);
    expect(decimalsRow?.hidden).toBe(false);
    expect(thousands?.closest('label')?.hidden).toBe(true);

    handle.detach();
  });

  it('lists number categories in Excel order without the custom Date & Time category', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    expect(
      Array.from(document.querySelectorAll<HTMLButtonElement>('button[data-fc-cat]')).map(
        (button) => button.dataset.fcCat,
      ),
    ).toEqual([
      'general',
      'fixed',
      'currency',
      'accounting',
      'date',
      'time',
      'percent',
      'fraction',
      'scientific',
      'text',
      'special',
      'custom',
    ]);
    expect(document.querySelector('button[data-fc-cat="datetime"]')).toBeNull();
    handle.detach();
  });

  it('persists and reopens all nine Excel fraction presets as Fraction', () => {
    const patterns = [
      '# ?/?',
      '# ??/??',
      '# ???/???',
      '# ?/2',
      '# ?/4',
      '# ?/8',
      '# ?/16',
      '# ?/10',
      '# ?/100',
    ];
    const handle = attachFormatDialog({ host, store });
    handle.open();
    document
      .querySelector<HTMLButtonElement>('button[data-fc-cat="fraction"]')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    const options = Array.from(
      document.querySelectorAll<HTMLButtonElement>('button[data-fc-pattern]'),
    );
    expect(options.map((button) => button.dataset.fcPattern)).toEqual(patterns);
    const selectedPattern = '# ?/16';
    options
      .find((button) => button.dataset.fcPattern === selectedPattern)
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    document
      .querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.numFmt,
    ).toEqual({ kind: 'custom', pattern: selectedPattern });

    handle.open();
    expect(
      document
        .querySelector<HTMLButtonElement>('button[data-fc-cat="fraction"]')
        ?.getAttribute('aria-selected'),
    ).toBe('true');
    expect(
      Array.from(document.querySelectorAll<HTMLButtonElement>('button[data-fc-pattern]'))
        .find((button) => button.dataset.fcPattern === selectedPattern)
        ?.getAttribute('aria-selected'),
    ).toBe('true');
    handle.detach();
  });

  it('Number category can persist Excel-style 1000 separator option', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    document
      .querySelector<HTMLButtonElement>('button[data-fc-cat="fixed"]')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    const decimalsInput = document.querySelector<HTMLInputElement>(
      'input[type="number"][min="0"][max="10"]',
    );
    const thousands = document.querySelector<HTMLInputElement>('input[data-fc-check="thousands"]');
    if (!decimalsInput || !thousands) throw new Error('number controls missing');
    decimalsInput.value = '2';
    decimalsInput.dispatchEvent(new Event('input', { bubbles: true }));
    thousands.checked = true;
    thousands.dispatchEvent(new Event('change', { bubbles: true }));

    document
      .querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.numFmt,
    ).toEqual({
      kind: 'fixed',
      decimals: 2,
      thousands: true,
    });
    handle.detach();
  });

  it('Number category can persist Excel-style negative number style', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    document
      .querySelector<HTMLButtonElement>('button[data-fc-cat="fixed"]')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    const redParens = document.querySelector<HTMLButtonElement>(
      'button[data-fc-negative-style="red-parens"]',
    );
    if (!redParens) throw new Error('negative style option missing');
    redParens.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(redParens.getAttribute('aria-selected')).toBe('true');

    document
      .querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.numFmt,
    ).toEqual({
      kind: 'fixed',
      decimals: 2,
      negativeStyle: 'red-parens',
    });
    handle.detach();
  });

  it('Number category persists and rehydrates Special formats as Special', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    document
      .querySelector<HTMLButtonElement>('button[data-fc-cat="special"]')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    const preset = document.querySelector<HTMLSelectElement>(
      'select[data-fc-select="patternPreset"]',
    );
    if (!preset) throw new Error('special preset select missing');
    preset.value = '00000-0000';
    preset.dispatchEvent(new Event('change', { bubbles: true }));

    document
      .querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.numFmt,
    ).toEqual({ kind: 'special', pattern: '00000-0000' });
    handle.detach();

    const reopened = attachFormatDialog({ host, store });
    reopened.open();
    expect(
      document
        .querySelector<HTMLButtonElement>('button[data-fc-cat="special"]')
        ?.getAttribute('aria-selected'),
    ).toBe('true');
    reopened.detach();
  });

  it('clicking on cat list outside button is a no-op', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();
    const catList = document.querySelector<HTMLElement>('.fc-fmtdlg__cat');
    catList?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    const generalBtn = document.querySelector<HTMLButtonElement>('button[data-fc-cat="general"]');
    expect(generalBtn?.getAttribute('aria-selected')).toBe('true');
    handle.detach();
  });

  it('number categories support Excel-style arrow, Home, and End keyboard navigation', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const generalBtn = document.querySelector<HTMLButtonElement>('button[data-fc-cat="general"]');
    const fixedBtn = document.querySelector<HTMLButtonElement>('button[data-fc-cat="fixed"]');
    const customBtn = document.querySelector<HTMLButtonElement>('button[data-fc-cat="custom"]');
    expect(generalBtn?.tabIndex).toBe(0);
    expect(fixedBtn?.tabIndex).toBe(-1);

    generalBtn?.focus();
    generalBtn?.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    expect(fixedBtn?.getAttribute('aria-selected')).toBe('true');
    expect(fixedBtn?.tabIndex).toBe(0);
    expect(document.activeElement).toBe(fixedBtn);

    fixedBtn?.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true }));
    expect(customBtn?.getAttribute('aria-selected')).toBe('true');
    expect(document.activeElement).toBe(customBtn);

    customBtn?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', bubbles: true }));
    expect(generalBtn?.getAttribute('aria-selected')).toBe('true');
    expect(document.activeElement).toBe(generalBtn);
    handle.detach();
  });

  it('decimals input clamps to [0, 10]', () => {
    const history = new History();
    const handle = attachFormatDialog({ host, store, history });
    handle.open();
    const fixedBtn = document.querySelector<HTMLButtonElement>('button[data-fc-cat="fixed"]');
    fixedBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const decimalsInput = document.querySelector<HTMLInputElement>(
      'input[type="number"][min="0"][max="10"]',
    ) as HTMLInputElement;
    decimalsInput.value = '99';
    decimalsInput.dispatchEvent(new Event('input', { bubbles: true }));

    const okBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.numFmt).toEqual({ kind: 'fixed', decimals: 10 });
    handle.detach();
  });

  it('decimals input ignores non-numeric values', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();
    const fixedBtn = document.querySelector<HTMLButtonElement>('button[data-fc-cat="fixed"]');
    fixedBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const decimalsInput = document.querySelector<HTMLInputElement>(
      'input[type="number"][min="0"][max="10"]',
    ) as HTMLInputElement;
    decimalsInput.value = 'abc';
    decimalsInput.dispatchEvent(new Event('input', { bubbles: true }));

    const okBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.numFmt).toEqual({ kind: 'fixed', decimals: 2 });
    handle.detach();
  });

  it('symbol select updates currency symbol', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();
    const currencyBtn = document.querySelector<HTMLButtonElement>('button[data-fc-cat="currency"]');
    currencyBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const symbolSelect = document.querySelector<HTMLSelectElement>('select') as HTMLSelectElement;
    symbolSelect.value = '¥';
    symbolSelect.dispatchEvent(new Event('change', { bubbles: true }));

    const okBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.numFmt).toEqual({ kind: 'currency', decimals: 2, symbol: '¥' });
    handle.detach();
  });

  it('keeps number pattern list buttons on the shared dialog option primitive', () => {
    const source = readFileSync(join(root, 'src/interact/format-dialog.ts'), 'utf8');
    expect(source).toContain('appendDialogOptionButton(patternList');
    expect(source).not.toContain("const item = document.createElement('button')");
  });
});
