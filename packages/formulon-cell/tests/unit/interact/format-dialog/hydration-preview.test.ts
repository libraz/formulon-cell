import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { CellValue } from '../../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { en } from '../../../../src/i18n/strings/en.js';
import { attachFormatDialog } from '../../../../src/interact/format-dialog.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import { setActive } from './fixtures.js';

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

  it('hydrates draft from active cell format on open', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        numFmt: { kind: 'fixed', decimals: 4, thousands: true },
        align: 'right',
        shrinkToFit: true,
        bold: true,
        italic: true,
        underline: true,
        strike: true,
        fontFamily: 'Georgia',
        fontSize: 14,
        color: '#ff0000',
        fill: '#00ff00',
        borders: { top: true, right: true, bottom: false, left: false },
      },
    );
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const decimalsInput = document.querySelector<HTMLInputElement>(
      'input[type="number"][min="0"][max="10"]',
    );
    expect(decimalsInput?.value).toBe('4');
    expect(
      document.querySelector<HTMLInputElement>('input[data-fc-check="thousands"]')?.checked,
    ).toBe(true);

    const boldInput = document.querySelector<HTMLInputElement>('input[data-fc-check="bold"]');
    expect(boldInput?.checked).toBe(true);

    const familyInput = document.querySelector<HTMLInputElement>('input[data-fc-input="family"]');
    expect(familyInput?.value).toBe('Georgia');

    const sizeInput = document.querySelector<HTMLInputElement>(
      'input[type="number"][min="1"][max="409"]',
    );
    expect(sizeInput?.value).toBe('14');

    const rightAlign = document.querySelector<HTMLInputElement>(
      'input[type="radio"][value="right"]',
    );
    expect(rightAlign?.checked).toBe(true);
    expect(
      document.querySelector<HTMLInputElement>('input[data-fc-check="shrinkToFit"]')?.checked,
    ).toBe(true);

    handle.detach();
  });

  it('hydrates draft from currency numFmt with symbol', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { numFmt: { kind: 'currency', decimals: 3, symbol: '€' } },
    );
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const decimalsInput = document.querySelector<HTMLInputElement>(
      'input[type="number"][min="0"][max="10"]',
    );
    expect(decimalsInput?.value).toBe('3');
    const symbolSelect = document.querySelector<HTMLSelectElement>('select');
    expect(symbolSelect?.value).toBe('€');
    handle.detach();
  });

  it('hydrates draft from percent numFmt', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { numFmt: { kind: 'percent', decimals: 1 } },
    );
    const handle = attachFormatDialog({ host, store });
    handle.open();
    const decimalsInput = document.querySelector<HTMLInputElement>(
      'input[type="number"][min="0"][max="10"]',
    );
    expect(decimalsInput?.value).toBe('1');
    handle.detach();
  });

  it('falls back to general/defaults when active cell has no format', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();
    const decimalsInput = document.querySelector<HTMLInputElement>(
      'input[type="number"][min="0"][max="10"]',
    );
    expect(decimalsInput?.value).toBe('2');
    const symbolSelect = document.querySelector<HTMLSelectElement>('select');
    expect(symbolSelect?.value).toBe('$');
    handle.detach();
  });

  it('previews the active numeric value instead of a synthetic sample', () => {
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'number', value: -1234.5 });
    const handle = attachFormatDialog({ host, store });
    handle.open();

    expect(document.querySelector<HTMLElement>('.fc-fmtdlg__preview-cell')?.textContent).toBe(
      '-1234.5',
    );
    handle.detach();
  });

  it.each([
    [123456789012, '1.23457E+11'],
    [-123456789012, '-1.23457E+11'],
    [0.0000000001234, '1.23400E-10'],
    [-0.0000000001234, '-1.23400E-10'],
    [1234.5, '1234.5'],
    [0, '0'],
  ])('previews General value %s with the shared numeric notation', (value, expected) => {
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'number', value });
    const handle = attachFormatDialog({ host, store, strings: en });
    try {
      handle.open();
      expect(document.querySelector('.fc-fmtdlg__preview-cell')?.textContent).toBe(expected);
    } finally {
      handle.detach();
    }
  });

  it('uses the active blank, text, boolean, and error values in the preview', () => {
    const cases: [CellValue, string][] = [
      [{ kind: 'blank' }, ''],
      [{ kind: 'text', value: 'sample' }, 'sample'],
      [{ kind: 'bool', value: true }, 'TRUE'],
      [{ kind: 'error', code: 2, text: '#VALUE!' }, '#VALUE!'],
    ];

    for (const [value, expected] of cases) {
      mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, value);
      const handle = attachFormatDialog({ host, store });
      handle.open();
      expect(document.querySelector<HTMLElement>('.fc-fmtdlg__preview-cell')?.textContent).toBe(
        expected,
      );
      handle.detach();
    }
  });

  it('uses a cached formula result before consulting the workbook', () => {
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'number', value: 7.25 }, '=A1*2');
    const getValue = vi.fn(() => ({ kind: 'number' as const, value: 99 }));
    const wb = { getValue } as unknown as WorkbookHandle;
    const handle = attachFormatDialog({ host, store, getWb: () => wb });
    handle.open();

    expect(document.querySelector<HTMLElement>('.fc-fmtdlg__preview-cell')?.textContent).toBe(
      '7.25',
    );
    expect(getValue).not.toHaveBeenCalled();
    handle.detach();
  });

  it('falls back to the workbook value when the active cell is not cached', () => {
    const getValue = vi.fn(() => ({ kind: 'number' as const, value: 42.75 }));
    const wb = { getValue } as unknown as WorkbookHandle;
    const handle = attachFormatDialog({ host, store, getWb: () => wb });
    handle.open();

    expect(document.querySelector<HTMLElement>('.fc-fmtdlg__preview-cell')?.textContent).toBe(
      '42.75',
    );
    expect(getValue).toHaveBeenCalledWith({ sheet: 0, row: 0, col: 0 });
    handle.detach();
  });

  it('formats the active numeric value as decimals and a selected currency symbol', () => {
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'number', value: 1234.5 });
    const handle = attachFormatDialog({ host, store });
    handle.open();

    document.querySelector<HTMLButtonElement>('button[data-fc-cat="fixed"]')?.click();
    const decimalsInput = document.querySelector<HTMLInputElement>(
      'input[type="number"][min="0"][max="10"]',
    );
    const thousands = document.querySelector<HTMLInputElement>('input[data-fc-check="thousands"]');
    if (!decimalsInput || !thousands) throw new Error('number controls missing');
    decimalsInput.value = '3';
    decimalsInput.dispatchEvent(new Event('input', { bubbles: true }));
    expect(document.querySelector<HTMLElement>('.fc-fmtdlg__preview-cell')?.textContent).toBe(
      '1234.500',
    );

    thousands.checked = true;
    thousands.dispatchEvent(new Event('change', { bubbles: true }));
    expect(document.querySelector<HTMLElement>('.fc-fmtdlg__preview-cell')?.textContent).toBe(
      '1,234.500',
    );

    document.querySelector<HTMLButtonElement>('button[data-fc-cat="currency"]')?.click();
    const symbol = document.querySelector<HTMLSelectElement>('select[aria-label="記号"]');
    if (!symbol) throw new Error('currency symbol select missing');
    symbol.value = '€';
    symbol.dispatchEvent(new Event('change', { bubbles: true }));
    expect(document.querySelector<HTMLElement>('.fc-fmtdlg__preview-cell')?.textContent).toBe(
      '€1,234.500',
    );
    handle.detach();
  });

  it('formats negative candidates with the draft locale, grouping, decimals, and red minus', () => {
    const handle = attachFormatDialog({ host, store, getLocale: () => 'de-DE' });
    handle.open();
    document.querySelector<HTMLButtonElement>('button[data-fc-cat="fixed"]')?.click();
    const thousands = document.querySelector<HTMLInputElement>('input[data-fc-check="thousands"]');
    if (!thousands) throw new Error('thousands checkbox missing');
    thousands.checked = true;
    thousands.dispatchEvent(new Event('change', { bubbles: true }));

    expect(
      document.querySelector<HTMLButtonElement>('button[data-fc-negative-style="minus"]')
        ?.textContent,
    ).toBe('-1.234,00');
    expect(
      document.querySelector<HTMLButtonElement>('button[data-fc-negative-style="red"]')
        ?.textContent,
    ).toBe('-1.234,00');
    expect(
      document.querySelector<HTMLButtonElement>('button[data-fc-negative-style="red-parens"]')
        ?.textContent,
    ).toBe('(1.234,00)');
    handle.detach();
  });

  it('resets preview red coloring when switching a negative style back to minus', () => {
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'number', value: -1234.5 });
    const handle = attachFormatDialog({ host, store });
    handle.open();
    document.querySelector<HTMLButtonElement>('button[data-fc-cat="fixed"]')?.click();
    const thousands = document.querySelector<HTMLInputElement>('input[data-fc-check="thousands"]');
    if (!thousands) throw new Error('thousands checkbox missing');
    thousands.checked = true;
    thousands.dispatchEvent(new Event('change', { bubbles: true }));

    const preview = document.querySelector<HTMLElement>('.fc-fmtdlg__preview-cell');
    const redParens = document.querySelector<HTMLButtonElement>(
      'button[data-fc-negative-style="red-parens"]',
    );
    const minus = document.querySelector<HTMLButtonElement>(
      'button[data-fc-negative-style="minus"]',
    );
    if (!preview || !redParens || !minus) throw new Error('negative preview controls missing');
    redParens.click();
    expect(preview.textContent).toBe('(1,234.50)');
    expect(preview.style.color).toBe('#c00000');

    minus.click();
    expect(preview.textContent).toBe('-1,234.50');
    expect(preview.style.color).toBe('');
    handle.detach();
  });

  it('preview reflects current draft and includes formatted number', () => {
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'number', value: 12345 });
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const preview = document.querySelector<HTMLElement>('.fc-fmtdlg__preview') as HTMLElement;

    const fixedBtn = document.querySelector<HTMLButtonElement>('button[data-fc-cat="fixed"]');
    fixedBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(preview.textContent).toMatch(/12,?345\.00/);

    const bold = document.querySelector<HTMLInputElement>(
      'input[data-fc-check="bold"]',
    ) as HTMLInputElement;
    const italic = document.querySelector<HTMLInputElement>(
      'input[data-fc-check="italic"]',
    ) as HTMLInputElement;
    const underline = document.querySelector<HTMLSelectElement>(
      'select[data-fc-input="underline"]',
    ) as HTMLSelectElement;
    const strike = document.querySelector<HTMLInputElement>(
      'input[data-fc-check="strike"]',
    ) as HTMLInputElement;
    bold.checked = true;
    bold.dispatchEvent(new Event('change', { bubbles: true }));
    expect(preview.style.fontWeight).toBe('bold');

    italic.checked = true;
    italic.dispatchEvent(new Event('change', { bubbles: true }));
    expect(preview.style.fontStyle).toBe('italic');

    underline.value = 'double';
    underline.dispatchEvent(new Event('change', { bubbles: true }));
    strike.checked = true;
    strike.dispatchEvent(new Event('change', { bubbles: true }));
    expect(preview.style.textDecoration).toMatch(/underline/);
    expect(preview.style.textDecoration).toMatch(/line-through/);

    const right = document.querySelector<HTMLInputElement>(
      'input[type="radio"][value="right"]',
    ) as HTMLInputElement;
    right.checked = true;
    right.dispatchEvent(new Event('change', { bubbles: true }));
    expect(preview.style.textAlign).toBe('right');

    handle.detach();
  });
});
