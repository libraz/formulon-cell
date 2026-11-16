import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { addrKey } from '../../../../src/engine/workbook-handle.js';
import { attachFormatDialog } from '../../../../src/interact/format-dialog.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import { setRange } from './fixtures.js';

// Desktop Excel: with a mixed selection, applying the dialog changes only the
// properties the user set; every cell keeps the rest of its own format.
// Sizes are stored in CSS px; the dialog shows points.
const PX_20 = (20 * 4) / 3;

describe('Format Cells apply on a mixed selection', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  const at = (row: number) => ({ sheet: 0, row, col: 0 });
  const fmt = (row: number) => store.getState().format.formats.get(addrKey(at(row)));

  beforeEach(() => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    mutators.setCellFormat(store, at(0), {
      bold: true,
      fontSize: 14,
      numFmt: { kind: 'fixed', decimals: 2 },
    });
    mutators.setCellFormat(store, at(1), {
      italic: true,
      fontSize: 10,
      numFmt: { kind: 'percent', decimals: 0 },
    });
    mutators.setCellFormat(store, at(2), {
      numFmt: { kind: 'fixed', decimals: 2 },
      fill: '#ffff00',
    });
    setRange(store, 0, 0, 2, 0);
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  const ok = (): void => {
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
  };

  it('setting only the font size leaves every other property alone', () => {
    const handle = attachFormatDialog({ getLocale: () => 'en-US', host, store });
    handle.open('font');
    document.querySelector<HTMLButtonElement>('[data-fc-font-size="20"]')?.click();
    ok();
    expect(fmt(0)).toMatchObject({
      bold: true,
      numFmt: { kind: 'fixed', decimals: 2 },
    });
    expect(fmt(1)).toMatchObject({
      italic: true,
      numFmt: { kind: 'percent', decimals: 0 },
    });
    expect(fmt(2)).toMatchObject({
      fill: '#ffff00',
      numFmt: { kind: 'fixed', decimals: 2 },
    });
    expect(fmt(2)?.bold).toBeFalsy();
    for (const row of [0, 1, 2]) expect(fmt(row)?.fontSize).toBeCloseTo(PX_20, 6);
    handle.detach?.();
  });

  it('setting only the number format leaves font and fill alone', () => {
    const handle = attachFormatDialog({ getLocale: () => 'en-US', host, store });
    handle.open('number');
    document.querySelector<HTMLButtonElement>('[data-fc-cat="fixed"]')?.click();
    const decimals = document.querySelector<HTMLInputElement>(
      '.fc-fmtdlg__panel-tab[data-fc-tab="number"] input[type="number"][min="0"][max="10"]',
    );
    if (!decimals) throw new Error('decimals field missing');
    decimals.value = '1';
    decimals.dispatchEvent(new Event('input', { bubbles: true }));
    decimals.dispatchEvent(new Event('change', { bubbles: true }));
    ok();
    for (const row of [0, 1, 2]) {
      expect(fmt(row)?.numFmt, `row ${row}`).toMatchObject({ kind: 'fixed', decimals: 1 });
    }
    expect(fmt(0)).toMatchObject({ bold: true, fontSize: 14 });
    expect(fmt(1)).toMatchObject({ italic: true, fontSize: 10 });
    expect(fmt(2)).toMatchObject({ fill: '#ffff00' });
    handle.detach?.();
  });

  it('leaves the number format alone when only the font changes', () => {
    const handle = attachFormatDialog({ getLocale: () => 'en-US', host, store });
    handle.open('font');
    document.querySelector<HTMLButtonElement>('[data-fc-font-size="20"]')?.click();
    ok();
    expect(fmt(1)?.numFmt).toEqual({ kind: 'percent', decimals: 0 });
    handle.detach?.();
  });
});
