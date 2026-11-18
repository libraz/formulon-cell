import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../../src/commands/history.js';
import { addrKey, WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { en } from '../../../../src/i18n/strings.js';
import { attachFormatDialog } from '../../../../src/interact/format-dialog.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../../src/store/store.js';
import { setActive } from './fixtures.js';

describe('data validation dialog with an engine that refuses the rule', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    setActive(store, 0, 0);
    wb = await WorkbookHandle.createDefault();
  });

  afterEach(() => {
    wb.dispose();
    document.body.innerHTML = '';
  });

  const applyCustom = (formula: string, history: History): void => {
    const handle = attachFormatDialog({
      getLocale: () => 'en-US',
      host,
      store,
      strings: en,
      history,
      getWb: () => wb,
    });
    handle.open('more', { mode: 'dataValidation', focus: 'validation' });
    const kind = Array.from(document.querySelectorAll<HTMLSelectElement>('select')).find((s) =>
      Array.from(s.options).some((o) => o.value === 'textLength'),
    );
    const input = document
      .querySelector('[data-range-picker="format-validation-formula"]')
      ?.closest('.fc-range-picker')
      ?.querySelector('input');
    if (!kind || !input) throw new Error('validation controls missing');
    kind.value = 'custom';
    kind.dispatchEvent(new Event('change', { bubbles: true }));
    input.value = formula;
    input.dispatchEvent(new Event('input', { bubbles: true }));
    Array.from(document.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__footer button'))
      .find((b) => b.textContent === 'OK')
      ?.click();
  };

  it('keeps a refused rule out of the store, history, and engine and keeps the dialog open', async () => {
    const history = new History();
    applyCustom('=SUM()', history);
    await Promise.resolve();

    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))).toBe(
      undefined,
    );
    expect(history.canUndo()).toBe(false);
    expect(wb.getValidationsForSheet(0)).toHaveLength(0);
    expect(document.querySelector<HTMLElement>('.fc-fmtdlg')?.hidden).toBe(false);
    const alert = document.querySelector<HTMLElement>('.fc-fmtdlg [role="alert"]');
    expect(alert?.hidden).toBe(false);
    expect(alert?.textContent).toBe(en.commandRefusal.engineFailed);
  });

  it('takes an accepted rule and closes', async () => {
    const history = new History();
    applyCustom('=A1>0', history);
    await Promise.resolve();

    expect(wb.getValidationsForSheet(0)).toHaveLength(1);
    expect(history.canUndo()).toBe(true);
    expect(document.querySelector<HTMLElement>('.fc-fmtdlg')?.hidden).toBe(true);
  });
});
