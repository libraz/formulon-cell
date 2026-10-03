import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { sparklineAt } from '../../../src/commands/sparkline.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings/en.js';
import { ja } from '../../../src/i18n/strings/ja.js';
import { attachMacSparklineDialog } from '../../../src/interact/mac-sparkline-dialog.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const makeWorkbook = (): WorkbookHandle =>
  ({
    sheetCount: 1,
    sheetName: () => 'Sheet1',
  }) as unknown as WorkbookHandle;

describe('attachMacSparklineDialog', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;

  beforeEach(() => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('validates the one-cell destination and commits the selected type', () => {
    const handle = attachMacSparklineDialog({
      host,
      store,
      getWb: makeWorkbook,
    });
    handle.open();

    const source = host.ownerDocument.querySelector<HTMLInputElement>('.fc-macsparkdlg__source');
    const destination = host.ownerDocument.querySelector<HTMLInputElement>(
      '.fc-macsparkdlg__destination',
    );
    const type = host.ownerDocument.querySelector<HTMLSelectElement>('.fc-macsparkdlg__type');
    const ok = host.ownerDocument.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    expect(source?.value).toBe('A1:C1');
    expect(destination?.value).toBe('D1');
    expect(ok?.disabled).toBe(false);
    expect(type).not.toBeNull();
    expect(ok).not.toBeNull();
    if (!type || !ok) throw new Error('sparkline chooser controls missing');
    type.value = 'column';
    type.dispatchEvent(new Event('change', { bubbles: true }));
    ok.click();

    expect(sparklineAt(store.getState(), { sheet: 0, row: 0, col: 3 })).toEqual({
      kind: 'column',
      source: 'A1:C1',
      showNegative: true,
    });
    handle.detach();
  });

  it('blocks a multi-cell destination before mutating the store', () => {
    const handle = attachMacSparklineDialog({
      host,
      store,
      getWb: makeWorkbook,
      strings: en,
    });
    handle.open();
    const destination = host.ownerDocument.querySelector<HTMLInputElement>(
      '.fc-macsparkdlg__destination',
    );
    const ok = host.ownerDocument.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    expect(destination).not.toBeNull();
    expect(ok).not.toBeNull();
    if (!destination || !ok) throw new Error('sparkline chooser controls missing');
    destination.value = 'D1:E1';
    destination.dispatchEvent(new Event('input', { bubbles: true }));
    expect(ok.disabled).toBe(true);
    expect(document.querySelector('[data-mac-sparkline-error]')?.textContent).toContain('one cell');
    ok.click();
    expect(store.getState().sparkline.sparklines.size).toBe(0);
    handle.detach();
  });

  it('keeps the strings set last across refresh and keeps the chosen type', () => {
    const handle = attachMacSparklineDialog({ host, store, getWb: makeWorkbook, strings: ja });
    handle.open();
    const type = document.querySelector<HTMLSelectElement>('.fc-macsparkdlg__type');
    expect(type).not.toBeNull();
    if (!type) throw new Error('type select missing');
    type.value = 'column';
    handle.setStrings(en);
    handle.refresh();
    const overlay = document.querySelector<HTMLElement>('.fc-macsparkdlg');
    expect(overlay?.querySelector('.fc-fmtdlg__header')?.textContent).toBe(en.macSparkline.title);
    expect(overlay?.getAttribute('aria-label')).toBe(en.macSparkline.title);
    expect(overlay?.querySelector('.fc-fmtdlg__btn--primary')?.textContent).toBe(
      en.macSparkline.ok,
    );
    expect(type.value).toBe('column');
    handle.detach();
  });

  it('applies the shared dialog class to the overlay without leaking it onto the panel', () => {
    const handle = attachMacSparklineDialog({ host, store, getWb: makeWorkbook });
    const overlay = document.querySelector<HTMLElement>('.fc-macsparkdlg');
    expect(overlay?.classList.contains('fc-fmtdlg')).toBe(true);
    expect(
      overlay?.querySelector('.fc-macsparkdlg__panel')?.classList.contains('fc-macsparkdlg'),
    ).toBe(false);
    handle.detach();
  });
});
