import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachMacSparklineDialog } from '../../../src/interact/mac-sparkline-dialog.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const workbook = {
  sheetCount: 2,
  sheetName: (i: number) => ['Sheet1', 'Sheet2'][i],
  sheetIndexByName: (name: string) => ['Sheet1', 'Sheet2'].indexOf(name),
} as unknown as WorkbookHandle;

describe('Edit Sparklines dialog', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  const field = (selector: string): HTMLInputElement => {
    const found = document.querySelector<HTMLInputElement>(selector);
    if (!found) throw new Error(`missing ${selector}`);
    return found;
  };

  it('refuses a location on another sheet than the group, and honours canCommit', () => {
    const handle = attachMacSparklineDialog({
      host,
      store: createSpreadsheetStore(),
      getWb: () => workbook,
      canCommit: () => true,
    });
    handle.open({ groupId: 'g', sheet: 0, data: 'A1:C2', location: 'E1:E2' });
    const ok = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    expect(ok?.disabled).toBe(false);

    const location = field('.fc-macsparkdlg__destination');
    location.value = 'Sheet2!E1:E2';
    location.dispatchEvent(new Event('input', { bubbles: true }));
    expect(ok?.disabled).toBe(true);
    expect(document.querySelector('[data-mac-sparkline-error]')?.textContent).not.toBe('');

    location.value = 'Sheet1!E1:E2';
    location.dispatchEvent(new Event('input', { bubbles: true }));
    expect(ok?.disabled).toBe(false);
    handle.detach();
  });

  it('refuses a location over rows and columns with an alert, keeping the dialog open', () => {
    // Excel for Mac: OK on a 2 x 2 location raises strman_0008 key 44 and leaves the dialog up.
    const handle = attachMacSparklineDialog({
      host,
      store: createSpreadsheetStore(),
      getWb: () => workbook,
      canCommit: () => true,
    });
    handle.open({ groupId: 'g', sheet: 0, data: 'A1:B4', location: 'E1:F2' });
    const ok = document.querySelector<HTMLButtonElement>(
      '.fc-macsparkdlg .fc-fmtdlg__btn--primary',
    );
    expect(ok?.disabled).toBe(false);
    ok?.click();
    const alert = document.querySelector('.fc-tb__dlg--alert');
    expect(alert?.querySelector('.fc-tb__dlg__message')?.textContent).toBe(
      'Location reference is not valid because the cells are not all in the same column or row. Select cells that are all in a single column or row.',
    );
    expect(document.querySelector<HTMLElement>('.fc-macsparkdlg')?.hidden).toBe(false);
    handle.detach();
  });

  it('applies canCommit to an edit', () => {
    const handle = attachMacSparklineDialog({
      host,
      store: createSpreadsheetStore(),
      getWb: () => workbook,
      canCommit: () => false,
    });
    handle.open({ groupId: 'g', sheet: 0, data: 'A1:C2', location: 'E1:E2' });
    expect(document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.disabled).toBe(
      true,
    );
    handle.detach();
  });

  it('refuses a source over the size cap', () => {
    const handle = attachMacSparklineDialog({
      host,
      store: createSpreadsheetStore(),
      getWb: () => workbook,
    });
    handle.open();
    const source = field('.fc-macsparkdlg__source');
    const location = field('.fc-macsparkdlg__destination');
    location.value = 'Z1';
    source.value = 'A1:Z1048576';
    source.dispatchEvent(new Event('input', { bubbles: true }));
    expect(document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.disabled).toBe(
      true,
    );
    handle.detach();
  });
});
