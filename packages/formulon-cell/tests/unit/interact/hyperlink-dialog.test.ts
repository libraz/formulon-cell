import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import { attachHyperlinkDialog } from '../../../src/interact/hyperlink-dialog.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('attachHyperlinkDialog', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('mounts hidden; open() reveals the dialog and focuses the URL input', async () => {
    const store = createSpreadsheetStore();
    const handle = attachHyperlinkDialog({ host, store });
    const overlay = host.querySelector<HTMLElement>('.fc-hldlg');
    expect(overlay?.hidden).toBe(true);

    handle.open();
    expect(overlay?.hidden).toBe(false);
    handle.detach();
  });

  it('OK with a non-empty URL writes hyperlink onto the active cell', () => {
    const store = createSpreadsheetStore();
    const handle = attachHyperlinkDialog({ host, store });
    handle.open();
    const input = host.querySelector<HTMLInputElement>('.fc-hldlg input');
    if (!input) throw new Error('expected url input');
    input.value = 'https://example.com';
    const okBtn = host.querySelectorAll<HTMLButtonElement>('.fc-hldlg button')[2];
    if (!okBtn) throw new Error('expected ok button');
    okBtn.click();

    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.hyperlink).toBe('https://example.com');
    handle.detach();
  });

  it('OK on an empty URL surfaces the error and does not write', () => {
    const store = createSpreadsheetStore();
    const handle = attachHyperlinkDialog({ host, store });
    handle.open();
    const okBtn = host.querySelectorAll<HTMLButtonElement>('.fc-hldlg button')[2];
    if (!okBtn) throw new Error('expected ok button');
    okBtn.click();

    expect(host.querySelector<HTMLElement>('.fc-hldlg__error')?.hidden).toBe(false);
    expect(store.getState().format.formats.size).toBe(0);
    handle.detach();
  });

  it('Remove button clears the hyperlink', () => {
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      format: {
        formats: new Map([[addrKey({ sheet: 0, row: 0, col: 0 }), { hyperlink: 'https://old' }]]),
      },
    }));
    const handle = attachHyperlinkDialog({ host, store });
    handle.open();
    const removeBtn = host.querySelectorAll<HTMLButtonElement>('.fc-hldlg button')[0];
    if (!removeBtn) throw new Error('expected remove button');
    removeBtn.click();

    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.hyperlink).toBeUndefined();
    handle.detach();
  });

  it('Remove button is hidden when the active cell has no hyperlink yet', () => {
    const store = createSpreadsheetStore();
    const handle = attachHyperlinkDialog({ host, store });
    handle.open();
    const removeBtn = host.querySelectorAll<HTMLButtonElement>('.fc-hldlg button')[0];
    if (!removeBtn) throw new Error('expected remove button');
    expect(removeBtn.hidden).toBe(true);
    handle.detach();
  });

  it('Escape closes the overlay', () => {
    const store = createSpreadsheetStore();
    const handle = attachHyperlinkDialog({ host, store });
    handle.open();
    const overlay = host.querySelector<HTMLElement>('.fc-hldlg');
    if (!overlay) throw new Error('expected overlay');
    overlay.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(overlay.hidden).toBe(true);
    handle.detach();
  });
});
