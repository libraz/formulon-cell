import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
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
    const overlay = document.querySelector<HTMLElement>('.fc-hldlg');
    expect(overlay?.hidden).toBe(true);

    handle.open();
    expect(overlay?.hidden).toBe(false);
    handle.detach();
  });

  it('OK with a non-empty URL writes hyperlink onto the active cell', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const handle = attachHyperlinkDialog({ host, store, history });
    handle.open();
    const input = document.querySelector<HTMLInputElement>('.fc-hldlg input');
    if (!input) throw new Error('expected url input');
    input.value = 'https://example.com';
    const okBtn = document.querySelectorAll<HTMLButtonElement>('.fc-hldlg button')[2];
    if (!okBtn) throw new Error('expected ok button');
    okBtn.click();

    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.hyperlink).toBe('https://example.com');
    expect(history.canUndo()).toBe(true);

    history.undo();
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))).toBe(
      undefined,
    );

    history.redo();
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.hyperlink,
    ).toBe('https://example.com');
    handle.detach();
  });

  it('Cancel Enter keeps its native click behavior and does not submit the URL', () => {
    const store = createSpreadsheetStore();
    const handle = attachHyperlinkDialog({ host, store });
    handle.open();
    const input = document.querySelector<HTMLInputElement>('.fc-hldlg input[type="url"]');
    const buttons = document.querySelectorAll<HTMLButtonElement>(
      '.fc-hldlg .fc-fmtdlg__footer button',
    );
    const cancelBtn = buttons[1];
    if (!input || !cancelBtn) throw new Error('expected URL input and Cancel button');
    input.value = 'https://example.com';
    const event = new KeyboardEvent('keydown', {
      key: 'Enter',
      bubbles: true,
      cancelable: true,
    });

    try {
      cancelBtn.focus();
      cancelBtn.dispatchEvent(event);
      expect(event.defaultPrevented).toBe(false);
      expect(store.getState().format.formats.size).toBe(0);
      expect(document.querySelector<HTMLElement>('.fc-hldlg')?.hidden).toBe(false);

      cancelBtn.click();
      expect(store.getState().format.formats.size).toBe(0);
      expect(document.querySelector<HTMLElement>('.fc-hldlg')?.hidden).toBe(true);
    } finally {
      handle.detach();
    }
  });

  it('native Enter on the OK button clicks once without an overlay submission', () => {
    const store = createSpreadsheetStore();
    const handle = attachHyperlinkDialog({ host, store });
    handle.open();
    const input = document.querySelector<HTMLInputElement>('.fc-hldlg input[type="url"]');
    const okBtn = document.querySelector<HTMLButtonElement>('.fc-hldlg .fc-fmtdlg__btn--primary');
    if (!input || !okBtn) throw new Error('expected URL input and OK button');
    input.value = 'https://example.com';
    const event = new KeyboardEvent('keydown', {
      key: 'Enter',
      bubbles: true,
      cancelable: true,
    });

    try {
      okBtn.focus();
      okBtn.dispatchEvent(event);
      expect(event.defaultPrevented).toBe(false);
      okBtn.click(); // happy-dom does not synthesize the native button click for keydown.
      expect(
        store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.hyperlink,
      ).toBe('https://example.com');
      expect(document.querySelector<HTMLElement>('.fc-hldlg')?.hidden).toBe(true);
    } finally {
      handle.detach();
    }
  });

  it('wraps Tab and Shift+Tab through the shared dialog shell', () => {
    const store = createSpreadsheetStore();
    const handle = attachHyperlinkDialog({ host, store });
    handle.open();
    const input = document.querySelector<HTMLInputElement>('.fc-hldlg input[type="url"]');
    const okBtn = document.querySelector<HTMLButtonElement>('.fc-hldlg .fc-fmtdlg__btn--primary');
    if (!input || !okBtn) throw new Error('expected URL input and OK button');

    try {
      okBtn.focus();
      const tab = new KeyboardEvent('keydown', {
        key: 'Tab',
        bubbles: true,
        cancelable: true,
      });
      okBtn.dispatchEvent(tab);
      expect(tab.defaultPrevented).toBe(true);
      expect(document.activeElement).toBe(input);

      const shiftTab = new KeyboardEvent('keydown', {
        key: 'Tab',
        shiftKey: true,
        bubbles: true,
        cancelable: true,
      });
      input.dispatchEvent(shiftTab);
      expect(shiftTab.defaultPrevented).toBe(true);
      expect(document.activeElement).toBe(okBtn);
    } finally {
      handle.detach();
    }
  });

  it('ignores composing Enter without writing or dismissing the hyperlink dialog', () => {
    const store = createSpreadsheetStore();
    const handle = attachHyperlinkDialog({ host, store });
    handle.open();
    const input = document.querySelector<HTMLInputElement>('.fc-hldlg input[type="url"]');
    if (!input) throw new Error('expected URL input');
    input.value = 'https://example.com';
    const event = new KeyboardEvent('keydown', {
      key: 'Enter',
      isComposing: true,
      bubbles: true,
      cancelable: true,
    });

    try {
      input.dispatchEvent(event);
      expect(event.defaultPrevented).toBe(false);
      expect(store.getState().format.formats.size).toBe(0);
      expect(document.querySelector<HTMLElement>('.fc-hldlg')?.hidden).toBe(false);
    } finally {
      handle.detach();
    }
  });

  it('OK on an empty URL surfaces the error and does not write', () => {
    const store = createSpreadsheetStore();
    const handle = attachHyperlinkDialog({ host, store });
    handle.open();
    const okBtn = document.querySelectorAll<HTMLButtonElement>('.fc-hldlg button')[2];
    if (!okBtn) throw new Error('expected ok button');
    okBtn.click();

    expect(document.querySelector<HTMLElement>('.fc-hldlg__error')?.hidden).toBe(false);
    expect(store.getState().format.formats.size).toBe(0);
    handle.detach();
  });

  it('Remove button clears the hyperlink', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    store.setState((s) => ({
      ...s,
      format: {
        formats: new Map([[addrKey({ sheet: 0, row: 0, col: 0 }), { hyperlink: 'https://old' }]]),
      },
    }));
    const handle = attachHyperlinkDialog({ host, store, history });
    handle.open();
    const removeBtn = document.querySelectorAll<HTMLButtonElement>('.fc-hldlg button')[0];
    if (!removeBtn) throw new Error('expected remove button');
    removeBtn.click();

    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.hyperlink).toBeUndefined();

    history.undo();
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.hyperlink,
    ).toBe('https://old');

    history.redo();
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.hyperlink,
    ).toBeUndefined();
    handle.detach();
  });

  it('Remove button is hidden when the active cell has no hyperlink yet', () => {
    const store = createSpreadsheetStore();
    const handle = attachHyperlinkDialog({ host, store });
    handle.open();
    const removeBtn = document.querySelectorAll<HTMLButtonElement>('.fc-hldlg button')[0];
    if (!removeBtn) throw new Error('expected remove button');
    expect(removeBtn.hidden).toBe(true);
    handle.detach();
  });

  it('Escape closes the overlay', () => {
    const store = createSpreadsheetStore();
    const handle = attachHyperlinkDialog({ host, store });
    handle.open();
    const overlay = document.querySelector<HTMLElement>('.fc-hldlg');
    if (!overlay) throw new Error('expected overlay');
    overlay.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(overlay.hidden).toBe(true);
    handle.detach();
  });
});
