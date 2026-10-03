import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { attachFormatDialog } from '../../../../src/interact/format-dialog.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../../src/store/store.js';
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

  it('picks a font color from the palette flyout and closes it', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open('font');

    const toggle = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__color-toggle');
    const flyout = document.querySelector<HTMLElement>('.fc-fmtdlg__color-flyout');
    const colorInput = document.querySelector<HTMLInputElement>('input[data-fc-color="font"]');
    if (!toggle || !flyout || !colorInput) throw new Error('font color control missing');

    // The palette is out of flow until opened — the tab has no room for it inline.
    expect(flyout.hidden).toBe(true);
    expect(toggle.getAttribute('aria-expanded')).toBe('false');

    toggle.click();
    expect(flyout.hidden).toBe(false);
    expect(toggle.getAttribute('aria-expanded')).toBe('true');

    const swatch = flyout.querySelector<HTMLButtonElement>('[data-color]');
    const picked = swatch?.dataset.color;
    swatch?.click();
    expect(colorInput.value).toBe(picked);
    expect(flyout.hidden).toBe(true);
    expect(document.activeElement).toBe(toggle);

    handle.detach();
  });

  it('picks a line color from the border palette flyout and closes it', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open('border');

    const panel = document.querySelector<HTMLElement>(
      '[data-fc-tab="border"].fc-fmtdlg__panel-tab',
    );
    const toggle = panel?.querySelector<HTMLButtonElement>('.fc-fmtdlg__color-toggle');
    const flyout = panel?.querySelector<HTMLElement>('.fc-fmtdlg__color-flyout');
    const colorInput = panel?.querySelector<HTMLInputElement>('input[data-fc-color="border"]');
    if (!toggle || !flyout || !colorInput) throw new Error('border color control missing');

    expect(flyout.hidden).toBe(true);
    toggle.click();
    expect(flyout.hidden).toBe(false);

    const swatch = flyout.querySelector<HTMLButtonElement>('[data-color]');
    const picked = swatch?.dataset.color;
    swatch?.click();
    expect(colorInput.value).toBe(picked);
    expect(flyout.hidden).toBe(true);
    expect(document.activeElement).toBe(toggle);

    handle.detach();
  });

  it('lets Escape close the palette flyout without closing the dialog', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open('font');

    const overlay = document.querySelector<HTMLElement>('.fc-fmtdlg');
    const toggle = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__color-toggle');
    const flyout = document.querySelector<HTMLElement>('.fc-fmtdlg__color-flyout');
    if (!overlay || !toggle || !flyout) throw new Error('font color control missing');

    toggle.click();
    overlay.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(flyout.hidden).toBe(true);
    expect(overlay.hidden).toBe(false);

    overlay.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(overlay.hidden).toBe(true);

    handle.detach();
  });

  it('closes the palette flyout when another tab is opened', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open('font');

    const toggle = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__color-toggle');
    const flyout = document.querySelector<HTMLElement>('.fc-fmtdlg__color-flyout');
    if (!toggle || !flyout) throw new Error('font color control missing');

    toggle.click();
    expect(flyout.hidden).toBe(false);
    document.querySelector<HTMLButtonElement>('button[data-fc-tab="border"]')?.click();
    expect(flyout.hidden).toBe(true);

    handle.detach();
  });
});
