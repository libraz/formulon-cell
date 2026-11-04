import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { defaultStrings } from '../../../src/i18n/strings.js';
import { attachPageSetupDialog } from '../../../src/interact/page-setup-dialog.js';
import { createSpreadsheetStore, getPageSetup, mutators } from '../../../src/store/store.js';

const t = defaultStrings.pageSetup;
/** Visible within its tab panel; the panel's own tab state is ignored. */
const visible = (el: Element | null): boolean => {
  for (let node = el; node && node.getAttribute('role') !== 'tabpanel'; node = node.parentElement) {
    if ((node as HTMLElement).hidden) return false;
  }
  return !!el;
};
const byLabel = (label: string): HTMLElement | null =>
  document.querySelector<HTMLElement>(`.fc-pgsetup [aria-label="${label}"]`);

describe('Page Setup dialog on the Mac profile', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    host.dataset.fcPlatform = 'mac';
    document.body.appendChild(host);
  });

  afterEach(() => {
    while (document.body.firstChild) document.body.removeChild(document.body.firstChild);
  });

  const open = () => {
    const store = createSpreadsheetStore();
    const handle = attachPageSetupDialog({ host, store });
    handle.open();
    return { store, handle };
  };

  it('offers orientation as two radios and drops paper size and print quality', () => {
    const { handle } = open();
    const radios = Array.from(
      document.querySelectorAll<HTMLInputElement>(
        '.fc-pgsetup input[name^="fc-pgsetup-orientation"]',
      ),
    );
    expect(radios.map((r) => r.value)).toEqual(['portrait', 'landscape']);
    expect(radios.map((r) => r.checked)).toEqual([true, false]);
    expect(visible(byLabel(t.paperSize))).toBe(false);
    expect(visible(byLabel(t.printQuality))).toBe(false);
    handle.detach();
  });

  it('keeps the paper size and orientation edits flowing into the setup', () => {
    const { store, handle } = open();
    const landscape = document.querySelector<HTMLInputElement>(
      '.fc-pgsetup input[name^="fc-pgsetup-orientation"][value="landscape"]',
    );
    if (!landscape) throw new Error('missing landscape radio');
    landscape.checked = true;
    landscape.dispatchEvent(new Event('change', { bubbles: true }));
    document.querySelector<HTMLButtonElement>('.fc-pgsetup .fc-fmtdlg__btn--primary')?.click();
    expect(getPageSetup(store.getState(), 0).orientation).toBe('landscape');
    expect(getPageSetup(store.getState(), 0).paperSize).toBe('A4');
    handle.detach();
  });

  it('drops the printer minimum margins row from the Margins tab', () => {
    const { handle } = open();
    expect(visible(document.querySelector('.fc-pgsetup__printable'))).toBe(false);
    handle.detach();
  });

  it('shows the header and footer slot editors only for a custom preset', () => {
    const { store, handle } = open();
    const slot = (): HTMLElement | null =>
      document.querySelector<HTMLElement>(`.fc-pgsetup input[aria-label^="${t.headerLabel} "]`);
    expect(visible(slot())).toBe(false);
    handle.close();
    mutators.setPageSetup(store, 0, { headerLeft: 'Left' });
    handle.open();
    expect(visible(slot())).toBe(true);
    handle.detach();
  });

  it('lays out the Page tab like the desktop app', () => {
    const store = createSpreadsheetStore();
    const handle = attachPageSetupDialog({
      host,
      store,
      getPrinterProfiles: () => [{ id: 'demo', name: 'Demo' }],
      refreshPrinterProfiles: async () => [],
    } as unknown as Parameters<typeof attachPageSetupDialog>[0]);
    handle.open();
    const panel = document.querySelector<HTMLElement>('.fc-pgsetup [role="tabpanel"]');
    expect(visible(panel?.querySelector('[data-pgsetup-printer]') ?? null)).toBe(false);
    const text = panel?.textContent ?? '';
    for (const label of [
      '拡大/縮小:',
      '次のページ数に合わせる:',
      'ページ (横) ×',
      'ページ (縦)',
      '先頭ページ番号:',
    ]) {
      expect(text).toContain(label);
    }
    expect(text).not.toContain('% 標準サイズ');
    expect(panel?.querySelectorAll('.fc-pgsetup__orient-icon')).toHaveLength(2);
    const fitWidth = byLabel(t.fitWidth) as HTMLInputElement;
    const scale = byLabel(t.scale) as HTMLInputElement;
    expect(fitWidth.disabled).toBe(true);
    expect(scale.disabled).toBe(false);
    const fit = document.querySelector<HTMLInputElement>('.fc-pgsetup input[value="fit"]');
    fit?.click();
    expect(fitWidth.disabled).toBe(false);
    expect(scale.disabled).toBe(true);
    expect((byLabel('先頭ページ番号:') as HTMLInputElement).value).toBe('自動');
    handle.detach();
  });
});
