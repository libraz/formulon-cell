import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { addrKey } from '../../../../src/engine/workbook-handle.js';
import { en } from '../../../../src/i18n/strings/en.js';
import { attachFormatDialog } from '../../../../src/interact/format-dialog.js';
import {
  type CellFormat,
  createSpreadsheetStore,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import { trackConnectedListenerLeaks } from '../connected-listener-leaks.js';
import { flushRaf, setActive } from './fixtures.js';

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

  it('mounts a hidden overlay on attach', () => {
    const handle = attachFormatDialog({ host, store });
    const overlay = document.querySelector<HTMLElement>('.fc-fmtdlg');
    expect(overlay).not.toBeNull();
    expect(overlay?.hidden).toBe(true);
    handle.detach();
  });

  it('open() reveals the overlay and focuses the active tab button', async () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();
    const overlay = document.querySelector<HTMLElement>('.fc-fmtdlg');
    expect(overlay?.hidden).toBe(false);

    await flushRaf();
    const numberTab = document.querySelector<HTMLButtonElement>('button[data-fc-tab="number"]');
    expect(document.activeElement).toBe(numberTab);
    handle.detach();
  });

  it('open(tab) opens the requested tab and focuses it', async () => {
    const handle = attachFormatDialog({ host, store });
    handle.open('more');
    const moreTab = document.querySelector<HTMLButtonElement>('button[data-fc-tab="more"]');
    const morePanel = document.querySelector<HTMLDivElement>('div[data-fc-tab="more"]');
    expect(moreTab?.getAttribute('aria-selected')).toBe('true');
    expect(morePanel?.hidden).toBe(false);

    await flushRaf();
    expect(document.activeElement).toBe(moreTab);
    handle.detach();
  });

  it('opens the shared validation editor as a dedicated Data Validation dialog', async () => {
    const handle = attachFormatDialog({ host, store });
    handle.open('more', { mode: 'dataValidation', focus: 'validation' });

    const overlay = document.querySelector<HTMLElement>('.fc-fmtdlg');
    const title = document.querySelector<HTMLElement>('.fc-fmtdlg__title');
    const tabs = document.querySelector<HTMLElement>('.fc-fmtdlg__tabs');
    const preview = document.querySelector<HTMLElement>('.fc-fmtdlg__preview');
    const sections = Array.from(
      document.querySelectorAll<HTMLElement>(
        '.fc-fmtdlg__panel-tab[data-fc-tab="more"] > .fc-fmtdlg__section',
      ),
    );
    const validationKind = document.querySelector<HTMLSelectElement>(
      '.fc-fmtdlg__panel-tab[data-fc-tab="more"] select[aria-label="種類"]',
    );

    expect(overlay?.classList.contains('fc-fmtdlg--data-validation')).toBe(true);
    expect(title?.textContent).toBe('入力規則');
    expect(tabs?.hidden).toBe(true);
    expect(preview?.hidden).toBe(true);
    expect(sections.map((section) => section.hidden)).toEqual([true, true, false]);
    expect(sections[2]?.classList.contains('fc-fmtdlg__section--standalone')).toBe(true);

    await flushRaf();
    expect(document.activeElement).toBe(validationKind);

    handle.detach();
  });

  it('returns number formats and borders through dxf mode without mutating the active cell', () => {
    const handle = attachFormatDialog({ host, store, strings: en });
    let applied: Partial<CellFormat> | null = null;
    handle.open('number', {
      mode: 'dxf',
      onApplyDxf: (format) => {
        applied = format;
      },
    });
    expect(document.querySelector<HTMLElement>('.fc-fmtdlg__preview-cell')?.textContent).toBe(
      '12,345',
    );
    document.querySelector<HTMLButtonElement>('button[data-fc-cat="fixed"]')?.click();
    document.querySelector<HTMLButtonElement>('button[data-fc-tab="border"]')?.click();
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__border-preset--outline')?.click();
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

    expect(applied).toMatchObject({
      numFmt: { kind: 'fixed', decimals: 2 },
      borders: {
        top: { style: 'thin' },
        right: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
      },
    });
    expect(store.getState().format.formats).toHaveLength(0);
    handle.detach();
  });

  it('restores the normal Format Cells chrome after Data Validation mode', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open('more', { mode: 'dataValidation', focus: 'validation' });
    handle.close();
    handle.open('font');

    const overlay = document.querySelector<HTMLElement>('.fc-fmtdlg');
    expect(overlay?.classList.contains('fc-fmtdlg--data-validation')).toBe(false);
    expect(document.querySelector<HTMLElement>('.fc-fmtdlg__title')?.textContent).toBe(
      'セルの書式設定',
    );
    expect(document.querySelector<HTMLElement>('.fc-fmtdlg__tabs')?.hidden).toBe(false);
    expect(document.querySelector<HTMLElement>('.fc-fmtdlg__preview')?.hidden).toBe(false);
    expect(
      Array.from(
        document.querySelectorAll<HTMLElement>(
          '.fc-fmtdlg__panel-tab[data-fc-tab="more"] > .fc-fmtdlg__section',
        ),
      ).map((section) => section.hidden),
    ).toEqual([false, false, false]);

    handle.detach();
  });

  it('close() hides the overlay and refocuses host', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();
    handle.close();
    const overlay = document.querySelector<HTMLElement>('.fc-fmtdlg');
    expect(overlay?.hidden).toBe(true);
    expect(document.activeElement).toBe(host);
    handle.detach();
  });

  it('header close button hides the overlay without applying changes', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const fontTab = document.querySelector<HTMLButtonElement>('button[data-fc-tab="font"]');
    fontTab?.click();
    const boldInput = document.querySelector<HTMLInputElement>('input[data-fc-check="bold"]');
    if (!boldInput) throw new Error('bold input missing');
    boldInput.checked = true;
    boldInput.dispatchEvent(new Event('change', { bubbles: true }));

    const close = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__close');
    expect(close?.textContent).toBe('');
    close?.click();

    expect(document.querySelector<HTMLElement>('.fc-fmtdlg')?.hidden).toBe(true);
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))).toBe(
      undefined,
    );
    handle.detach();
  });

  it('detach() removes the overlay from DOM', () => {
    const handle = attachFormatDialog({ host, store });
    expect(document.querySelector('.fc-fmtdlg')).not.toBeNull();
    handle.detach();
    expect(document.querySelector('.fc-fmtdlg')).toBeNull();
  });

  it('switching tabs toggles aria-selected and panel visibility', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const fontTab = document.querySelector<HTMLButtonElement>('button[data-fc-tab="font"]');
    fontTab?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(fontTab?.getAttribute('aria-selected')).toBe('true');
    expect(fontTab?.tabIndex).toBe(0);
    expect(fontTab?.getAttribute('aria-controls')).toBe('fc-fmtdlg-panel-font');
    const numberTab = document.querySelector<HTMLButtonElement>('button[data-fc-tab="number"]');
    expect(numberTab?.getAttribute('aria-selected')).toBe('false');
    expect(numberTab?.tabIndex).toBe(-1);

    const fontPanel = document.querySelector<HTMLDivElement>('div[data-fc-tab="font"]');
    const numberPanel = document.querySelector<HTMLDivElement>('div[data-fc-tab="number"]');
    expect(fontPanel?.hidden).toBe(false);
    expect(fontPanel?.getAttribute('aria-labelledby')).toBe('fc-fmtdlg-tab-font');
    expect(numberPanel?.hidden).toBe(true);

    handle.detach();
  });

  it('format tabs support Excel-style arrow, Home, and End keyboard navigation', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const numberTab = document.querySelector<HTMLButtonElement>('button[data-fc-tab="number"]');
    const alignTab = document.querySelector<HTMLButtonElement>('button[data-fc-tab="align"]');
    const moreTab = document.querySelector<HTMLButtonElement>('button[data-fc-tab="more"]');
    numberTab?.focus();
    numberTab?.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }));
    expect(alignTab?.getAttribute('aria-selected')).toBe('true');
    expect(document.activeElement).toBe(alignTab);

    alignTab?.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true }));
    expect(moreTab?.getAttribute('aria-selected')).toBe('true');
    expect(document.activeElement).toBe(moreTab);

    moreTab?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', bubbles: true }));
    expect(numberTab?.getAttribute('aria-selected')).toBe('true');
    expect(document.activeElement).toBe(numberTab);

    handle.detach();
  });

  it('labels Format Cells controls for keyboard and assistive navigation', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const namedControls = Array.from(
      document.querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>(
        '.fc-fmtdlg input[aria-label], .fc-fmtdlg select[aria-label], .fc-fmtdlg textarea[aria-label]',
      ),
    );
    expect(namedControls.length).toBeGreaterThan(20);

    const colorInputs = Array.from(
      document.querySelectorAll<HTMLInputElement>('.fc-fmtdlg input[type="color"]'),
    );
    expect(colorInputs.length).toBeGreaterThan(0);
    for (const input of colorInputs) expect(input.getAttribute('aria-label')).toBeTruthy();

    const textareas = Array.from(
      document.querySelectorAll<HTMLTextAreaElement>('.fc-fmtdlg textarea'),
    );
    expect(textareas.length).toBeGreaterThan(0);
    for (const area of textareas) expect(area.getAttribute('aria-label')).toBeTruthy();

    const radioGroups = Array.from(
      document.querySelectorAll<HTMLElement>('.fc-fmtdlg__choice-grid[role="radiogroup"]'),
    );
    expect(radioGroups.length).toBeGreaterThanOrEqual(2);
    for (const group of radioGroups) expect(group.getAttribute('aria-label')).toBeTruthy();
    handle.detach();
  });

  it('describes the More tab as a secondary entry point', () => {
    const handle = attachFormatDialog({ host, store, strings: en });
    handle.open('more');
    expect(document.querySelector<HTMLElement>('.fc-fmtdlg__more-hint')?.textContent).toContain(
      'secondary editor',
    );
    handle.detach();
  });

  it('clicking tab strip outside button does nothing', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();
    const tabsStrip = document.querySelector<HTMLElement>('.fc-fmtdlg__tabs');
    tabsStrip?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    const numberTab = document.querySelector<HTMLButtonElement>('button[data-fc-tab="number"]');
    expect(numberTab?.getAttribute('aria-selected')).toBe('true');
    handle.detach();
  });

  it('Cancel button closes the dialog without writing', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const buttons = Array.from(document.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__btn'));
    const cancelBtn = buttons.find((b) => b.textContent === 'キャンセル') as HTMLButtonElement;
    cancelBtn.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const overlay = document.querySelector<HTMLElement>('.fc-fmtdlg');
    expect(overlay?.hidden).toBe(true);
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 })),
    ).toBeUndefined();
    handle.detach();
  });

  it('backdrop click closes the dialog; panel click does not', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const overlay = document.querySelector<HTMLElement>('.fc-fmtdlg') as HTMLElement;
    const panel = document.querySelector<HTMLElement>('.fc-fmtdlg__panel') as HTMLElement;

    panel.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(overlay.hidden).toBe(false);

    overlay.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(overlay.hidden).toBe(true);

    handle.detach();
  });

  it('Escape closes the dialog and stops propagation', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const overlay = document.querySelector<HTMLElement>('.fc-fmtdlg') as HTMLElement;
    const hostKey = vi.fn();
    host.addEventListener('keydown', hostKey);

    const ev = new KeyboardEvent('keydown', { key: 'Escape', bubbles: true });
    overlay.dispatchEvent(ev);

    expect(overlay.hidden).toBe(true);
    expect(hostKey).not.toHaveBeenCalled();
    handle.detach();
  });

  it('Enter on input applies and closes', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const fixedBtn = document.querySelector<HTMLButtonElement>('button[data-fc-cat="fixed"]');
    fixedBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const overlay = document.querySelector<HTMLElement>('.fc-fmtdlg') as HTMLElement;
    const decimalsInput = document.querySelector<HTMLInputElement>(
      'input[type="number"][min="0"][max="10"]',
    ) as HTMLInputElement;
    decimalsInput.value = '5';
    decimalsInput.dispatchEvent(new Event('input', { bubbles: true }));

    const ev = new KeyboardEvent('keydown', { key: 'Enter', bubbles: true });
    Object.defineProperty(ev, 'target', { value: decimalsInput });
    overlay.dispatchEvent(ev);

    expect(overlay.hidden).toBe(true);
    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.numFmt).toEqual({ kind: 'fixed', decimals: 5 });

    handle.detach();
  });

  it('Enter on a button does not apply (lets button click handle it)', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const overlay = document.querySelector<HTMLElement>('.fc-fmtdlg') as HTMLElement;
    const someBtn = document.querySelector<HTMLButtonElement>(
      'button[data-fc-tab="font"]',
    ) as HTMLButtonElement;

    const ev = new KeyboardEvent('keydown', { key: 'Enter', bubbles: true });
    Object.defineProperty(ev, 'target', { value: someBtn });
    overlay.dispatchEvent(ev);

    expect(overlay.hidden).toBe(false);
    handle.detach();
  });

  it('detach() removes wired listeners (clicking after detach is inert)', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();
    handle.detach();
    expect(document.querySelector('.fc-fmtdlg')).toBeNull();
  });

  it('detach() removes every listener left on still-connected targets', () => {
    const { registered, leaked } = trackConnectedListenerLeaks(() => {
      const handle = attachFormatDialog({ host, store });
      handle.open('font');
      handle.detach();
    });
    expect(registered).toBeGreaterThan(50);
    expect(leaked).toBe(0);
  });
});
