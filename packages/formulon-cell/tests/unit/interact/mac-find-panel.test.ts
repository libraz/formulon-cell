import { afterEach, describe, expect, it } from 'vitest';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { defaultStrings, ja } from '../../../src/i18n/strings.js';
import { attachFindReplace } from '../../../src/interact/find-replace.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

afterEach(() => {
  document.body.replaceChildren();
});

const mount = async (platform: 'mac' | 'default', texts: [number, number, string][] = []) => {
  const fcHost = document.createElement('div');
  fcHost.className = 'fc-host';
  fcHost.dataset.fcPlatform = platform;
  const host = document.createElement('div');
  host.tabIndex = -1;
  fcHost.appendChild(host);
  document.body.appendChild(fcHost);
  const store = createSpreadsheetStore();
  const wb = await WorkbookHandle.createDefault({ preferStub: true });
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    for (const [row, col, value] of texts) {
      const addr = { sheet: 0, row, col };
      wb.setText(addr, value);
      cells.set(addrKey(addr), { value: { kind: 'text', value }, formula: null });
    }
    return { ...s, data: { ...s.data, cells } };
  });
  const handle = attachFindReplace({ host, store, wb, strings: ja });
  const panel = () => host.querySelector<HTMLElement>('.fc-findpanel');
  const input = () => {
    const el = host.querySelector<HTMLInputElement>('.fc-findpanel__input');
    if (!el) throw new Error('no panel input');
    return el;
  };
  const dialog = () => host.querySelector<HTMLElement>('.fc-find');
  const active = () => store.getState().selection.active;
  const type = (text: string): void => {
    input().value = text;
    input().dispatchEvent(new Event('input', { bubbles: true }));
  };
  const enter = (shiftKey = false): void => {
    input().dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', shiftKey, bubbles: true }));
  };
  const click = (selector: string): void => {
    host.querySelector<HTMLElement>(selector)?.click();
  };
  const menuLabels = (): (string | null)[] =>
    [...document.body.querySelectorAll('.fc-popup-menu__item')].map((b) => b.textContent);
  const pick = (label: string): void => {
    [...document.body.querySelectorAll<HTMLElement>('.fc-popup-menu__item')]
      .find((b) => b.textContent === label)
      ?.click();
  };
  return { host, handle, panel, input, dialog, active, type, enter, click, menuLabels, pick };
};

const SHEET = [
  [0, 0, 'apple'],
  [2, 1, 'apple pie'],
  [4, 0, 'banana'],
] as [number, number, string][];

describe('the Mac find panel', () => {
  // Excel for Mac: ⌘F opens a compact 検索する panel, not the Find and Replace dialog.
  it('opens for Find on the Mac, with the sheet scope as placeholder', async () => {
    const m = await mount('mac');
    m.handle.open();
    expect(m.panel()?.hidden).toBe(false);
    expect(m.dialog()?.hidden).toBe(true);
    expect(m.panel()?.querySelector('.fc-findpanel__title')?.textContent).toBe('検索する');
    expect(m.input().placeholder).toBe('シート内を検索する');
    expect(document.activeElement).toBe(m.input());
  });

  it('keeps the dialog for Find off the Mac', async () => {
    const m = await mount('default');
    m.handle.open();
    expect(m.panel()).toBeNull();
    expect(m.dialog()?.hidden).toBe(false);
  });

  it('finds the next match on Return and Shift+Return, wrapping, and stays open', async () => {
    const m = await mount('mac', SHEET);
    m.handle.open();
    m.type('a');
    m.enter();
    expect(m.active()).toEqual({ sheet: 0, row: 2, col: 1 });
    // The text stays selected for the next query.
    expect(document.activeElement).toBe(m.input());
    expect([m.input().selectionStart, m.input().selectionEnd]).toEqual([0, 1]);
    // Shift+Return also moves forward.
    m.enter(true);
    expect(m.active()).toEqual({ sheet: 0, row: 4, col: 0 });
    m.enter();
    expect(m.active()).toEqual({ sheet: 0, row: 0, col: 0 });
    expect(m.panel()?.hidden).toBe(false);
  });

  it('steps back with the previous arrow and forward with the next arrow', async () => {
    const m = await mount('mac', SHEET);
    m.handle.open();
    m.type('a');
    m.click('.fc-findpanel__prev');
    expect(m.active()).toEqual({ sheet: 0, row: 4, col: 0 });
    m.click('.fc-findpanel__prev');
    expect(m.active()).toEqual({ sheet: 0, row: 2, col: 1 });
    m.click('.fc-findpanel__next');
    expect(m.active()).toEqual({ sheet: 0, row: 4, col: 0 });
  });

  it('shows the not-found alert when nothing matches', async () => {
    const m = await mount('mac', SHEET);
    m.handle.open();
    m.type('zzz');
    m.enter();
    expect(
      document.body.querySelector('[role="alertdialog"] .fc-tb__dlg__headline')?.textContent,
    ).toBe('検索しているデータが見つかりません。');
  });

  it('switches the scope from the magnifier menu, shared with the dialog', async () => {
    const m = await mount('mac');
    m.handle.open();
    m.click('.fc-findpanel__scope');
    expect(menuState()).toEqual([
      ['シート内を検索する', 'true'],
      ['ブック内を検索する', 'false'],
    ]);
    m.pick('ブック内を検索する');
    expect(m.input().placeholder).toBe('ブック内を検索する');
    m.handle.open('find');
    expect(m.host.querySelector<HTMLSelectElement>('#fc-find-within')?.value).toBe('workbook');
  });

  it('leads to the dialog from More Options, carrying the text', async () => {
    const m = await mount('mac');
    m.handle.open();
    m.type('pie');
    m.click('.fc-findpanel__more');
    expect(m.menuLabels()).toEqual(['高度な検索...', '置換...']);
    m.pick('高度な検索...');
    expect(m.panel()?.hidden).toBe(true);
    expect(m.dialog()?.hidden).toBe(false);
    expect(m.dialog()?.dataset.fcFindTab).toBe('find');
    expect(m.host.querySelector<HTMLInputElement>('#fc-find-find-input')?.value).toBe('pie');

    m.handle.close();
    m.handle.open();
    m.click('.fc-findpanel__more');
    m.pick('置換...');
    expect(m.dialog()?.dataset.fcFindTab).toBe('replace');
  });

  it('clears the field with the clear button', async () => {
    const m = await mount('mac');
    m.handle.open();
    const clear = () => m.host.querySelector<HTMLButtonElement>('.fc-findpanel__clear');
    expect(clear()?.hidden).toBe(true);
    m.type('abc');
    expect(clear()?.hidden).toBe(false);
    clear()?.click();
    expect(m.input().value).toBe('');
    expect(clear()?.hidden).toBe(true);
  });

  it('closes on Escape while it has focus, handing focus to the grid', async () => {
    const m = await mount('mac');
    m.handle.open();
    m.input().dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(m.panel()?.hidden).toBe(true);
    expect(document.activeElement).toBe(m.host);
  });
});

const menuState = (): [string | null, string | null][] =>
  [...document.body.querySelectorAll('.fc-popup-menu__item')].map((b) => [
    b.textContent,
    b.getAttribute('aria-checked'),
  ]);

describe('the Mac Find and Replace dialog', () => {
  const dialogFind = (m: Awaited<ReturnType<typeof mount>>, query: string): void => {
    const field = m.host.querySelector<HTMLInputElement>('#fc-find-find-input');
    if (!field) throw new Error('no field');
    field.value = query;
    field.dispatchEvent(new Event('input', { bubbles: true }));
    field.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
  };

  it('drops down this session’s searches, most recent first, shared with the panel', async () => {
    const m = await mount('mac', SHEET);
    m.handle.open('find');
    const history = () => m.host.querySelector<HTMLButtonElement>('.fc-find__history');
    expect(history()?.disabled).toBe(true);
    dialogFind(m, 'apple');
    dialogFind(m, 'banana');
    m.handle.close();
    m.handle.open();
    m.type('pie');
    m.enter();
    m.handle.close();
    m.handle.open('find');
    dialogFind(m, 'apple');
    expect(history()?.disabled).toBe(false);
    history()?.click();
    expect(m.menuLabels()).toEqual(['apple', 'pie', 'banana']);
    m.pick('banana');
    expect(m.host.querySelector<HTMLInputElement>('#fc-find-find-input')?.value).toBe('banana');
  });

  it('styles Previous and Next as primary on the Mac only', async () => {
    const primary = (m: Awaited<ReturnType<typeof mount>>) =>
      ['.fc-find__btn--prev', '.fc-find__btn--next'].map((s) =>
        m.host.querySelector(s)?.classList.contains('fc-fmtdlg__btn--primary'),
      );
    const mac = await mount('mac');
    expect(primary(mac)).toEqual([true, true]);
    document.body.replaceChildren();
    const win = await mount('default');
    win.handle.setStrings(defaultStrings);
    expect(primary(win)).toEqual([false, false]);
    expect(win.host.querySelector('.fc-find__history')).toBeNull();
  });
});
