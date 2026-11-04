import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { EngineCapabilities } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { ja } from '../../../src/i18n/strings/ja.js';
import { attachSheetTabsController } from '../../../src/mount/sheet-tabs-controller.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

const mountMac = (toggleSheetProtection?: () => void, hiddenSheets: number[] = []) => {
  const host = document.createElement('div');
  const sheetTabs = document.createElement('div');
  const sheetMenu = document.createElement('div');
  sheetMenu.hidden = true;
  host.append(sheetTabs, sheetMenu);
  document.body.appendChild(host);
  const sheets = ['A', 'B', 'C', 'D'];
  const wb = {
    get sheetCount() {
      return sheets.length;
    },
    sheetName: (idx: number) => sheets[idx] ?? '',
    setSheetTabHidden: () => true,
    setSheetVisibility: () => true,
    spreadsheetProfileId: () => 'mac-ja_JP',
    capabilities: {
      sheetMutate: true,
      sheetTabHidden: true,
      sheetVisibility: true,
    } as EngineCapabilities,
  } as unknown as WorkbookHandle;
  const store = createSpreadsheetStore();
  const controller = attachSheetTabsController({
    addSheetBtn: document.createElement('button'),
    firstSheet: document.createElement('button'),
    getStrings: () => ja,
    getWb: () => wb,
    history: null as never,
    hydrateActiveSheet: vi.fn(),
    invalidate: vi.fn(),
    lastSheet: document.createElement('button'),
    refreshStatusBar: vi.fn(),
    sheetMenu,
    sheetTabs,
    store,
    toggleSheetProtection,
  });
  store.setState((st) => ({
    ...st,
    layout: { ...st.layout, hiddenSheets: new Set(hiddenSheets) },
  }));
  controller.update();
  return { controller, host, sheetMenu, sheetTabs, store };
};

describe('Mac sheet tab menu', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('lists the entries in Excel order', () => {
    const { controller, sheetMenu, sheetTabs } = mountMac(() => undefined);
    const tab = sheetTabs.querySelector<HTMLButtonElement>('.fc-host__sheetbar-tab');
    if (!tab) throw new Error('missing tab');
    controller.showMenu(0, tab, 10, 10);
    const labels = Array.from(
      sheetMenu.querySelectorAll<HTMLElement>('.fc-sheetmenu__item'),
      (el) => el.textContent,
    );
    expect(labels).toEqual([
      'シートの挿入 (⇧F11)',
      '削除',
      '名前の変更',
      '移動またはコピー...',
      'コードの表示',
      'シートの保護...',
      'シート見出しの色',
      '表示しない',
      '再表示...',
      'すべてのシートを選択',
    ]);
  });

  it('runs the protection flow from the protect entry and greys what the host lacks', () => {
    const toggle = vi.fn();
    const { controller, sheetMenu, sheetTabs } = mountMac(toggle);
    const tab = sheetTabs.querySelector<HTMLButtonElement>('.fc-host__sheetbar-tab');
    if (!tab) throw new Error('missing tab');
    controller.showMenu(0, tab, 10, 10);
    const items = Array.from(sheetMenu.querySelectorAll<HTMLButtonElement>('.fc-sheetmenu__item'));
    const byText = (text: string) => items.find((item) => item.textContent === text);
    byText('シートの保護...')?.click();
    expect(toggle).toHaveBeenCalledOnce();
    expect(byText('コードの表示')?.disabled).toBe(true);
    expect(byText('すべてのシートを選択')?.disabled).toBe(true);
    expect(byText('再表示...')?.disabled).toBe(true);
  });

  const openMenu = (m: ReturnType<typeof mountMac>) => {
    const tab = m.sheetTabs.querySelector<HTMLButtonElement>('.fc-host__sheetbar-tab');
    if (!tab) throw new Error('missing tab');
    m.controller.showMenu(0, tab, 10, 10);
  };
  const item = (m: ReturnType<typeof mountMac>, text: string) => {
    const found = Array.from(
      m.sheetMenu.querySelectorAll<HTMLButtonElement>('.fc-sheetmenu__item'),
    ).find((el) => el.textContent === text);
    if (!found) throw new Error(`missing ${text}`);
    return found;
  };

  it('opens the tab colour as a submenu with the theme and standard palettes', () => {
    const m = mountMac(() => undefined);
    openMenu(m);
    expect(m.sheetMenu.querySelector('.fc-sheetmenu__swatch')).toBeNull();
    const trigger = item(m, 'シート見出しの色');
    expect(trigger.getAttribute('aria-haspopup')).toBe('menu');
    expect(m.sheetMenu.querySelector('.fc-sheetmenu__sub')).toBeNull();
    trigger.click();
    const sub = m.sheetMenu.querySelector<HTMLElement>('.fc-sheetmenu__sub');
    if (!sub) throw new Error('missing submenu');
    expect(trigger.getAttribute('aria-expanded')).toBe('true');
    expect(sub.querySelector('.fc-colorpalette__contrast')?.textContent).toBe(
      'ハイ コントラストのみ',
    );
    expect(sub.querySelector('.fc-colorpalette__action--no-fill')?.textContent).toBe('色なし');
    expect(
      Array.from(sub.querySelectorAll('.fc-colorpalette__heading'), (h) => h.textContent),
    ).toEqual(['テーマの色', '標準の色']);
    expect(
      sub.querySelectorAll('.fc-colorpalette__grid--theme .fc-colorpalette__swatch'),
    ).toHaveLength(60);
    expect(
      sub.querySelectorAll('.fc-colorpalette__grid--standard .fc-colorpalette__swatch'),
    ).toHaveLength(10);
    expect(sub.querySelector('.fc-colorpalette__action--more')?.textContent).toBe('その他の色...');
  });

  it('colours the tab from a swatch and clears it with the no-colour button', () => {
    const m = mountMac(() => undefined);
    openMenu(m);
    item(m, 'シート見出しの色').click();
    m.sheetMenu
      .querySelector<HTMLButtonElement>('.fc-colorpalette__grid--standard .fc-colorpalette__swatch')
      ?.click();
    expect(m.store.getState().layout.sheetTabColors.get(0)).toBe('#c00000');
    expect(m.sheetMenu.hidden).toBe(true);
    openMenu(m);
    item(m, 'シート見出しの色').click();
    m.sheetMenu.querySelector<HTMLButtonElement>('.fc-colorpalette__action--no-fill')?.click();
    expect(m.store.getState().layout.sheetTabColors.has(0)).toBe(false);
  });

  it('makes the submenu 198pt wide', () => {
    const css = readFileSync(join(root, 'src/styles/core/app/overlays/sheet-menu.css'), 'utf8');
    expect(css).toMatch(/\.fc-sheetmenu__sub\s*\{[\s\S]*?width: 198px;/);
  });

  it('opens one list dialog of the hidden sheets from the enabled unhide entry', async () => {
    const m = mountMac(() => undefined, [1, 3]);
    openMenu(m);
    const labels = Array.from(
      m.sheetMenu.querySelectorAll('.fc-sheetmenu__item'),
      (el) => el.textContent,
    );
    expect(labels.filter((l) => l?.includes('再表示'))).toEqual(['再表示...']);
    const unhide = item(m, '再表示...');
    expect(unhide.disabled).toBe(false);
    unhide.click();
    const options = document.querySelectorAll<HTMLElement>('.fc-movecopy__list [role="option"]');
    expect(Array.from(options, (o) => o.textContent)).toEqual(['B', 'D']);
    options[1]?.click();
    document
      .querySelector<HTMLButtonElement>('.fc-fmtdlg__footer .fc-fmtdlg__btn--primary')
      ?.click();
    await new Promise((r) => setTimeout(r, 0));
    const hiddenNow = m.store.getState().layout.hiddenSheets;
    expect(hiddenNow.has(3)).toBe(false);
    expect(hiddenNow.has(1)).toBe(true);
    expect(m.store.getState().data.sheetIndex).toBe(3);
  });
});
