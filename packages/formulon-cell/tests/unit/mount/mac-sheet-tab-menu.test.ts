import { afterEach, describe, expect, it, vi } from 'vitest';
import type { EngineCapabilities } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { ja } from '../../../src/i18n/strings/ja.js';
import { attachSheetTabsController } from '../../../src/mount/sheet-tabs-controller.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const mountMac = (toggleSheetProtection?: () => void) => {
  const host = document.createElement('div');
  const sheetTabs = document.createElement('div');
  const sheetMenu = document.createElement('div');
  sheetMenu.hidden = true;
  host.append(sheetTabs, sheetMenu);
  document.body.appendChild(host);
  const sheets = ['A', 'B'];
  const wb = {
    get sheetCount() {
      return sheets.length;
    },
    sheetName: (idx: number) => sheets[idx] ?? '',
    spreadsheetProfileId: () => 'mac-ja_JP',
    capabilities: {
      sheetMutate: true,
      sheetTabHidden: true,
      sheetVisibility: true,
    } as EngineCapabilities,
  } as unknown as WorkbookHandle;
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
    store: createSpreadsheetStore(),
    toggleSheetProtection,
  });
  controller.update();
  return { controller, host, sheetMenu, sheetTabs };
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
      sheetMenu.querySelectorAll<HTMLElement>('.fc-sheetmenu__item, .fc-sheetmenu__colors-label'),
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
      ...Array.from(
        sheetMenu.querySelectorAll<HTMLElement>('.fc-sheetmenu__swatches .fc-sheetmenu__item'),
        (el) => el.textContent,
      ),
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
});
