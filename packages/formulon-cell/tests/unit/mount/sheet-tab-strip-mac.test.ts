import { afterEach, describe, expect, it, vi } from 'vitest';
import type { EngineCapabilities } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { ja } from '../../../src/i18n/strings/ja.js';
import { attachSheetTabsController } from '../../../src/mount/sheet-tabs-controller.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const mount = (names: string[]) => {
  const host = document.createElement('div');
  const sheetTabs = document.createElement('div');
  const sheetMenu = document.createElement('div');
  const firstSheet = document.createElement('button');
  const lastSheet = document.createElement('button');
  const addSheetBtn = document.createElement('button');
  host.append(firstSheet, lastSheet, sheetTabs, addSheetBtn, sheetMenu);
  document.body.appendChild(host);
  const sheets = [...names];
  const wb = {
    get sheetCount() {
      return sheets.length;
    },
    sheetName: (i: number) => sheets[i] ?? '',
    spreadsheetProfileId: () => 'mac-ja_JP',
    getTables: () => [],
    capabilities: { sheetMutate: true, sheetTabHidden: true } as EngineCapabilities,
    addSheet(name?: string) {
      sheets.push(name ?? `Sheet${sheets.length + 1}`);
      return sheets.length - 1;
    },
    moveSheet(from: number, to: number) {
      const [s] = sheets.splice(from, 1);
      if (s === undefined) return false;
      sheets.splice(to, 0, s);
      return true;
    },
  } as unknown as WorkbookHandle;
  const store = createSpreadsheetStore();
  const controller = attachSheetTabsController({
    addSheetBtn,
    firstSheet,
    getStrings: () => ja,
    getWb: () => wb,
    history: null as never,
    hydrateActiveSheet: vi.fn(),
    invalidate: vi.fn(),
    lastSheet,
    refreshStatusBar: vi.fn(),
    sheetMenu,
    sheetTabs,
    store,
  });
  controller.update();
  return { controller, sheets, store, sheetTabs, firstSheet, lastSheet, addSheetBtn };
};

/** Lay the tabs out 79 pt apart inside a 160 pt strip. */
const layout = (sheetTabs: HTMLElement, state: { scroll: number }, count: number) => {
  Object.defineProperty(sheetTabs, 'clientWidth', { value: 160, configurable: true });
  Object.defineProperty(sheetTabs, 'scrollWidth', { value: count * 79, configurable: true });
  Object.defineProperty(sheetTabs, 'scrollLeft', {
    get: () => state.scroll,
    set: (v: number) => {
      state.scroll = Math.max(0, Math.min(v, count * 79 - 160));
    },
    configurable: true,
  });
  sheetTabs.getBoundingClientRect = () => new DOMRect(0, 0, 160, 24);
  sheetTabs.querySelectorAll<HTMLElement>('.fc-host__sheetbar-tab').forEach((tab, i) => {
    tab.getBoundingClientRect = () => new DOMRect(i * 79 - state.scroll, 0, 79, 24);
  });
};

describe('Mac sheet tab strip', () => {
  afterEach(() => document.body.replaceChildren());

  it('inserts the + button sheet after the active sheet and activates it', () => {
    const h = mount(['Sheet1', 'Sheet2', 'Sheet3']);
    h.addSheetBtn.click();
    expect(h.sheets).toEqual(['Sheet1', 'Sheet4', 'Sheet2', 'Sheet3']);
    expect(h.store.getState().data.sheetIndex).toBe(1);
  });

  it('scrolls the strip by one tab from the arrows instead of switching sheets', () => {
    const names = Array.from({ length: 6 }, (_, i) => `Sheet${i + 1}`);
    const h = mount(names);
    const state = { scroll: 0 };
    layout(h.sheetTabs, state, names.length);
    h.sheetTabs.dispatchEvent(new Event('scroll'));
    expect(h.firstSheet.disabled).toBe(true);
    expect(h.lastSheet.disabled).toBe(false);

    h.lastSheet.click();
    expect(state.scroll).toBe(79);
    layout(h.sheetTabs, state, names.length);
    h.lastSheet.click();
    expect(state.scroll).toBe(158);
    h.firstSheet.click();
    expect(state.scroll).toBe(79);
    expect(h.store.getState().data.sheetIndex).toBe(0);
    expect(h.firstSheet.disabled).toBe(false);
  });

  it('reorders by dragging a tab to the left half of another and activates it', () => {
    const h = mount(['Sheet1', 'Sheet2', 'Sheet3']);
    layout(h.sheetTabs, { scroll: 0 }, 3);
    const tabs = h.sheetTabs.querySelectorAll<HTMLElement>('.fc-host__sheetbar-tab');
    const ptr = (type: string, x: number, target: EventTarget = document) =>
      target.dispatchEvent(new MouseEvent(type, { clientX: x, button: 0, bubbles: true }));
    ptr('pointerdown', 190, tabs[2]);
    ptr('pointermove', 100);
    ptr('pointermove', 20);
    const mark = document.querySelector<HTMLElement>('.fc-host__sheetbar-drop');
    expect(mark?.style.left).toBe('0px');
    ptr('pointerup', 20);
    expect(h.sheets).toEqual(['Sheet3', 'Sheet1', 'Sheet2']);
    expect(h.store.getState().data.sheetIndex).toBe(0);
    expect(document.querySelector('.fc-host__sheetbar-drop')).toBeNull();
  });
});
