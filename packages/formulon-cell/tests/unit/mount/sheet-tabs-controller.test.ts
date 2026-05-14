import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import type { EngineCapabilities } from '../../../src/engine/capabilities.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import en from '../../../src/i18n/en.js';
import type { Strings } from '../../../src/i18n/strings.js';
import { attachSheetTabsController } from '../../../src/mount/sheet-tabs-controller.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

interface FakeWbState {
  sheets: string[];
  capabilities: Partial<EngineCapabilities>;
  failRename?: boolean;
}

const makeFakeWb = (state: FakeWbState): WorkbookHandle => {
  return {
    get sheetCount() {
      return state.sheets.length;
    },
    sheetName(idx: number): string {
      return state.sheets[idx] ?? '';
    },
    capabilities: {
      sheetMutate: true,
      sheetTabHidden: true,
      ...state.capabilities,
    } as EngineCapabilities,
    addSheet(name?: string): number {
      state.sheets.push(name ?? `Sheet${state.sheets.length + 1}`);
      return state.sheets.length - 1;
    },
    renameSheet(idx: number, name: string): boolean {
      if (state.failRename) return false;
      state.sheets[idx] = name;
      return true;
    },
    removeSheet(idx: number): boolean {
      if (state.sheets.length <= 1) return false;
      state.sheets.splice(idx, 1);
      return true;
    },
    moveSheet(from: number, to: number): boolean {
      const [s] = state.sheets.splice(from, 1);
      if (s === undefined) return false;
      state.sheets.splice(to, 0, s);
      return true;
    },
    clearViewportHint(): void {
      /* noop in tests */
    },
  } as unknown as WorkbookHandle;
};

interface Harness {
  controller: ReturnType<typeof attachSheetTabsController>;
  state: FakeWbState;
  store: SpreadsheetStore;
  sheetTabs: HTMLDivElement;
  sheetMenu: HTMLDivElement;
  firstSheet: HTMLButtonElement;
  lastSheet: HTMLButtonElement;
  addSheetBtn: HTMLButtonElement;
  hydrateActiveSheet: ReturnType<typeof vi.fn>;
  invalidate: ReturnType<typeof vi.fn>;
  refreshStatusBar: ReturnType<typeof vi.fn>;
  host: HTMLElement;
  detach: () => void;
}

const mount = (state: FakeWbState, opts?: { initialIndex?: number }): Harness => {
  const host = document.createElement('div');
  document.body.appendChild(host);

  const firstSheet = document.createElement('button');
  const lastSheet = document.createElement('button');
  const addSheetBtn = document.createElement('button');
  const sheetTabs = document.createElement('div');
  const sheetMenu = document.createElement('div');
  sheetMenu.hidden = true;
  host.append(firstSheet, lastSheet, addSheetBtn, sheetTabs, sheetMenu);

  const store = createSpreadsheetStore();
  if (opts?.initialIndex !== undefined && opts.initialIndex > 0) {
    store.setState((s) => ({
      ...s,
      data: { ...s.data, sheetIndex: opts.initialIndex ?? 0 },
    }));
  }

  const wb = makeFakeWb(state);
  const hydrateActiveSheet = vi.fn();
  const invalidate = vi.fn();
  const refreshStatusBar = vi.fn();

  const controller = attachSheetTabsController({
    addSheetBtn,
    firstSheet,
    getStrings: () => en as Strings,
    getWb: () => wb,
    history: null as never,
    host,
    hydrateActiveSheet,
    invalidate,
    lastSheet,
    refreshStatusBar,
    sheetMenu,
    sheetTabs,
    store,
  });
  controller.update();

  return {
    controller,
    state,
    store,
    sheetTabs,
    sheetMenu,
    firstSheet,
    lastSheet,
    addSheetBtn,
    hydrateActiveSheet,
    invalidate,
    refreshStatusBar,
    host,
    detach: () => {
      controller.detach();
      host.remove();
    },
  };
};

describe('mount/sheet-tabs-controller', () => {
  let h: Harness | undefined;

  afterEach(() => {
    h?.detach();
    h = undefined;
  });

  describe('update()', () => {
    it('renders one button per visible sheet with active=selected', () => {
      h = mount({ sheets: ['Alpha', 'Beta', 'Gamma'], capabilities: {} });
      const tabs = h.sheetTabs.querySelectorAll('.fc-host__sheetbar-tab');
      expect(tabs.length).toBe(3);
      expect(tabs[0]?.textContent).toBe('Alpha');
      expect(tabs[0]?.getAttribute('aria-selected')).toBe('true');
      expect((tabs[0] as HTMLButtonElement).tabIndex).toBe(0);
      expect(tabs[1]?.getAttribute('aria-selected')).toBe('false');
      expect((tabs[1] as HTMLButtonElement).tabIndex).toBe(-1);
    });

    it('skips hidden sheets from layout.hiddenSheets', () => {
      h = mount({ sheets: ['A', 'B', 'C'], capabilities: {} });
      h.store.setState((s) => ({
        ...s,
        layout: { ...s.layout, hiddenSheets: new Set([1]) },
      }));
      h.controller.update();
      const labels = Array.from(h.sheetTabs.querySelectorAll('.fc-host__sheetbar-tab')).map(
        (t) => t.textContent,
      );
      expect(labels).toEqual(['A', 'C']);
    });

    it('disables nav buttons at the boundaries', () => {
      h = mount({ sheets: ['A', 'B'], capabilities: {} });
      expect(h.firstSheet.disabled).toBe(true); // active=0 -> no prev
      expect(h.lastSheet.disabled).toBe(false);
    });
  });

  describe('switchSheet()', () => {
    it('mutates store.sheetIndex and triggers side effects', () => {
      h = mount({ sheets: ['A', 'B', 'C'], capabilities: {} });
      h.controller.switchSheet(2);
      expect(h.store.getState().data.sheetIndex).toBe(2);
      expect(h.hydrateActiveSheet).toHaveBeenCalledTimes(1);
      expect(h.invalidate).toHaveBeenCalledTimes(1);
      expect(h.refreshStatusBar).toHaveBeenCalledTimes(1);
    });

    it('is a no-op when target equals current index', () => {
      h = mount({ sheets: ['A', 'B'], capabilities: {} });
      h.controller.switchSheet(0);
      expect(h.hydrateActiveSheet).not.toHaveBeenCalled();
      expect(h.invalidate).not.toHaveBeenCalled();
    });

    it('ignores out-of-range indices', () => {
      h = mount({ sheets: ['A', 'B'], capabilities: {} });
      h.controller.switchSheet(99);
      h.controller.switchSheet(-1);
      expect(h.store.getState().data.sheetIndex).toBe(0);
      expect(h.hydrateActiveSheet).not.toHaveBeenCalled();
    });

    it('ignores hidden sheets', () => {
      h = mount({ sheets: ['A', 'B', 'C'], capabilities: {} });
      h.store.setState((s) => ({
        ...s,
        layout: { ...s.layout, hiddenSheets: new Set([1]) },
      }));
      h.controller.switchSheet(1);
      expect(h.store.getState().data.sheetIndex).toBe(0);
    });
  });

  describe('switchRelative()', () => {
    it('navigates to adjacent visible sheets', () => {
      h = mount({ sheets: ['A', 'B', 'C'], capabilities: {} });
      h.controller.switchRelative(1);
      expect(h.store.getState().data.sheetIndex).toBe(1);
      h.controller.switchRelative(1);
      expect(h.store.getState().data.sheetIndex).toBe(2);
      h.controller.switchRelative(-1);
      expect(h.store.getState().data.sheetIndex).toBe(1);
    });

    it('clamps at edges (does nothing past last)', () => {
      h = mount({ sheets: ['A', 'B'], capabilities: {} });
      h.controller.switchRelative(1);
      h.hydrateActiveSheet.mockClear();
      h.controller.switchRelative(1);
      expect(h.store.getState().data.sheetIndex).toBe(1);
      expect(h.hydrateActiveSheet).not.toHaveBeenCalled();
    });
  });

  describe('tab click + keyboard', () => {
    it('click on a tab switches the sheet', () => {
      h = mount({ sheets: ['A', 'B'], capabilities: {} });
      const tabB = h.sheetTabs.querySelectorAll<HTMLButtonElement>('.fc-host__sheetbar-tab')[1];
      tabB?.click();
      expect(h.store.getState().data.sheetIndex).toBe(1);
    });

    it('ArrowRight on active tab switches forward', () => {
      h = mount({ sheets: ['A', 'B'], capabilities: {} });
      const tabs = h.sheetTabs.querySelectorAll<HTMLButtonElement>('.fc-host__sheetbar-tab');
      tabs[0]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }));
      expect(h.store.getState().data.sheetIndex).toBe(1);
    });

    it('F2 begins rename — input replaces the tab', () => {
      h = mount({ sheets: ['A', 'B'], capabilities: {} });
      const tab = h.sheetTabs.querySelectorAll<HTMLButtonElement>('.fc-host__sheetbar-tab')[0];
      tab?.dispatchEvent(new KeyboardEvent('keydown', { key: 'F2', bubbles: true }));
      const input = h.sheetTabs.querySelector<HTMLInputElement>('input.fc-host__sheetbar-rename');
      expect(input).not.toBeNull();
      expect(input?.value).toBe('A');
    });
  });

  describe('rename flow', () => {
    it('Enter commits a rename', () => {
      h = mount({ sheets: ['A', 'B'], capabilities: {} });
      const tab = h.sheetTabs.querySelectorAll<HTMLButtonElement>('.fc-host__sheetbar-tab')[0];
      tab?.dispatchEvent(new KeyboardEvent('keydown', { key: 'F2', bubbles: true }));
      const input = h.sheetTabs.querySelector<HTMLInputElement>('input.fc-host__sheetbar-rename');
      if (!input) throw new Error('rename input not found');
      input.value = 'Renamed';
      input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
      expect(h.state.sheets[0]).toBe('Renamed');
      // Tab should be back, showing the new label.
      const tabs = h.sheetTabs.querySelectorAll<HTMLButtonElement>('.fc-host__sheetbar-tab');
      expect(tabs[0]?.textContent).toBe('Renamed');
    });

    it('Escape cancels — name unchanged', () => {
      h = mount({ sheets: ['A', 'B'], capabilities: {} });
      const tab = h.sheetTabs.querySelectorAll<HTMLButtonElement>('.fc-host__sheetbar-tab')[0];
      tab?.dispatchEvent(new KeyboardEvent('keydown', { key: 'F2', bubbles: true }));
      const input = h.sheetTabs.querySelector<HTMLInputElement>('input.fc-host__sheetbar-rename');
      if (!input) throw new Error('rename input not found');
      input.value = 'Discarded';
      input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
      expect(h.state.sheets[0]).toBe('A');
    });

    it('marks aria-invalid when engine refuses the rename', () => {
      h = mount({ sheets: ['A', 'B'], capabilities: {}, failRename: true });
      const tab = h.sheetTabs.querySelectorAll<HTMLButtonElement>('.fc-host__sheetbar-tab')[0];
      tab?.dispatchEvent(new KeyboardEvent('keydown', { key: 'F2', bubbles: true }));
      const input = h.sheetTabs.querySelector<HTMLInputElement>('input.fc-host__sheetbar-rename');
      if (!input) throw new Error('rename input not found');
      input.value = 'BadName';
      input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
      expect(input.getAttribute('aria-invalid')).toBe('true');
      expect(h.state.sheets[0]).toBe('A');
    });
  });

  describe('context menu / showMenu()', () => {
    it('opens with rename / insert / move / delete / hide buttons', () => {
      h = mount({ sheets: ['A', 'B', 'C'], capabilities: {} });
      const tab = h.sheetTabs.querySelectorAll<HTMLButtonElement>('.fc-host__sheetbar-tab')[1];
      if (!tab) throw new Error('tab not found');
      h.controller.showMenu(1, tab, 0, 0);

      expect(h.sheetMenu.hidden).toBe(false);
      const items = Array.from(
        h.sheetMenu.querySelectorAll<HTMLButtonElement>('.fc-sheetmenu__item'),
      );
      // 7 menu buttons: rename, insert, moveLeft, moveRight, delete, hide, unhide.
      expect(items.length).toBe(7);
      // Active is sheet 1, so moveLeft + moveRight are both enabled, delete enabled
      // (sheetCount > 1), unhide disabled (no hidden sheets yet).
      const enabled = items.map((b) => !b.disabled);
      // [rename, insert, moveLeft, moveRight, delete, hide, unhide]
      expect(enabled).toEqual([true, true, true, true, true, true, false]);
    });

    it('disables mutation entries when capabilities forbid them', () => {
      h = mount({
        sheets: ['A', 'B'],
        capabilities: { sheetMutate: false, sheetTabHidden: false },
      });
      const tab = h.sheetTabs.querySelectorAll<HTMLButtonElement>('.fc-host__sheetbar-tab')[0];
      if (!tab) throw new Error('tab not found');
      h.controller.showMenu(0, tab, 0, 0);
      const items = Array.from(
        h.sheetMenu.querySelectorAll<HTMLButtonElement>('.fc-sheetmenu__item'),
      );
      // rename, moveLeft, moveRight, delete, hide, unhide all gated on caps.
      // insert is unconditional (uses engine addSheet directly).
      expect(items[0]?.disabled).toBe(true); // rename
      expect(items[1]?.disabled).toBe(false); // insert
      expect(items[4]?.disabled).toBe(true); // delete (capability)
      expect(items[5]?.disabled).toBe(true); // hide
      expect(items[6]?.disabled).toBe(true); // unhide
    });

    it('closeMenu() hides the menu and clears children', () => {
      h = mount({ sheets: ['A', 'B'], capabilities: {} });
      const tab = h.sheetTabs.querySelectorAll<HTMLButtonElement>('.fc-host__sheetbar-tab')[0];
      if (!tab) throw new Error('tab not found');
      h.controller.showMenu(0, tab, 0, 0);
      expect(h.sheetMenu.hidden).toBe(false);
      h.controller.closeMenu();
      expect(h.sheetMenu.hidden).toBe(true);
      expect(h.sheetMenu.children.length).toBe(0);
    });

    it('Escape on document closes an open menu', () => {
      h = mount({ sheets: ['A', 'B'], capabilities: {} });
      const tab = h.sheetTabs.querySelectorAll<HTMLButtonElement>('.fc-host__sheetbar-tab')[0];
      if (!tab) throw new Error('tab not found');
      h.controller.showMenu(0, tab, 0, 0);
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
      expect(h.sheetMenu.hidden).toBe(true);
    });
  });

  describe('add-sheet button', () => {
    it('appends a sheet and switches to it', () => {
      h = mount({ sheets: ['A'], capabilities: {} });
      h.addSheetBtn.click();
      expect(h.state.sheets.length).toBe(2);
      expect(h.store.getState().data.sheetIndex).toBe(1);
      expect(h.sheetTabs.querySelectorAll('.fc-host__sheetbar-tab').length).toBe(2);
    });
  });

  describe('detach()', () => {
    it('removes document-level listeners and the menu element', () => {
      h = mount({ sheets: ['A', 'B'], capabilities: {} });
      const menu = h.sheetMenu;
      expect(menu.isConnected).toBe(true);
      h.controller.detach();
      expect(menu.isConnected).toBe(false);
      // After detach, document-level Escape should NOT reopen the listener.
      // (Best-effort: dispatching Escape no longer touches the menu, which is
      // already removed — this is a regression check that detach() doesn't throw.)
      expect(() =>
        document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' })),
      ).not.toThrow();
      // Restore so afterEach cleanup doesn't double-detach.
      h.detach = () => h?.host.remove();
    });
  });

  describe('aria labels', () => {
    it('applies strings from getStrings()', () => {
      h = mount({ sheets: ['A', 'B'], capabilities: {} });
      const t = (en as Strings).sheetTabs;
      expect(h.firstSheet.getAttribute('aria-label')).toBe(t.previousSheet);
      expect(h.lastSheet.getAttribute('aria-label')).toBe(t.nextSheet);
      expect(h.addSheetBtn.getAttribute('aria-label')).toBe(t.addSheet);
    });
  });

  describe('integration: rename then switch', () => {
    let local: Harness;
    beforeEach(() => {
      local = mount({ sheets: ['A', 'B', 'C'], capabilities: {} });
      h = local;
    });
    it('committed rename survives a switch round trip', () => {
      const tab = local.sheetTabs.querySelectorAll<HTMLButtonElement>('.fc-host__sheetbar-tab')[0];
      tab?.dispatchEvent(new KeyboardEvent('keydown', { key: 'F2', bubbles: true }));
      const input = local.sheetTabs.querySelector<HTMLInputElement>(
        'input.fc-host__sheetbar-rename',
      );
      if (!input) throw new Error('rename input not found');
      input.value = 'First';
      input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
      local.controller.switchSheet(1);
      local.controller.switchSheet(0);
      const tabs = local.sheetTabs.querySelectorAll<HTMLButtonElement>('.fc-host__sheetbar-tab');
      expect(tabs[0]?.textContent).toBe('First');
    });
  });
});
