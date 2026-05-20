import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { setWorkbookStructureProtected } from '../../../src/commands/protection.js';
import type { EngineCapabilities } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import en from '../../../src/i18n/en.js';
import type { Strings } from '../../../src/i18n/strings.js';
import { attachSheetTabsController } from '../../../src/mount/sheet-tabs-controller.js';
import {
  SHEET_TAB_COLOR_CHOICES,
  sheetTabColorChoiceLabel,
} from '../../../src/sheet-tab-colors.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

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
    setSheetTabHidden(): boolean {
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
      expect(h.sheetTabs.dataset.fcSheetOverflow).toBe('true');
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

    it('marks single-sheet workbooks as non-overflowing', () => {
      h = mount({ sheets: ['Only'], capabilities: {} });
      expect(h.sheetTabs.dataset.fcSheetOverflow).toBe('false');
    });

    it('keeps tab button DOM on the shared sheet tab helper', () => {
      const source = readFileSync(join(root, 'src/mount/sheet-tabs-controller.ts'), 'utf8');
      expect(source).toContain('createSheetTabButton({');
      expect(source).not.toContain("const tab = document.createElement('button')");
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

  describe('sheet navigation buttons', () => {
    afterEach(() => {
      vi.useRealTimers();
    });

    it('keeps click navigation as a single-step sheet switch', () => {
      h = mount({ sheets: ['A', 'B', 'C'], capabilities: {} });
      h.lastSheet.click();
      expect(h.store.getState().data.sheetIndex).toBe(1);
      h.firstSheet.click();
      expect(h.store.getState().data.sheetIndex).toBe(0);
    });

    it('repeats sheet navigation while the pointer is held', () => {
      vi.useFakeTimers();
      h = mount({ sheets: ['A', 'B', 'C', 'D'], capabilities: {} });

      h.lastSheet.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true, button: 0 }));
      vi.advanceTimersByTime(350);
      expect(h.store.getState().data.sheetIndex).toBe(1);
      vi.advanceTimersByTime(120);
      expect(h.store.getState().data.sheetIndex).toBe(2);
      vi.advanceTimersByTime(120);
      expect(h.store.getState().data.sheetIndex).toBe(3);

      h.lastSheet.dispatchEvent(new MouseEvent('pointerup', { bubbles: true, button: 0 }));
      h.lastSheet.click();
      expect(h.store.getState().data.sheetIndex).toBe(3);
    });

    it('cancels a pending repeat when the pointer leaves the button', () => {
      vi.useFakeTimers();
      h = mount({ sheets: ['A', 'B', 'C'], capabilities: {} });

      h.lastSheet.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true, button: 0 }));
      vi.advanceTimersByTime(200);
      h.lastSheet.dispatchEvent(new MouseEvent('pointerleave', { bubbles: true, button: 0 }));
      vi.advanceTimersByTime(350);

      expect(h.store.getState().data.sheetIndex).toBe(0);
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

    it('Home and End move to the first and last visible sheet tabs', () => {
      h = mount({ sheets: ['A', 'B', 'C', 'D'], capabilities: {} }, { initialIndex: 1 });
      h.store.setState((s) => ({
        ...s,
        layout: { ...s.layout, hiddenSheets: new Set([2]) },
      }));
      h.controller.update();
      h.sheetTabs
        .querySelector<HTMLButtonElement>('.fc-host__sheetbar-tab[aria-selected="true"]')
        ?.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true }));
      expect(h.store.getState().data.sheetIndex).toBe(3);
      h.sheetTabs
        .querySelector<HTMLButtonElement>('.fc-host__sheetbar-tab[aria-selected="true"]')
        ?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', bubbles: true }));
      expect(h.store.getState().data.sheetIndex).toBe(0);
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
      expect(items[6]?.dataset.disabledReason).toBe(en.ribbonMenu.sheetUnhideRequiresHiddenSheet);
      expect(items[6]?.getAttribute('aria-description')).toBe(
        en.ribbonMenu.sheetUnhideRequiresHiddenSheet,
      );
      expect(items[6]?.title).toBe(
        `${en.sheetTabs.unhideSheet}\n${en.ribbonMenu.sheetUnhideRequiresHiddenSheet}`,
      );
      expect(h.sheetMenu.querySelector('.fc-sheetmenu__colors')).not.toBeNull();
    });

    it('sets and clears an Excel-style tab color from the sheet menu palette', () => {
      h = mount({ sheets: ['A', 'B'], capabilities: {} });
      const tab = h.sheetTabs.querySelectorAll<HTMLButtonElement>('.fc-host__sheetbar-tab')[0];
      if (!tab) throw new Error('tab not found');
      h.controller.showMenu(0, tab, 0, 0);

      const swatches = h.sheetMenu.querySelectorAll<HTMLButtonElement>('.fc-sheetmenu__swatch');
      expect(swatches.length).toBe(SHEET_TAB_COLOR_CHOICES.length);
      expect(
        Array.from(swatches).map((button) => ({
          checked: button.getAttribute('aria-checked'),
          color: button.style.getPropertyValue('--fc-sheet-tab-color') || null,
          label: button.getAttribute('aria-label'),
        })),
      ).toEqual(
        SHEET_TAB_COLOR_CHOICES.map((choice) => ({
          checked: choice.color === null ? 'true' : 'false',
          color: choice.color,
          label:
            choice.color === null
              ? en.sheetTabs.noColor
              : `${en.sheetTabs.tabColor}: ${sheetTabColorChoiceLabel(choice, en.sheetTabs)} ${
                  choice.color
                }`,
        })),
      );
      swatches[1]?.click();

      expect(h.store.getState().layout.sheetTabColors.get(0)).toBe('#c00000');
      const coloredTab =
        h.sheetTabs.querySelectorAll<HTMLButtonElement>('.fc-host__sheetbar-tab')[0];
      expect(coloredTab?.dataset.fcSheetTabColor).toBe('true');
      expect(coloredTab?.style.getPropertyValue('--fc-sheet-tab-color')).toBe('#c00000');

      h.controller.showMenu(0, coloredTab as HTMLButtonElement, 0, 0);
      h.sheetMenu.querySelector<HTMLButtonElement>('.fc-sheetmenu__swatch--none')?.click();
      expect(h.store.getState().layout.sheetTabColors.has(0)).toBe(false);
    });

    it('shows each hidden sheet as an unhide target in the sheet menu', () => {
      h = mount({ sheets: ['A', 'B', 'C', 'D'], capabilities: {} });
      h.store.setState((s) => ({
        ...s,
        layout: { ...s.layout, hiddenSheets: new Set([1, 3]) },
      }));
      h.controller.update();
      const tab = h.sheetTabs.querySelectorAll<HTMLButtonElement>('.fc-host__sheetbar-tab')[0];
      if (!tab) throw new Error('tab not found');
      h.controller.showMenu(0, tab, 0, 0);

      const items = Array.from(
        h.sheetMenu.querySelectorAll<HTMLButtonElement>('.fc-sheetmenu__item'),
      );
      expect(items.map((item) => item.textContent)).toContain('Unhide B');
      expect(items.map((item) => item.textContent)).toContain('Unhide D');

      items.find((item) => item.textContent === 'Unhide D')?.click();
      expect(h.store.getState().layout.hiddenSheets.has(3)).toBe(false);
      expect(h.store.getState().layout.hiddenSheets.has(1)).toBe(true);
      expect(h.store.getState().data.sheetIndex).toBe(3);
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
      expect(items[0]?.dataset.disabledReason).toBe(en.ribbonMenu.sheetActionUnavailable);
      expect(items[4]?.dataset.disabledReason).toBe(en.ribbonMenu.sheetMutationUnavailable);
      expect(items[5]?.dataset.disabledReason).toBe(en.ribbonMenu.sheetActionUnavailable);
    });

    it('projects workbook protection reasons on sheet menu actions', () => {
      h = mount({ sheets: ['A', 'B'], capabilities: {} });
      setWorkbookStructureProtected(h.store, true);
      h.controller.update();
      const tab = h.sheetTabs.querySelectorAll<HTMLButtonElement>('.fc-host__sheetbar-tab')[0];
      if (!tab) throw new Error('tab not found');
      h.controller.showMenu(0, tab, 0, 0);
      const items = Array.from(
        h.sheetMenu.querySelectorAll<HTMLButtonElement>('.fc-sheetmenu__item'),
      );

      expect(items[0]?.disabled).toBe(true);
      expect(items[1]?.disabled).toBe(true);
      expect(items[0]?.dataset.disabledReason).toBe(
        en.ribbonMenu.workbookStructureProtectedBlocked,
      );
      expect(items[1]?.dataset.disabledReason).toBe(
        en.ribbonMenu.workbookStructureProtectedBlocked,
      );
      expect(items[0]?.getAttribute('aria-description')).toBe(
        en.ribbonMenu.workbookStructureProtectedBlocked,
      );
      expect(items[0]?.title).toBe(
        `${en.sheetTabs.rename}\n${en.ribbonMenu.workbookStructureProtectedBlocked}`,
      );
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

    it('Escape returns focus to the sheet tab that opened the menu', () => {
      h = mount({ sheets: ['A', 'B'], capabilities: {} });
      const tab = h.sheetTabs.querySelectorAll<HTMLButtonElement>('.fc-host__sheetbar-tab')[0];
      if (!tab) throw new Error('tab not found');
      h.controller.showMenu(0, tab, 0, 0);
      expect(document.activeElement).not.toBe(tab);

      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));

      expect(document.activeElement).toBe(tab);
    });

    it('supports Arrow/Home/End keyboard navigation inside the sheet menu', () => {
      h = mount({ sheets: ['A', 'B', 'C'], capabilities: {} });
      const tab = h.sheetTabs.querySelectorAll<HTMLButtonElement>('.fc-host__sheetbar-tab')[1];
      if (!tab) throw new Error('tab not found');
      h.controller.showMenu(1, tab, 0, 0);
      const items = Array.from(
        h.sheetMenu.querySelectorAll<HTMLButtonElement>('[role^="menuitem"]'),
      ).filter((item) => !item.disabled);
      expect(document.activeElement).toBe(items[0]);

      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
      expect(document.activeElement).toBe(items[1]);
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true }));
      expect(document.activeElement).toBe(items.at(-1));
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
      expect(document.activeElement).toBe(items[0]);
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', bubbles: true }));
      expect(document.activeElement).toBe(items[0]);
    });

    it('Enter and Space activate the focused sheet menu item', () => {
      h = mount({ sheets: ['A', 'B', 'C'], capabilities: {} });
      const tab = h.sheetTabs.querySelectorAll<HTMLButtonElement>('.fc-host__sheetbar-tab')[1];
      if (!tab) throw new Error('tab not found');
      h.controller.showMenu(1, tab, 0, 0);
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
      expect(h.state.sheets).toEqual(['B', 'A', 'C']);

      const movedTab = h.sheetTabs.querySelectorAll<HTMLButtonElement>('.fc-host__sheetbar-tab')[0];
      if (!movedTab) throw new Error('moved tab not found');
      h.controller.showMenu(0, movedTab, 0, 0);
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
      document.dispatchEvent(new KeyboardEvent('keydown', { key: ' ', bubbles: true }));
      expect(h.state.sheets).toEqual(['A', 'B', 'C']);
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

    it('projects disabled reasons on navigation and add-sheet buttons', () => {
      h = mount({ sheets: ['A', 'B'], capabilities: {} });
      expect(h.firstSheet.disabled).toBe(true);
      expect(h.firstSheet.dataset.disabledReason).toBe(en.sheetTabs.previousSheetUnavailable);
      expect(h.firstSheet.getAttribute('aria-description')).toBe(
        en.sheetTabs.previousSheetUnavailable,
      );
      expect(h.firstSheet.title).toBe(
        `${en.sheetTabs.previousSheet}\n${en.sheetTabs.previousSheetUnavailable}`,
      );
      expect(h.lastSheet.disabled).toBe(false);
      expect(h.lastSheet.dataset.disabledReason).toBeUndefined();
      expect(h.lastSheet.title).toBe(en.sheetTabs.nextSheet);

      setWorkbookStructureProtected(h.store, true);
      h.controller.update();
      expect(h.addSheetBtn.disabled).toBe(true);
      expect(h.addSheetBtn.dataset.disabledReason).toBe(
        en.ribbonMenu.workbookStructureProtectedBlocked,
      );
      expect(h.addSheetBtn.getAttribute('aria-description')).toBe(
        en.ribbonMenu.workbookStructureProtectedBlocked,
      );
      expect(h.addSheetBtn.title).toBe(
        `${en.sheetTabs.addSheet}\n${en.ribbonMenu.workbookStructureProtectedBlocked}`,
      );
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
