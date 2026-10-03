import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createDefaultRibbonMenus } from '../../../../src/mount/toolbar-defaults.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators } from '../../../../src/store/store.js';
import {
  RIBBON_MENU_FACTORY_FOR_COMMAND,
  RIBBON_MENU_FACTORY_KEYS,
  RIBBON_MENU_FOR_COMMAND,
} from '../../../../src/toolbar/ribbon/activation.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { dynamicDropdownNoopOverrides, seedText, stubHelpers } from './fixtures.js';

vi.setConfig({ testTimeout: 20_000 });

describe('Spreadsheet.mountToolbar', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    sheet.dispose();
    host.remove();
  });

  it('renders the ribbon shell and returns an imperative instance', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });

    expect(tb.host).toBe(host);
    expect(tb.instance).toBe(sheet.instance);
    expect(host.querySelector('.fc-tb__ribbon-shell')).toBeTruthy();
    expect(tb.getActiveTab()).toBe('home');
    expect(tb.getCollapsed()).toBe(false);
    expect(tb.getFormulaBarVisible()).toBe(true);
    expect(tb.getTheme()).toBe('paper');

    tb.dispose();
    expect(host.children.length).toBe(0);
  });

  it('rehomes an omitted File/Help tab when the inherited platform changes to Mac', () => {
    sheet.instance.host.dataset.fcPlatform = 'default';
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    tb.setActiveTab('file');
    expect(tb.getActiveTab()).toBe('file');
    expect(host.querySelector('[data-ribbon-tab="file"]')).not.toBeNull();

    sheet.instance.host.dataset.fcPlatform = 'mac';
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 1, col: 0 });

    expect(host.dataset.fcPlatform).toBe('mac');
    expect(tb.getActiveTab()).toBe('home');
    expect(host.querySelector('[data-ribbon-tab="file"]')).toBeNull();
    expect(host.querySelector('[data-ribbon-tab="home"]')).not.toBeNull();
  });

  it('provides default menu factories for every shared ribbon menu slot', () => {
    const menus = createDefaultRibbonMenus(sheet.instance);
    const missing = RIBBON_MENU_FACTORY_KEYS.filter((key) => typeof menus[key] !== 'function');

    expect(missing).toEqual([]);
  });

  it('keeps default menu factories returning each activation menu id', () => {
    const menus = createDefaultRibbonMenus(sheet.instance);
    const mismatches: string[] = [];

    for (const [command, menuId] of Object.entries(RIBBON_MENU_FOR_COMMAND)) {
      const key = RIBBON_MENU_FACTORY_FOR_COMMAND[command];
      const menu = key ? menus[key]?.(command) : null;
      if (!menu) mismatches.push(`${command}:missing-factory`);
      else if (menu.id !== menuId) mismatches.push(`${command}:${key}:${menu.id}->${menuId}`);
    }

    expect(mismatches).toEqual([]);
  });

  it('dispatches ribbon commands and fires onCommand', () => {
    const onCommand = vi.fn();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
      onCommand,
    });

    // 'undoHome' is a built-in command handled by core (no hooks needed). It
    // returns true even when the history is empty — `undo()` returns false
    // but the dispatcher still claims the click.
    const applied = tb.applyCommand('undoHome');
    expect(applied).toBe(true);
    expect(onCommand).toHaveBeenCalledWith('undoHome', true);

    // Unknown ids fall through.
    const unknown = tb.applyCommand('not-a-real-command');
    expect(unknown).toBe(false);
    expect(onCommand).toHaveBeenLastCalledWith('not-a-real-command', false);

    tb.dispose();
  });

  it('routes PivotTable Fields command to the active pivot field list with fallback', () => {
    const openActivePivotFieldList = vi
      .spyOn(sheet.instance, 'openActivePivotFieldList')
      .mockReturnValue(false);
    const openWorkbookObjects = vi.spyOn(sheet.instance, 'openWorkbookObjects');
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });

    expect(tb.applyCommand('pivotFieldListView')).toBe(true);
    expect(openActivePivotFieldList).toHaveBeenCalledTimes(1);
    expect(openWorkbookObjects).toHaveBeenCalledTimes(1);

    openActivePivotFieldList.mockReturnValue(true);
    expect(tb.applyCommand('pivotFieldListView')).toBe(true);
    expect(openActivePivotFieldList).toHaveBeenCalledTimes(2);
    expect(openWorkbookObjects).toHaveBeenCalledTimes(1);

    tb.dispose();
  });

  it('routes direct ribbon commands through default hooks', async () => {
    seedText(sheet, 0, 0, 'teh  teh');
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
    });

    expect(tb.applyCommand('formatTableHome')).toBe(true);
    await Promise.resolve();
    expect(document.body.textContent).toContain('Create Table');
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    expect(sheet.instance.store.getState().tables.tables).toMatchObject([
      { style: 'medium', range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 } },
    ]);

    expect(tb.applyCommand('recordActions')).toBe(true);
    await Promise.resolve();
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      'Recorded selected range action',
    );
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();

    expect(tb.applyCommand('allScripts')).toBe(true);
    await Promise.resolve();
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      'Built-in scripts',
    );
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();

    expect(tb.applyCommand('outlineGroup')).toBe(true);
    expect(sheet.instance.store.getState().layout.outlineRows.get(0)).toBe(1);
    expect(tb.applyCommand('outlineHideDetail')).toBe(true);
    expect(sheet.instance.store.getState().layout.hiddenRows.has(0)).toBe(true);
    expect(tb.applyCommand('outlineShowDetail')).toBe(true);
    expect(sheet.instance.store.getState().layout.hiddenRows.has(0)).toBe(false);

    expect(tb.applyCommand('sheetViewSave')).toBe(true);
    expect(sheet.instance.store.getState().sheetViews.views).toHaveLength(1);
    const saved = sheet.instance.store.getState().sheetViews.views[0];
    sheet.instance.store.setState((state) => ({
      ...state,
      sheetViews: { ...state.sheetViews, activeViewId: saved?.id ?? null },
    }));
    expect(tb.applyCommand('sheetViewDelete')).toBe(true);
    expect(sheet.instance.store.getState().sheetViews.views).toHaveLength(0);

    expect(tb.applyCommand('spellingReview')).toBe(true);
    await Promise.resolve();
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      'Possible typo',
    );
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();

    expect(tb.applyCommand('filter')).toBe(true);
    expect(sheet.instance.store.getState().ui.filterRange).toEqual({
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 2,
      c1: 1,
    });

    tb.dispose();
  });

  it('clicks on ribbon tabs switch the active tab and rerender', () => {
    const onTabChange = vi.fn();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
      onTabChange,
    });

    const insertTab = host.querySelector<HTMLButtonElement>('[data-ribbon-tab="insert"]');
    expect(insertTab).toBeTruthy();
    insertTab?.click();

    expect(tb.getActiveTab()).toBe('insert');
    expect(onTabChange).toHaveBeenCalledWith('insert');

    tb.dispose();
  });

  it('focuses the active ribbon tab for F6 landmark navigation', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });

    expect(tb.focusActiveTab()).toBe(true);
    expect(document.activeElement).toBe(host.querySelector('[data-ribbon-tab="home"]'));

    tb.setActiveTab('data');
    expect(tb.focusActiveTab()).toBe(true);
    expect(document.activeElement).toBe(host.querySelector('[data-ribbon-tab="data"]'));

    tb.dispose();
  });

  it('reveals collapsed tabs without replacing buttons or changing the saved display mode', () => {
    const onTabChange = vi.fn();
    const onDisplayModeChange = vi.fn();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
      platform: 'mac',
      ribbonDisplayMode: 'tabsOnly',
      onTabChange,
      onDisplayModeChange,
    });
    const home = host.querySelector<HTMLButtonElement>('[data-ribbon-tab="home"]');
    const insert = host.querySelector<HTMLButtonElement>('[data-ribbon-tab="insert"]');
    if (!home || !insert) throw new Error('Missing ribbon tabs');
    home.click();
    expect(host.querySelector('.fc-tb__ribbon-shell--peek')).toBeTruthy();
    expect(host.querySelector('[data-ribbon-tab="home"]')).toBe(home);
    expect(tb.getDisplayMode()).toBe('tabsOnly');
    expect(onDisplayModeChange).not.toHaveBeenCalled();
    expect(onTabChange).not.toHaveBeenCalled();
    insert.click();
    expect(host.querySelector('[data-ribbon-tab="insert"]')).toBe(insert);
    expect(tb.getActiveTab()).toBe('insert');
    expect(onTabChange).toHaveBeenCalledExactlyOnceWith('insert');
    expect(host.querySelector<HTMLElement>('[data-ribbon-panel="home"]')?.hidden).toBe(true);
    expect(host.querySelector<HTMLElement>('[data-ribbon-panel="insert"]')?.hidden).toBe(false);
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(host.querySelector('.fc-tb__ribbon-shell--peek')).toBeFalsy();
    expect(document.activeElement).toBe(insert);
    tb.setActiveTab('data');
    expect(host.querySelector('.fc-tb__ribbon-shell--peek')).toBeFalsy();
    expect(tb.getDisplayMode()).toBe('tabsOnly');
    insert.disabled = true;
    tb.setActiveTab('insert');
    expect(tb.getActiveTab()).toBe('data');
    home.click();
    expect(host.querySelector('.fc-tb__ribbon-shell--peek')).toBeTruthy();
    expect(tb.applyCommand('bold')).toBe(true);
    expect(host.querySelector('.fc-tb__ribbon-shell--peek')).toBeFalsy();
    tb.dispose();
  });

  it('chooses an available initial tab when a host limits the ribbon tabs', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
      ribbonTabs: ['home'],
      activeTab: 'data',
    });
    expect(tb.getActiveTab()).toBe('home');
    expect(host.querySelectorAll('[data-ribbon-tab][aria-selected="true"]')).toHaveLength(1);
    expect(host.querySelector<HTMLElement>('[data-ribbon-panel="home"]')?.hidden).toBe(false);
    tb.dispose();
  });

  it.each([false, true])(
    'dismisses an intercepted command unless it opens a menu (%s)',
    (opensMenu) => {
      const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
        helpers: stubHelpers(),
        platform: 'mac',
        ribbonDisplayMode: 'tabsOnly',
        interceptCommand: (id) => {
          if (id !== 'bold') return false;
          if (opensMenu) {
            const menu = host.querySelector<HTMLElement>('#menu-paste');
            if (!menu) throw new Error('Missing paste menu');
            menu.hidden = false;
          }
          return true;
        },
      });
      host.querySelector<HTMLButtonElement>('[data-ribbon-tab="home"]')?.click();
      host.querySelector<HTMLButtonElement>('[data-ribbon-command="bold"]')?.click();
      expect(!!host.querySelector('.fc-tb__ribbon-shell--peek')).toBe(opensMenu);
      tb.dispose();
    },
  );

  it('supports Excel-style ribbon display modes', () => {
    const onDisplayModeChange = vi.fn();
    const onCollapsedChange = vi.fn();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
      onCollapsedChange,
      onDisplayModeChange,
    });

    expect(tb.getDisplayMode()).toBe('full');
    expect(tb.getCollapsed()).toBe(false);
    expect(
      host
        .querySelector('[data-ribbon-panel="home"]')
        ?.classList.contains('fc-tb__ribbon--office365-home'),
    ).toBe(false);
    expect(
      Array.from(
        host.querySelectorAll<HTMLElement>('[data-ribbon-panel="home"] .fc-tb__ribbon-label'),
      )
        .map((label) => label.textContent)
        .filter(Boolean),
    ).toContain('Clipboard');

    tb.setDisplayMode('singleLine');
    expect(tb.getDisplayMode()).toBe('singleLine');
    expect(tb.getCollapsed()).toBe(false);
    expect(host.querySelector('.fc-tb__ribbon-shell--singleLine')).toBeTruthy();
    expect(onDisplayModeChange).toHaveBeenLastCalledWith('singleLine');

    tb.setCollapsed(true);
    expect(tb.getDisplayMode()).toBe('tabsOnly');
    expect(tb.getCollapsed()).toBe(true);
    expect(host.querySelector('.fc-tb__ribbon-shell--tabsOnly')).toBeTruthy();
    expect(onCollapsedChange).toHaveBeenLastCalledWith(true);

    tb.setDisplayMenuOpen(true);
    const autoHideButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-display-option="autoHide"]',
    );
    expect(autoHideButton).toBeTruthy();
    autoHideButton?.click();
    expect(tb.getDisplayMode()).toBe('autoHide');
    expect(tb.getCollapsed()).toBe(true);
    expect(host.querySelector('.fc-tb__ribbon-shell--autoHide')).toBeTruthy();

    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Alt', bubbles: true }));
    expect(host.querySelector('.fc-tb__ribbon-shell--autoHidePeek')).toBeTruthy();
    expect(
      host.querySelector('.fc-tb__ribbon-shell')?.getAttribute('data-ribbon-auto-hide-peek'),
    ).toBe('true');

    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(host.querySelector('.fc-tb__ribbon-shell--autoHidePeek')).toBeFalsy();

    tb.dispose();
  });

  it('routes hook calls into opts.hooks when present', () => {
    const copy = vi.fn();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
      hooks: { clipboard: { copy, cut: vi.fn(), paste: vi.fn() } },
    });

    tb.applyCommand('copy');
    expect(copy).toHaveBeenCalledTimes(1);

    tb.dispose();
  });

  it('dispose detaches the click listener and store subscription', () => {
    const onCommand = vi.fn();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
      onCommand,
    });

    tb.dispose();

    const tabBtn = document.createElement('button');
    tabBtn.dataset.ribbonTab = 'insert';
    host.appendChild(tabBtn);
    tabBtn.click();

    // After dispose, the click listener is gone so the active tab doesn't change.
    expect(tb.getActiveTab()).toBe('home');
    expect(onCommand).not.toHaveBeenCalled();
  });

  it('dispose detaches the document-level dynamic dropdown listeners', () => {
    const added = vi.spyOn(document, 'addEventListener');
    const removed = vi.spyOn(document, 'removeEventListener');
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: dynamicDropdownNoopOverrides(),
      helpers: stubHelpers(),
    });
    const attachedTypes = added.mock.calls.map(([type]) => type);
    const dynamicTypes = ['click', 'mousedown', 'focusin', 'mouseover', 'keydown'];
    for (const type of dynamicTypes) expect(attachedTypes).toContain(type);

    tb.dispose();

    const addedListeners = added.mock.calls.map(([, listener]) => listener);
    const removedListeners = removed.mock.calls.map(([, listener]) => listener);
    for (const listener of addedListeners) expect(removedListeners).toContain(listener);
    added.mockRestore();
    removed.mockRestore();
  });
});
