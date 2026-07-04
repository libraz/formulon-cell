import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { customCellStyleId } from '../../../src/commands/cell-styles.js';
import { captureSnapshot } from '../../../src/commands/clipboard/snapshot.js';
import { commentAt, setComment } from '../../../src/commands/comment.js';
import {
  customPivotTableStyleId,
  customTableStyleId,
} from '../../../src/commands/format-as-table.js';
import { hyperlinkAt, setHyperlink } from '../../../src/commands/hyperlinks.js';
import { setWorkbookStructureProtected } from '../../../src/commands/protection.js';
import { addrKey } from '../../../src/engine/address.js';
import { createDefaultRibbonMenus } from '../../../src/mount/toolbar-defaults.js';
import { Spreadsheet } from '../../../src/mount.js';
import { getPageSetup, mutators } from '../../../src/store/store.js';
import {
  RIBBON_AUDITED_DROPDOWN_COMMANDS,
  RIBBON_AUDITED_GALLERY_COMMANDS,
  RIBBON_AUDITED_PRIMARY_ACTION_SPLIT_COMMANDS,
  RIBBON_AUDITED_SPLIT_TOGGLE_COMMANDS,
  RIBBON_BORDERS_MENU_ID,
  RIBBON_DIALOG_COMMANDS,
  RIBBON_DISABLED_COMMANDS,
  RIBBON_DROPDOWN_COMMANDS,
  RIBBON_DYNAMIC_MENU_FIRST_COMMANDS,
  RIBBON_EXTERNAL_MENU_FIRST_COMMANDS,
  RIBBON_EXTERNAL_MENU_FOR_COMMAND,
  RIBBON_GALLERY_COMMANDS,
  RIBBON_MENU_FACTORY_FOR_COMMAND,
  RIBBON_MENU_FACTORY_KEYS,
  RIBBON_MENU_FOR_COMMAND,
  RIBBON_PRIMARY_ACTION_COMMANDS,
  RIBBON_PRIMARY_ACTION_SPLIT_COMMANDS,
  RIBBON_PRIMARY_FACE_MENU_COMMANDS,
  RIBBON_SPLIT_BUTTON_COMMANDS,
  RIBBON_SPLIT_TOGGLE_COMMANDS,
  RIBBON_TOGGLE_COMMANDS,
  ribbonActivationCategories,
  ribbonActivationForCommand,
} from '../../../src/toolbar/ribbon/activation.js';
import {
  RIBBON_DIALOG_OPENERS,
  RIBBON_FUNCTION_ARG_OPENERS,
  RIBBON_HOOK_DIALOG_COMMANDS,
  RIBBON_PRIMARY_SPLIT_DIALOG_COMMANDS,
} from '../../../src/toolbar/ribbon/command-tables.js';
import { SUPPORTED_CONDITIONAL_MENU_ACTIONS } from '../../../src/toolbar/ribbon/conditional-menu-action.js';
import {
  DYNAMIC_RIBBON_DROPDOWN_HANDLER_DATASET_KEYS,
  type DynamicDropdownsCtx,
  RIBBON_DROPDOWN_MENU_FOR_COMMAND,
} from '../../../src/toolbar/ribbon/dynamic-dropdowns.js';
import type { RibbonRenderHelpers } from '../../../src/toolbar/ribbon/render-ribbon.js';
import {
  HOME_MIXED_LAYOUT_GROUP_VARIANTS,
  HOME_STACKED_LAYOUT_GROUP_VARIANTS,
  HOME_TILE_LAYOUT_GROUP_VARIANTS,
  ribbonActivatableSurfaceCommandIds,
} from '../../../src/toolbar/ribbon-model.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

vi.setConfig({ testTimeout: 20_000 });

// Minimal helpers stub: enough for the renderer to emit a shell, no real
// dropdown DOM. The toolbar still needs `createSelect/Color/Icon/makeSvg`
// because every command path may reach them.
const stubHelpers = (): RibbonRenderHelpers => ({
  createSelect: () => document.createElement('div'),
  createColor: () => document.createElement('div'),
  createIcon: () => null,
  makeSvg: (_viewBox, _pathData, className) => {
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.setAttribute('class', className);
    return svg;
  },
  chevronPath: 'M0 0',
});

const waitFor = async (predicate: () => boolean, timeoutMs = 250): Promise<void> => {
  const deadline = Date.now() + timeoutMs;
  while (!predicate()) {
    if (Date.now() >= deadline) throw new Error('Timed out waiting for expected state.');
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
};

const seedNumber = (sheet: MountedStubSheet, row: number, col: number, value: number): void => {
  sheet.workbook.setNumber({ sheet: 0, row, col }, value);
  sheet.instance.store.setState((state) => {
    const cells = new Map(state.data.cells);
    cells.set(addrKey({ sheet: 0, row, col }), {
      value: { kind: 'number', value },
      formula: null,
    });
    return { ...state, data: { ...state.data, cells } };
  });
};

const seedText = (sheet: MountedStubSheet, row: number, col: number, value: string): void => {
  sheet.workbook.setText({ sheet: 0, row, col }, value);
  sheet.instance.store.setState((state) => {
    const cells = new Map(state.data.cells);
    cells.set(addrKey({ sheet: 0, row, col }), {
      value: { kind: 'text', value },
      formula: null,
    });
    return { ...state, data: { ...state.data, cells } };
  });
};

const dynamicDropdownNoopOverrides = (): Partial<DynamicDropdownsCtx> => ({
  applyRibbonPasteAction: vi.fn(),
  updatePasteMenu: vi.fn(),
  applyPivotTableAction: vi.fn(),
  applyDefinedNameAction: vi.fn(),
  applyLinksAction: vi.fn(),
  applyFillSeries: vi.fn(),
  updateFillMenu: vi.fn(),
  applyFillDirection: vi.fn(),
  applyClearAction: vi.fn(),
  updateClearMenu: vi.fn(),
  applyFreezeAction: vi.fn(),
  updateFreezeMenu: vi.fn(),
  applyTextOrientationAction: vi.fn(),
  updateTextOrientationMenu: vi.fn(),
  applyCellInsertAction: vi.fn(),
  updateCellInsertMenu: vi.fn(),
  applyCellDeleteAction: vi.fn(),
  updateCellDeleteMenu: vi.fn(),
  applyCellFormatAction: vi.fn(),
  applyPageBreakAction: vi.fn(),
  applySheetBackgroundAction: vi.fn(),
  applyPrintAreaAction: vi.fn(),
  applyArrangeAction: vi.fn(),
  applyUiTheme: vi.fn(),
  applySortMenuAction: vi.fn(),
  applyFindSelectAction: vi.fn(),
  applyAutoSumFormula: vi.fn(),
  applyFormulaAuditAction: vi.fn(),
  applyWatchAction: vi.fn(),
  applyReviewCommentAction: vi.fn(),
  applyProtectAction: vi.fn(),
  applyCalcOptionAction: vi.fn(),
  updateArrangeMenu: vi.fn(),
  updateCellStylesMenu: vi.fn(),
  updateCurrencyMenu: vi.fn(),
  updatePageBreaksMenu: vi.fn(),
  updatePrintAreaMenu: vi.fn(),
  updateProtectMenu: vi.fn(),
  updatePageThemeMenu: vi.fn(),
  updateReviewCommentsMenu: vi.fn(),
  updateSortMenu: vi.fn(),
  updateTableStylesMenu: vi.fn(),
  updateWatchMenu: vi.fn(),
  createRecommendedChartFromSelection: vi.fn(),
  createChartFromSelection: vi.fn(),
  chartKindFromAction: vi.fn((_action: string): 'column' => 'column'),
  insertPictureFromRibbon: vi.fn(),
  insertShapeFromRibbon: vi.fn(),
  insertScreenshotFromRibbon: vi.fn(),
  applyScriptAction: vi.fn(),
  applyPdfAction: vi.fn(),
  createTableFromSelection: vi.fn(),
  openTableStyleFooterAction: vi.fn(),
  applyCellStyleFromRibbon: vi.fn(),
  openCellStyleFooterAction: vi.fn(),
  applyCurrencyPreset: vi.fn(),
  openCurrencyFooterAction: vi.fn(),
  splitTextToColumns: vi.fn(),
  splitTextToColumnsCustom: vi.fn(),
  applyDataValidationAction: vi.fn(),
  applyAddInAction: vi.fn(),
  applyConditionalMenuAction: vi.fn(),
  applySymbolAction: vi.fn(),
  updateClearArrowsMenu: vi.fn(),
  updateDataValidationMenu: vi.fn(),
  updateLinksMenu: vi.fn(),
  updateErrorCheckingMenu: vi.fn(),
  updateFormatCellsMenu: vi.fn(),
});

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
    expect(host.querySelector('.demo__ribbon-shell')).toBeTruthy();
    expect(tb.getActiveTab()).toBe('home');
    expect(tb.getCollapsed()).toBe(false);
    expect(tb.getFormulaBarVisible()).toBe(true);
    expect(tb.getTheme()).toBe('light');

    tb.dispose();
    expect(host.children.length).toBe(0);
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

  it('projects active formatting onto ribbon toggle buttons', () => {
    mutators.setCellFormat(sheet.instance.store, { sheet: 0, row: 0, col: 0 }, { bold: true });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });

    const bold = host.querySelector<HTMLButtonElement>('[data-ribbon-command="bold"]');
    expect(bold).toBeTruthy();
    expect(bold?.classList.contains('demo__rb--active')).toBe(true);
    expect(bold?.getAttribute('aria-pressed')).toBe('true');

    mutators.setActive(sheet.instance.store, { sheet: 0, row: 1, col: 0 });
    expect(bold?.classList.contains('demo__rb--active')).toBe(false);
    expect(bold?.getAttribute('aria-pressed')).toBe('false');

    tb.dispose();
  });

  it('keeps an empty-cell format toggle visually active as pending input format', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });

    const bold = host.querySelector<HTMLButtonElement>('[data-ribbon-command="bold"]');
    expect(bold).toBeTruthy();
    expect(bold?.getAttribute('aria-pressed')).toBe('false');
    bold?.click();

    expect(bold?.classList.contains('demo__rb--active')).toBe(true);
    expect(bold?.getAttribute('aria-pressed')).toBe('true');
    expect(sheet.instance.store.getState().ui.pendingFormat).toEqual({
      addr: { sheet: 0, row: 0, col: 0 },
      format: { bold: true },
    });
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')).toBeUndefined();

    tb.dispose();
  });

  it('keeps Underline as a split toggle with single and double underline menu actions', async () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const underline = host.querySelector<HTMLButtonElement>('[data-ribbon-command="underline"]');
    expect(underline).toBeTruthy();
    expect(underline?.dataset.ribbonActivation).toBe('splitToggle');
    expect(underline?.dataset.ribbonMenuId).toBe('menu-underline');
    expect(underline?.getAttribute('aria-haspopup')).toBe('menu');
    expect(underline?.getAttribute('aria-pressed')).toBe('false');

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'underline', menuId: 'menu-underline' },
      underline,
    );
    const underlineItems = Array.from(
      host.querySelectorAll<HTMLButtonElement>('#menu-underline .app__menu-item--iconic'),
    );
    expect(underlineItems.map((item) => item.textContent)).toEqual([
      'Underline',
      'Double Underline',
    ]);

    const single = host.querySelector<HTMLButtonElement>(
      '#menu-underline [data-underline-action="single"]',
    );
    const singleEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(singleEvent, 'target', { value: single });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(singleEvent)).toBe(true);
    expect(sheet.instance.store.getState().ui.pendingFormat).toEqual({
      addr: { sheet: 0, row: 0, col: 0 },
      format: { underline: true },
    });

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'underline', menuId: 'menu-underline' },
      underline,
    );
    const double = host.querySelector<HTMLButtonElement>(
      '#menu-underline [data-underline-action="double"]',
    );
    const doubleEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(doubleEvent, 'target', { value: double });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(doubleEvent)).toBe(true);
    await Promise.resolve();
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'Double Underline',
    );
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dispose();
  });

  it('localizes the Underline split menu through the shared Home menu factory', async () => {
    sheet.dispose();
    sheet = await mountStubSheet({ locale: 'ja' });
    host.remove();
    host = document.createElement('div');
    document.body.appendChild(host);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const underline = host.querySelector<HTMLButtonElement>('[data-ribbon-command="underline"]');
    expect(underline).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'underline', menuId: 'menu-underline' },
      underline,
    );

    expect(
      Array.from(
        host.querySelectorAll<HTMLButtonElement>('#menu-underline .app__menu-item--iconic'),
      ).map((item) => item.textContent),
    ).toEqual(['下線', '二重下線']);

    const double = host.querySelector<HTMLButtonElement>(
      '#menu-underline [data-underline-action="double"]',
    );
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: double });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    await Promise.resolve();
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      '二重下線',
    );
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dispose();
  });

  it('localizes representative structured ribbon menus in Japanese', async () => {
    sheet.dispose();
    sheet = await mountStubSheet({ locale: 'ja' });
    host.remove();
    host = document.createElement('div');
    document.body.appendChild(host);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const menuTexts = (selector: string): string[] =>
      Array.from(host.querySelectorAll<HTMLButtonElement>(selector)).map((item) =>
        (item.textContent ?? '').replace('▶', '').trim(),
      );

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="paste"]')?.click();
    expect(menuTexts('#menu-paste .app__menu-item--iconic')).toEqual([
      '貼り付け',
      '数式',
      '数式と数値の書式',
      '値',
      '値と数値の書式',
      '書式',
      '行/列の入れ替え',
      '形式を選択して貼り付け…',
    ]);

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="conditional"]')?.click();
    expect(menuTexts('#menu-conditional > .app__menu-item')).toEqual([
      'セルの強調表示ルール',
      '上位/下位ルール',
      'データ バー',
      'カラー スケール',
      'アイコン セット',
      '新しいルール...',
      'ルールのクリア',
      'ルールの管理...',
    ]);

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="formatTableHome"]')?.click();
    expect(
      Array.from(
        host.querySelectorAll<HTMLElement>('#menu-table-style-home .app__tablestyle-heading'),
      ).map((heading) => heading.textContent),
    ).toEqual(['淡色', '中間', '濃色']);
    expect(menuTexts('#menu-table-style-home > .app__tablestyle-footer')).toEqual([
      '新しい表スタイル…',
      '新しいピボットテーブル スタイル…',
    ]);

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="findHome"]')?.click();
    expect(menuTexts('#menu-find-select .app__menu-item--iconic')).toEqual([
      '検索...',
      '置換...',
      'ジャンプ...',
      '条件を選択してジャンプ...',
      '数式',
      'コメントとメモ',
      '条件付き書式',
      '定数',
      'データの入力規則',
      'オブジェクトの選択',
      '選択ウィンドウ...',
    ]);

    tb.dispose();
  });

  it('opens the View Zoom dialog from the shared default hook', async () => {
    const refreshZoom = vi.fn();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
      refreshZoom,
    });
    tb.setActiveTab('view');

    const zoomButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="zoomDialog"]');
    expect(zoomButton).toBeTruthy();
    expect(zoomButton?.dataset.ribbonActivation).toBe('dialog');
    zoomButton?.click();

    const dialog = document.body.querySelector<HTMLElement>('.app__dlg');
    const input = dialog?.querySelector<HTMLInputElement>('input');
    expect(input).toBeTruthy();
    if (!input) throw new Error('Expected Zoom dialog input.');
    expect(input.value).toBe('100');
    input.value = '401';
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    expect(dialog?.textContent).toContain('Enter a zoom percentage from 50 to 400.');
    input.value = '125';
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    await Promise.resolve();

    expect(sheet.instance.store.getState().viewport.zoom).toBe(1.25);
    expect(refreshZoom).toHaveBeenCalled();

    tb.dispose();
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
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    expect(sheet.instance.store.getState().tables.tables).toMatchObject([
      { style: 'medium', range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 } },
    ]);

    expect(tb.applyCommand('recordActions')).toBe(true);
    await Promise.resolve();
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'Recorded selected range action',
    );
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

    expect(tb.applyCommand('allScripts')).toBe(true);
    await Promise.resolve();
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'Built-in scripts',
    );
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

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
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'Possible typo',
    );
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

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
        ?.classList.contains('demo__ribbon--office365-home'),
    ).toBe(false);
    expect(
      Array.from(
        host.querySelectorAll<HTMLElement>('[data-ribbon-panel="home"] .demo__ribbon-label'),
      )
        .map((label) => label.textContent)
        .filter(Boolean),
    ).toContain('Clipboard');

    tb.setDisplayMode('singleLine');
    expect(tb.getDisplayMode()).toBe('singleLine');
    expect(tb.getCollapsed()).toBe(false);
    expect(host.querySelector('.demo__ribbon-shell--singleLine')).toBeTruthy();
    expect(onDisplayModeChange).toHaveBeenLastCalledWith('singleLine');

    tb.setCollapsed(true);
    expect(tb.getDisplayMode()).toBe('tabsOnly');
    expect(tb.getCollapsed()).toBe(true);
    expect(host.querySelector('.demo__ribbon-shell--tabsOnly')).toBeTruthy();
    expect(onCollapsedChange).toHaveBeenLastCalledWith(true);

    tb.setDisplayMenuOpen(true);
    const autoHideButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-display-option="autoHide"]',
    );
    expect(autoHideButton).toBeTruthy();
    autoHideButton?.click();
    expect(tb.getDisplayMode()).toBe('autoHide');
    expect(tb.getCollapsed()).toBe(true);
    expect(host.querySelector('.demo__ribbon-shell--autoHide')).toBeTruthy();

    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Alt', bubbles: true }));
    expect(host.querySelector('.demo__ribbon-shell--autoHidePeek')).toBeTruthy();
    expect(
      host.querySelector('.demo__ribbon-shell')?.getAttribute('data-ribbon-auto-hide-peek'),
    ).toBe('true');

    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(host.querySelector('.demo__ribbon-shell--autoHidePeek')).toBeFalsy();

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

  it('sets, adds, and clears the current selection through the Print Area dropdown', () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 3 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('pageLayout');

    const printAreaButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="printArea"]',
    );
    expect(printAreaButton).toBeTruthy();
    printAreaButton?.click();
    const setPrintAreaButton = host.querySelector<HTMLButtonElement>(
      '[data-print-area-action="set"]',
    );
    const initialAddPrintAreaButton = host.querySelector<HTMLButtonElement>(
      '[data-print-area-action="add"]',
    );
    const initialClearPrintAreaButton = host.querySelector<HTMLButtonElement>(
      '[data-print-area-action="clear"]',
    );
    expect(setPrintAreaButton).toBeTruthy();
    expect(initialAddPrintAreaButton?.disabled).toBe(true);
    expect(initialAddPrintAreaButton?.getAttribute('aria-disabled')).toBe('true');
    expect(initialAddPrintAreaButton?.dataset.menuDisabledReason).toBe(
      'No print area has been set.',
    );
    expect(initialClearPrintAreaButton?.disabled).toBe(true);
    expect(initialClearPrintAreaButton?.dataset.menuDisabledReason).toBe(
      'No print area has been set.',
    );
    expect(
      host.querySelector<HTMLElement>('#menu-print-area .app__menu-icon--print-area-set'),
    ).toBeTruthy();
    const setEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(setEvent, 'target', { value: setPrintAreaButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(setEvent)).toBe(true);

    expect(getPageSetup(sheet.instance.store.getState(), 0).printArea).toBe('B2:D3');

    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 5, c0: 5, r1: 5, c1: 6 });
    printAreaButton?.click();
    const addPrintAreaButton = host.querySelector<HTMLButtonElement>(
      '[data-print-area-action="add"]',
    );
    expect(addPrintAreaButton).toBeTruthy();
    expect(addPrintAreaButton?.disabled).toBe(false);
    expect(addPrintAreaButton?.dataset.menuDisabledReason).toBeUndefined();
    expect(
      host
        .querySelector<HTMLButtonElement>('[data-print-area-action="clear"]')
        ?.getAttribute('aria-disabled'),
    ).toBe('false');
    const addEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(addEvent, 'target', { value: addPrintAreaButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(addEvent)).toBe(true);

    expect(getPageSetup(sheet.instance.store.getState(), 0).printArea).toBe('B2:D3,F6:G6');

    printAreaButton?.click();
    const clearPrintAreaButton = host.querySelector<HTMLButtonElement>(
      '[data-print-area-action="clear"]',
    );
    expect(clearPrintAreaButton).toBeTruthy();
    const clearEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(clearEvent, 'target', { value: clearPrintAreaButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(clearEvent)).toBe(true);

    expect(getPageSetup(sheet.instance.store.getState(), 0).printArea).toBeUndefined();

    printAreaButton?.click();
    expect(host.querySelector<HTMLButtonElement>('[data-print-area-action="add"]')?.disabled).toBe(
      true,
    );
    expect(
      host.querySelector<HTMLButtonElement>('[data-print-area-action="clear"]')?.disabled,
    ).toBe(true);

    tb.dispose();
  });

  it('renders Page Theme as a visual gallery and applies the selected theme', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('pageLayout');

    const themeButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="pageTheme"]');
    expect(themeButton).toBeTruthy();
    expect(themeButton?.dataset.ribbonActivation).toBe('gallery');
    themeButton?.click();
    const menu = host.querySelector<HTMLElement>('#menu-page-theme');
    expect(menu?.classList.contains('app__menu--visual')).toBe(true);
    expect(menu?.querySelectorAll('.app__visual-tile')).toHaveLength(3);

    const lightButton = host.querySelector<HTMLButtonElement>('[data-page-theme-action="light"]');
    const darkButton = host.querySelector<HTMLButtonElement>('[data-page-theme-action="dark"]');
    expect(lightButton?.getAttribute('role')).toBe('menuitemradio');
    expect(lightButton?.getAttribute('aria-checked')).toBe('true');
    expect(lightButton?.classList.contains('app__visual-tile--active')).toBe(true);
    expect(darkButton?.getAttribute('aria-checked')).toBe('false');
    expect(darkButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: darkButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(sheet.instance.store.getState().ui.theme).toBe('dark');

    themeButton?.click();
    expect(lightButton?.getAttribute('aria-checked')).toBe('false');
    expect(darkButton?.getAttribute('aria-checked')).toBe('true');
    expect(darkButton?.classList.contains('app__visual-tile--active')).toBe(true);
    const contrastButton = host.querySelector<HTMLButtonElement>(
      '[data-page-theme-action="contrast"]',
    );
    expect(contrastButton).toBeTruthy();
    const contrastEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(contrastEvent, 'target', { value: contrastButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(contrastEvent)).toBe(true);
    expect(sheet.instance.store.getState().ui.theme).toBe('contrast');

    tb.dispose();
  });

  it('inserts and deletes cells through the Home Cells dropdowns', () => {
    seedNumber(sheet, 1, 1, 10);
    seedNumber(sheet, 2, 1, 20);
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const insertButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="insertRows"]',
    );
    expect(insertButton).toBeTruthy();
    insertButton?.click();
    expect(host.querySelectorAll('#menu-insert-cells .app__menu-item--iconic').length).toBe(4);
    expect(host.querySelector<HTMLButtonElement>('[data-cell-insert="sheet"]')?.disabled).toBe(
      false,
    );
    const insertCellsButton = host.querySelector<HTMLButtonElement>('[data-cell-insert="cells"]');
    expect(insertCellsButton).toBeTruthy();
    const insertEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(insertEvent, 'target', { value: insertCellsButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(insertEvent)).toBe(true);
    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 1 }).kind).toBe('blank');
    expect(sheet.workbook.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({
      kind: 'number',
      value: 10,
    });

    const deleteButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="deleteRows"]',
    );
    expect(deleteButton).toBeTruthy();
    deleteButton?.click();
    expect(host.querySelectorAll('#menu-delete-cells .app__menu-item--iconic').length).toBe(6);
    const disabledDeleteSheet = host.querySelector<HTMLButtonElement>('[data-cell-delete="sheet"]');
    expect(disabledDeleteSheet?.disabled).toBe(true);
    expect(disabledDeleteSheet?.dataset.menuDisabledReason).toBe(
      'A workbook must contain at least one visible sheet.',
    );
    const deleteCellsButton = host.querySelector<HTMLButtonElement>('[data-cell-delete="cells"]');
    expect(deleteCellsButton).toBeTruthy();
    const deleteEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(deleteEvent, 'target', { value: deleteCellsButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(deleteEvent)).toBe(true);
    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({
      kind: 'number',
      value: 10,
    });

    tb.dispose();
  });

  it('inserts and deletes rows, columns, and sheets through the Home Cells dropdowns', () => {
    seedNumber(sheet, 1, 0, 10);
    seedNumber(sheet, 2, 0, 20);
    seedNumber(sheet, 0, 1, 11);
    seedNumber(sheet, 0, 2, 22);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const insertButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="insertRows"]',
    );
    const deleteButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="deleteRows"]',
    );
    expect(insertButton).toBeTruthy();
    expect(deleteButton).toBeTruthy();
    const clickInsert = (action: string): void => {
      tb.dropdownsApi?.openDynamicRibbonDropdown(
        { command: 'insertRows', menuId: 'menu-insert-cells' },
        insertButton as HTMLButtonElement,
      );
      const button = host.querySelector<HTMLButtonElement>(`[data-cell-insert="${action}"]`);
      expect(button).toBeTruthy();
      const event = new MouseEvent('click', { bubbles: true });
      Object.defineProperty(event, 'target', { value: button });
      expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    };
    const clickDelete = (action: string): void => {
      tb.dropdownsApi?.openDynamicRibbonDropdown(
        { command: 'deleteRows', menuId: 'menu-delete-cells' },
        deleteButton as HTMLButtonElement,
      );
      const button = host.querySelector<HTMLButtonElement>(`[data-cell-delete="${action}"]`);
      expect(button).toBeTruthy();
      expect(button?.disabled).toBe(false);
      const event = new MouseEvent('click', { bubbles: true });
      Object.defineProperty(event, 'target', { value: button });
      expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    };

    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 0 });
    clickInsert('rows');
    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 0 }).kind).toBe('blank');
    expect(sheet.workbook.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({
      kind: 'number',
      value: 10,
    });
    clickDelete('rows');
    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'number',
      value: 10,
    });

    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 });
    clickInsert('cols');
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 1 }).kind).toBe('blank');
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({
      kind: 'number',
      value: 11,
    });
    clickDelete('cols');
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
      kind: 'number',
      value: 11,
    });

    const initialSheetCount = sheet.workbook.sheetCount;
    clickInsert('sheet');
    expect(sheet.workbook.sheetCount).toBe(initialSheetCount + 1);
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'deleteRows', menuId: 'menu-delete-cells' },
      deleteButton as HTMLButtonElement,
    );
    const deleteSheetButton = host.querySelector<HTMLButtonElement>('[data-cell-delete="sheet"]');
    expect(deleteSheetButton?.disabled).toBe(true);
    expect(deleteSheetButton?.getAttribute('aria-disabled')).toBe('true');
    expect(deleteSheetButton?.dataset.menuDisabledReason).toBe(
      'This workbook engine cannot remove sheets.',
    );

    tb.dispose();
  });

  it('explains workbook structure protection on disabled Insert/Delete Sheet menu items', () => {
    setWorkbookStructureProtected(sheet.instance.store, true);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const insertButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="insertRows"]',
    );
    const deleteButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="deleteRows"]',
    );

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'insertRows', menuId: 'menu-insert-cells' },
      insertButton as HTMLButtonElement,
    );
    const insertSheetButton = host.querySelector<HTMLButtonElement>('[data-cell-insert="sheet"]');
    expect(insertSheetButton?.disabled).toBe(true);
    expect(insertSheetButton?.dataset.menuDisabledReason).toBe('Workbook structure is protected.');

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'deleteRows', menuId: 'menu-delete-cells' },
      deleteButton as HTMLButtonElement,
    );
    const deleteSheetButton = host.querySelector<HTMLButtonElement>('[data-cell-delete="sheet"]');
    expect(deleteSheetButton?.disabled).toBe(true);
    expect(deleteSheetButton?.dataset.menuDisabledReason).toBe('Workbook structure is protected.');

    tb.dispose();
  });

  it('sorts the active column through the Home Sort dropdown', () => {
    seedNumber(sheet, 0, 0, 3);
    seedNumber(sheet, 1, 0, 1);
    seedNumber(sheet, 2, 0, 2);
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 });
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const sortButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="sortFilterHome"]',
    );
    expect(sortButton).toBeTruthy();
    sortButton?.click();
    expect(host.querySelectorAll('#menu-sort-home .app__menu-item--iconic').length).toBe(11);
    const ascendingButton = host.querySelector<HTMLButtonElement>('[data-sort="asc"]');
    expect(ascendingButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: ascendingButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);

    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'number',
      value: 1,
    });
    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'number',
      value: 2,
    });
    expect(sheet.workbook.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({
      kind: 'number',
      value: 3,
    });

    tb.dispose();
  });

  it('reflects filter state in the Sort & Filter dropdown', () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const sortButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="sortFilterHome"]',
    );
    expect(sortButton).toBeTruthy();
    sortButton?.click();
    const menu = host.querySelector<HTMLElement>('#menu-sort-home');
    const filterButton = menu?.querySelector<HTMLButtonElement>('[data-sort="filter"]');
    const clearButton = menu?.querySelector<HTMLButtonElement>('[data-sort="filter-clear"]');
    const reapplyButton = menu?.querySelector<HTMLButtonElement>('[data-sort="filter-reapply"]');
    expect(filterButton?.getAttribute('aria-pressed')).toBe('false');
    expect(clearButton?.disabled).toBe(true);
    expect(clearButton?.dataset.menuDisabledReason).toBe('There is no filter to clear.');
    expect(reapplyButton?.disabled).toBe(true);
    expect(reapplyButton?.dataset.menuDisabledReason).toBe(
      'There are no filter criteria to reapply.',
    );

    const range = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 };
    mutators.setFilterRange(sheet.instance.store, range);
    sortButton?.click();
    expect(filterButton?.getAttribute('aria-pressed')).toBe('true');
    expect(filterButton?.classList.contains('app__menu-item--active')).toBe(true);
    expect(clearButton?.disabled).toBe(false);
    expect(clearButton?.dataset.menuDisabledReason).toBeUndefined();
    expect(reapplyButton?.disabled).toBe(true);

    sheet.instance.store.setState((state) => ({
      ...state,
      ui: {
        ...state.ui,
        filterCriteria: [{ range, byCol: 0, hiddenValues: ['1'] }],
      },
    }));
    sortButton?.click();
    expect(reapplyButton?.disabled).toBe(false);
    expect(reapplyButton?.getAttribute('aria-disabled')).toBe('false');
    expect(reapplyButton?.dataset.menuDisabledReason).toBeUndefined();

    tb.dispose();
  });

  it('routes filter and manager actions through the Home Sort dropdown', async () => {
    seedText(sheet, 0, 0, 'Name');
    seedText(sheet, 0, 1, 'Score');
    seedText(sheet, 1, 0, 'Alice');
    seedNumber(sheet, 1, 1, 10);
    seedText(sheet, 2, 0, 'Bob');
    seedNumber(sheet, 2, 1, 20);
    seedText(sheet, 4, 0, 'Name');
    seedText(sheet, 5, 0, 'Bob');
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 });
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 2, col: 0 });
    const openConditionalDialog = vi
      .spyOn(sheet.instance, 'openConditionalDialog')
      .mockImplementation(() => undefined);
    const openNamedRangeDialog = vi
      .spyOn(sheet.instance, 'openNamedRangeDialog')
      .mockImplementation(() => undefined);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const sortButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="sortFilterHome"]',
    );
    expect(sortButton).toBeTruthy();
    const clickSort = async (action: string): Promise<void> => {
      tb.dropdownsApi?.openDynamicRibbonDropdown(
        { command: 'sortFilterHome', menuId: 'menu-sort-home' },
        sortButton as HTMLButtonElement,
      );
      const button = host.querySelector<HTMLButtonElement>(`[data-sort="${action}"]`);
      expect(button).toBeTruthy();
      const event = new MouseEvent('click', { bubbles: true });
      Object.defineProperty(event, 'target', { value: button });
      expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
      await Promise.resolve();
    };

    await clickSort('filter');
    expect(sheet.instance.store.getState().ui.filterRange).toEqual({
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 2,
      c1: 1,
    });

    await clickSort('filter-by-value');
    expect(sheet.instance.store.getState().layout.hiddenRows.has(1)).toBe(true);
    expect(sheet.instance.store.getState().layout.hiddenRows.has(2)).toBe(false);

    await clickSort('filter-clear');
    expect(sheet.instance.store.getState().ui.filterRange).toBeNull();
    expect(sheet.instance.store.getState().layout.hiddenRows.has(1)).toBe(false);

    await clickSort('filter-advanced');
    const advancedDialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(advancedDialog?.textContent).toContain('Advanced Filter');
    expect(advancedDialog?.querySelector('.fc-advfilter__ranges')).toBeTruthy();
    const inputs = Array.from(advancedDialog?.querySelectorAll<HTMLInputElement>('input') ?? []);
    expect(inputs).toHaveLength(4);
    const listInput = inputs[0];
    const criteriaInput = inputs[1];
    if (!listInput || !criteriaInput) throw new Error('Expected Advanced Filter range inputs.');
    listInput.value = 'A1:B3';
    criteriaInput.value = 'A5:A6';
    advancedDialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    expect(sheet.instance.store.getState().layout.hiddenRows.has(1)).toBe(true);
    expect(sheet.instance.store.getState().layout.hiddenRows.has(2)).toBe(false);

    await clickSort('conditional');
    await clickSort('named');
    expect(openConditionalDialog).toHaveBeenCalledTimes(1);
    expect(openNamedRangeDialog).toHaveBeenCalledTimes(1);

    tb.dispose();
  });

  it('opens custom sort and remove duplicates from the Home Sort dropdown', async () => {
    seedNumber(sheet, 0, 0, 2);
    seedNumber(sheet, 1, 0, 1);
    seedNumber(sheet, 2, 0, 1);
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 });
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const sortButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="sortFilterHome"]',
    );
    expect(sortButton).toBeTruthy();
    sortButton?.click();
    const customButton = host.querySelector<HTMLButtonElement>('[data-sort="custom"]');
    expect(customButton).toBeTruthy();
    const customEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(customEvent, 'target', { value: customButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(customEvent)).toBe(true);

    const sortDialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(sortDialog?.textContent).toContain('Sort by');
    expect(sortDialog?.textContent).toContain('Add Level');
    expect(sortDialog?.querySelectorAll('.fc-sortdlg__level')).toHaveLength(1);
    sortDialog?.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter' }));
    await Promise.resolve();
    await Promise.resolve();

    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'number',
      value: 1,
    });

    sortButton?.click();
    const dedupeButton = host.querySelector<HTMLButtonElement>('[data-sort="dedupe"]');
    expect(dedupeButton).toBeTruthy();
    const dedupeEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(dedupeEvent, 'target', { value: dedupeButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(dedupeEvent)).toBe(true);

    const dedupeDialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(dedupeDialog?.textContent).toContain('Remove Duplicates');
    expect(dedupeDialog?.querySelectorAll('.fc-dedupedlg__column')).toHaveLength(1);
    dedupeDialog?.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter' }));
    await Promise.resolve();
    await Promise.resolve();

    expect(sheet.workbook.getValue({ sheet: 0, row: 2, col: 0 }).kind).toBe('blank');

    tb.dispose();
  });

  it('opens Text to Columns delimiter dialog from primary click and keeps presets secondary', async () => {
    seedText(sheet, 0, 0, 'alpha,beta');
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('data');

    const textToColumnsButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="textToColumns"]',
    );
    expect(textToColumnsButton).toBeTruthy();
    textToColumnsButton?.click();
    const dialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(dialog?.textContent).toContain('Convert Text to Columns');
    expect(dialog?.textContent).toContain('Original data type');
    expect(dialog?.textContent).toContain('Data preview');
    const comma = dialog?.querySelector<HTMLInputElement>('[data-dialog-field="delimiter-,"]');
    expect(comma?.checked).toBe(true);
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();

    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'alpha',
    });
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
      kind: 'text',
      value: 'beta',
    });

    seedText(sheet, 1, 0, 'one,two');
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 0 });
    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'textToColumns',
      menuId: 'menu-text-to-columns',
    });
    expect(host.querySelectorAll('#menu-text-to-columns .app__menu-item--iconic').length).toBe(5);
    const commaButton = host.querySelector<HTMLButtonElement>(
      '[data-text-to-columns-delimiter=","]',
    );
    expect(commaButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: commaButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);

    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'text',
      value: 'one',
    });
    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({
      kind: 'text',
      value: 'two',
    });

    tb.dispose();
  });

  it('clears comments and hyperlinks through the Home Clear dropdown', () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setComment(sheet.instance.store, { sheet: 0, row: 0, col: 0 }, 'note', sheet.workbook);
    setHyperlink(
      sheet.instance.store,
      { sheet: 0, row: 0, col: 0 },
      'https://example.test',
      sheet.workbook,
    );
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const clearButton = host.querySelector<HTMLButtonElement>(
      '.demo__ribbon-group--editing [data-ribbon-command="clearFormat"]',
    );
    expect(clearButton).toBeTruthy();
    expect(clearButton?.dataset.ribbonActivation).toBe('dropdown');
    expect(clearButton?.dataset.ribbonMenuId).toBe('menu-clear');
    clearButton?.click();
    expect(host.querySelectorAll('#menu-clear .app__menu-item--iconic').length).toBe(7);
    const clearCommentsButton = host.querySelector<HTMLButtonElement>('[data-clear="comments"]');
    expect(clearCommentsButton).toBeTruthy();
    expect(clearCommentsButton?.getAttribute('aria-disabled')).toBe('false');
    const commentsEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(commentsEvent, 'target', { value: clearCommentsButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(commentsEvent)).toBe(true);
    expect(commentAt(sheet.instance.store.getState(), { sheet: 0, row: 0, col: 0 })).toBeNull();
    expect(hyperlinkAt(sheet.instance.store.getState(), { sheet: 0, row: 0, col: 0 })).toBe(
      'https://example.test',
    );

    clearButton?.click();
    expect(clearCommentsButton?.getAttribute('aria-disabled')).toBe('true');
    expect(clearCommentsButton?.getAttribute('aria-description')).toBe(
      'Nothing matching this clear option is selected.',
    );
    expect(clearCommentsButton?.dataset.menuDisabledReason).toBe(
      'Nothing matching this clear option is selected.',
    );
    const clearHyperlinksButton = host.querySelector<HTMLButtonElement>(
      '[data-clear="remove-hyperlinks"]',
    );
    expect(clearHyperlinksButton).toBeTruthy();
    expect(clearHyperlinksButton?.getAttribute('aria-disabled')).toBe('false');
    expect(clearHyperlinksButton?.getAttribute('aria-description')).toBeNull();
    expect(clearHyperlinksButton?.dataset.menuDisabledReason).toBeUndefined();
    const hyperlinksEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(hyperlinksEvent, 'target', { value: clearHyperlinksButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(hyperlinksEvent)).toBe(true);
    expect(hyperlinkAt(sheet.instance.store.getState(), { sheet: 0, row: 0, col: 0 })).toBeNull();
    clearButton?.click();
    expect(clearHyperlinksButton?.getAttribute('aria-disabled')).toBe('true');
    expect(clearHyperlinksButton?.dataset.menuDisabledReason).toBe(
      'Nothing matching this clear option is selected.',
    );

    tb.dispose();
  });

  it('does not enable Clear Formats for comment-only or empty format metadata', () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setComment(sheet.instance.store, { sheet: 0, row: 0, col: 0 }, 'note', sheet.workbook);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const clearButton = host.querySelector<HTMLButtonElement>(
      '.demo__ribbon-group--editing [data-ribbon-command="clearFormat"]',
    );
    expect(clearButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'clearFormat', menuId: 'menu-clear' },
      clearButton as HTMLButtonElement,
    );
    const clearFormats = host.querySelector<HTMLButtonElement>('[data-clear="formats"]');
    const clearComments = host.querySelector<HTMLButtonElement>('[data-clear="comments"]');
    const clearAll = host.querySelector<HTMLButtonElement>('[data-clear="all"]');
    expect(clearFormats?.getAttribute('aria-disabled')).toBe('true');
    expect(clearComments?.getAttribute('aria-disabled')).toBe('false');
    expect(clearAll?.getAttribute('aria-disabled')).toBe('false');

    setComment(sheet.instance.store, { sheet: 0, row: 0, col: 0 }, '', sheet.workbook);
    sheet.instance.store.setState((s) => {
      const formats = new Map(s.format.formats);
      formats.set('0:0:0', {});
      return { ...s, format: { ...s.format, formats } };
    });
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'clearFormat', menuId: 'menu-clear' },
      clearButton as HTMLButtonElement,
    );
    expect(clearFormats?.getAttribute('aria-disabled')).toBe('true');
    expect(clearComments?.getAttribute('aria-disabled')).toBe('true');
    expect(clearAll?.getAttribute('aria-disabled')).toBe('true');

    tb.dispose();
  });

  it('clears contents, formats, conditional rules, and all state through the Home Clear dropdown', () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const clearButton = host.querySelector<HTMLButtonElement>(
      '.demo__ribbon-group--editing [data-ribbon-command="clearFormat"]',
    );
    expect(clearButton).toBeTruthy();
    const clickClear = (action: string): void => {
      tb.dropdownsApi?.openDynamicRibbonDropdown(
        { command: 'clearFormat', menuId: 'menu-clear' },
        clearButton as HTMLButtonElement,
      );
      const button = host.querySelector<HTMLButtonElement>(`[data-clear="${action}"]`);
      expect(button).toBeTruthy();
      expect(button?.disabled).toBe(false);
      const event = new MouseEvent('click', { bubbles: true });
      Object.defineProperty(event, 'target', { value: button });
      expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    };

    seedNumber(sheet, 0, 0, 42);
    clickClear('contents');
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');

    mutators.setCellFormat(sheet.instance.store, { sheet: 0, row: 0, col: 0 }, { bold: true });
    clickClear('formats');
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')).toBeUndefined();

    mutators.addConditionalRule(sheet.instance.store, {
      kind: 'data-bar',
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      color: '#638ec6',
    });
    clickClear('conditional');
    expect(sheet.instance.store.getState().conditional.rules).toEqual([]);

    seedNumber(sheet, 0, 0, 7);
    mutators.setCellFormat(sheet.instance.store, { sheet: 0, row: 0, col: 0 }, { italic: true });
    setComment(sheet.instance.store, { sheet: 0, row: 0, col: 0 }, 'note', sheet.workbook);
    setHyperlink(
      sheet.instance.store,
      { sheet: 0, row: 0, col: 0 },
      'https://example.test',
      sheet.workbook,
    );
    mutators.addConditionalRule(sheet.instance.store, {
      kind: 'data-bar',
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      color: '#638ec6',
    });
    clickClear('all');
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')).toBeUndefined();
    expect(commentAt(sheet.instance.store.getState(), { sheet: 0, row: 0, col: 0 })).toBeNull();
    expect(hyperlinkAt(sheet.instance.store.getState(), { sheet: 0, row: 0, col: 0 })).toBeNull();
    expect(sheet.instance.store.getState().conditional.rules).toEqual([]);

    tb.dispose();
  });

  it('enables Clear Formats when the active empty cell only has pending input format', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const bold = host.querySelector<HTMLButtonElement>('[data-ribbon-command="bold"]');
    bold?.click();
    expect(sheet.instance.store.getState().ui.pendingFormat).toEqual({
      addr: { sheet: 0, row: 0, col: 0 },
      format: { bold: true },
    });

    const clearButton = host.querySelector<HTMLButtonElement>(
      '.demo__ribbon-group--editing [data-ribbon-command="clearFormat"]',
    );
    expect(clearButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'clearFormat', menuId: 'menu-clear' },
      clearButton,
    );
    const clearFormats = host.querySelector<HTMLButtonElement>('[data-clear="formats"]');
    expect(clearFormats?.disabled).toBe(false);
    expect(clearFormats?.getAttribute('aria-disabled')).toBe('false');

    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: clearFormats });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(sheet.instance.store.getState().ui.pendingFormat).toBeNull();
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')).toBeUndefined();

    tb.dispose();
  });

  it('enables Clear All when the active empty cell only has pending input format', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="bold"]')?.click();
    host.querySelector<HTMLButtonElement>('[data-ribbon-command="italic"]')?.click();
    expect(sheet.instance.store.getState().ui.pendingFormat).toEqual({
      addr: { sheet: 0, row: 0, col: 0 },
      format: { bold: true, italic: true },
    });

    const clearButton = host.querySelector<HTMLButtonElement>(
      '.demo__ribbon-group--editing [data-ribbon-command="clearFormat"]',
    );
    expect(clearButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'clearFormat', menuId: 'menu-clear' },
      clearButton as HTMLButtonElement,
    );
    const clearAll = host.querySelector<HTMLButtonElement>('[data-clear="all"]');
    expect(clearAll?.disabled).toBe(false);
    expect(clearAll?.getAttribute('aria-disabled')).toBe('false');

    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: clearAll });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(sheet.instance.store.getState().ui.pendingFormat).toBeNull();
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')).toBeUndefined();
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');

    tb.dispose();
  });

  it('clears pending Borders formatting through the shared Clear Formats dropdown', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const borders = host.querySelector<HTMLButtonElement>(
      '.demo__ribbon-group--font [data-ribbon-command="borders"]',
    );
    expect(borders).toBeTruthy();
    borders?.click();
    const bottomBorder = host.querySelector<HTMLButtonElement>('[data-border-preset="bottom"]');
    expect(bottomBorder).toBeTruthy();
    bottomBorder?.click();
    expect(sheet.instance.store.getState().ui.pendingFormat).toEqual({
      addr: { sheet: 0, row: 0, col: 0 },
      format: { borders: { bottom: { style: 'thin' } } },
    });
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')).toBeUndefined();

    const clearButton = host.querySelector<HTMLButtonElement>(
      '.demo__ribbon-group--editing [data-ribbon-command="clearFormat"]',
    );
    expect(clearButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'clearFormat', menuId: 'menu-clear' },
      clearButton,
    );
    const clearFormats = host.querySelector<HTMLButtonElement>('[data-clear="formats"]');
    expect(clearFormats?.disabled).toBe(false);

    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: clearFormats });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(sheet.instance.store.getState().ui.pendingFormat).toBeNull();
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')).toBeUndefined();

    tb.dispose();
  });

  it('updates the Clear menu for huge selections by scanning materialized entries only', () => {
    seedText(sheet, 900_000, 0, 'far');
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1048575, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const clearButton = host.querySelector<HTMLButtonElement>(
      '.demo__ribbon-group--editing [data-ribbon-command="clearFormat"]',
    );
    expect(clearButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'clearFormat', menuId: 'menu-clear' },
      clearButton,
    );

    const clearContents = host.querySelector<HTMLButtonElement>('[data-clear="contents"]');
    expect(clearContents?.disabled).toBe(false);
    expect(clearContents?.getAttribute('aria-disabled')).toBe('false');

    tb.dispose();
  });

  it('opens the Fill Series dialog from the Fill dropdown and applies it', async () => {
    seedNumber(sheet, 0, 0, 7);
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const fillButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="fillHome"]');
    expect(fillButton).toBeTruthy();
    fillButton?.click();
    expect(host.querySelectorAll('#menu-fill .app__menu-item--iconic').length).toBe(8);
    const fillDownButton = host.querySelector<HTMLButtonElement>('[data-fill="down"]');
    const fillRightButton = host.querySelector<HTMLButtonElement>('[data-fill="right"]');
    const fillGroupButton = host.querySelector<HTMLButtonElement>('[data-fill="group"]');
    const fillJustifyButton = host.querySelector<HTMLButtonElement>('[data-fill="justify"]');
    const flashFillButton = host.querySelector<HTMLButtonElement>('[data-fill="flash"]');
    expect(fillDownButton?.getAttribute('aria-disabled')).toBe('false');
    expect(fillRightButton?.getAttribute('aria-disabled')).toBe('true');
    expect(fillRightButton?.dataset.menuDisabledReason).toBe(
      'Select more than one column to fill left or right.',
    );
    expect(fillGroupButton?.disabled).toBe(true);
    expect(fillJustifyButton?.disabled).toBe(true);
    expect(flashFillButton).toBeTruthy();
    const seriesButton = host.querySelector<HTMLButtonElement>('[data-fill="series"]');
    expect(seriesButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: seriesButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    await new Promise((resolve) => requestAnimationFrame(resolve));

    const dialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(dialog?.textContent).toContain('Series');
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();

    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'number',
      value: 7,
    });
    expect(sheet.workbook.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({
      kind: 'number',
      value: 7,
    });

    tb.dispose();
  });

  it('applies Fill dropdown directions through the shared default action', () => {
    seedNumber(sheet, 0, 0, 7);
    seedNumber(sheet, 0, 2, 9);
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const fillButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="fillHome"]');
    expect(fillButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'fillHome', menuId: 'menu-fill' },
      fillButton as HTMLButtonElement,
    );
    const fillDownButton = host.querySelector<HTMLButtonElement>('[data-fill="down"]');
    expect(fillDownButton).toBeTruthy();
    expect(fillDownButton?.disabled).toBe(false);
    const downEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(downEvent, 'target', { value: fillDownButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(downEvent)).toBe(true);
    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'number',
      value: 7,
    });
    expect(sheet.workbook.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({
      kind: 'number',
      value: 7,
    });

    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'fillHome', menuId: 'menu-fill' },
      fillButton as HTMLButtonElement,
    );
    const fillLeftButton = host.querySelector<HTMLButtonElement>('[data-fill="left"]');
    expect(fillLeftButton).toBeTruthy();
    expect(fillLeftButton?.disabled).toBe(false);
    expect(fillLeftButton?.dataset.menuDisabledReason).toBeUndefined();
    const leftEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(leftEvent, 'target', { value: fillLeftButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(leftEvent)).toBe(true);
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'number',
      value: 9,
    });
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
      kind: 'number',
      value: 9,
    });

    tb.dispose();
  });

  it('keeps AutoSum presets as icon menu items through the shared dispatcher', () => {
    const dropdowns = dynamicDropdownNoopOverrides();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: dropdowns,
      helpers: stubHelpers(),
    });

    const autosumButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="autosum"]');
    expect(autosumButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'autosum', menuId: 'menu-autosum-home' },
      autosumButton as HTMLButtonElement,
    );
    expect(host.querySelectorAll('#menu-autosum-home [data-autosum-fn]').length).toBe(6);
    expect(
      host.querySelectorAll('#menu-autosum-home .app__menu-icon--svg .app__menu-icon-svg').length,
    ).toBe(1);
    const averageButton = host.querySelector<HTMLButtonElement>('[data-autosum-fn="AVERAGE"]');
    expect(averageButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: averageButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(dropdowns.applyAutoSumFormula).toHaveBeenCalledWith('AVERAGE');

    tb.dispose();
  });

  it('applies AutoSum dropdown presets through the shared default action', () => {
    seedNumber(sheet, 0, 0, 10);
    seedNumber(sheet, 1, 0, 20);
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 2, col: 0 });
    const openFunctionArguments = vi
      .spyOn(sheet.instance, 'openFunctionArguments')
      .mockImplementation(() => undefined);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const autosumButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="autosum"]');
    expect(autosumButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'autosum', menuId: 'menu-autosum-home' },
      autosumButton as HTMLButtonElement,
    );
    const averageButton = host.querySelector<HTMLButtonElement>('[data-autosum-fn="AVERAGE"]');
    const averageEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(averageEvent, 'target', { value: averageButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(averageEvent)).toBe(true);
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 2, col: 0 })).toBe('=AVERAGE(A1:A2)');
    expect(sheet.instance.store.getState().selection.active).toEqual({ sheet: 0, row: 2, col: 0 });

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'autosum', menuId: 'menu-autosum-home' },
      autosumButton as HTMLButtonElement,
    );
    const moreButton = host.querySelector<HTMLButtonElement>('[data-autosum-fn="MORE"]');
    const moreEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(moreEvent, 'target', { value: moreButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(moreEvent)).toBe(true);
    expect(openFunctionArguments).toHaveBeenCalledTimes(1);

    tb.dispose();
  });

  it('keeps Calculation Options radio items icon-backed and dispatchable', () => {
    const dropdowns = dynamicDropdownNoopOverrides();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: dropdowns,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('formulas');

    const calcButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="calcOptions"]');
    expect(calcButton).toBeTruthy();
    calcButton?.click();
    const menu = host.querySelector<HTMLElement>('#menu-calc-options');
    expect(menu?.querySelectorAll('.app__menu-item--iconic').length).toBe(6);
    const radios = Array.from(
      menu?.querySelectorAll<HTMLButtonElement>('[role="menuitemradio"]') ?? [],
    );
    expect(radios.map((button) => button.dataset.calcOption)).toEqual([
      'auto',
      'auto-no-table',
      'manual',
    ]);
    const manual = host.querySelector<HTMLButtonElement>('[data-calc-option="manual"]');
    expect(manual).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: manual });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(dropdowns.applyCalcOptionAction).toHaveBeenCalledWith('manual');

    tb.dispose();
  });

  it('reflects the current calculation mode in the Calculation Options menu', () => {
    vi.spyOn(sheet.workbook, 'calcMode').mockReturnValue(1);
    const setCalcMode = vi.spyOn(sheet.workbook, 'setCalcMode').mockReturnValue(true);
    const recalc = vi.spyOn(sheet.instance, 'recalc').mockImplementation(() => undefined);
    const openIterativeDialog = vi
      .spyOn(sheet.instance, 'openIterativeDialog')
      .mockImplementation(() => undefined);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('formulas');

    const calcButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="calcOptions"]');
    expect(calcButton).toBeTruthy();
    calcButton?.click();

    const auto = host.querySelector<HTMLButtonElement>('[data-calc-option="auto"]');
    const manual = host.querySelector<HTMLButtonElement>('[data-calc-option="manual"]');
    const autoNoTable = host.querySelector<HTMLButtonElement>('[data-calc-option="auto-no-table"]');
    expect(auto?.getAttribute('aria-checked')).toBe('false');
    expect(manual?.getAttribute('aria-checked')).toBe('true');
    expect(manual?.classList.contains('app__menu-item--active')).toBe(true);
    expect(autoNoTable?.getAttribute('aria-checked')).toBe('false');

    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: autoNoTable });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(setCalcMode).toHaveBeenCalledWith(2);

    const calculateNow = host.querySelector<HTMLButtonElement>(
      '[data-calc-option="calculate-now"]',
    );
    const calculateSheet = host.querySelector<HTMLButtonElement>(
      '[data-calc-option="calculate-sheet"]',
    );
    const iterative = host.querySelector<HTMLButtonElement>('[data-calc-option="iterative"]');
    expect(calculateNow).toBeTruthy();
    expect(calculateSheet).toBeTruthy();
    expect(iterative).toBeTruthy();

    const calculateNowEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(calculateNowEvent, 'target', { value: calculateNow });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(calculateNowEvent)).toBe(true);
    const calculateSheetEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(calculateSheetEvent, 'target', { value: calculateSheet });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(calculateSheetEvent)).toBe(true);
    expect(recalc).toHaveBeenCalledTimes(2);

    const iterativeEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(iterativeEvent, 'target', { value: iterative });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(iterativeEvent)).toBe(true);
    expect(openIterativeDialog).toHaveBeenCalledTimes(1);

    tb.dispose();
  });

  it('opens Data Validation from primary click and keeps circle actions secondary', () => {
    seedNumber(sheet, 0, 0, 5);
    seedNumber(sheet, 1, 0, 20);
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 });
    sheet.instance.store.setState((state) => {
      const formats = new Map(state.format.formats);
      formats.set(addrKey({ sheet: 0, row: 0, col: 0 }), {
        validation: { kind: 'whole', op: '<=', a: 10 },
      });
      formats.set(addrKey({ sheet: 0, row: 1, col: 0 }), {
        validation: { kind: 'whole', op: '<=', a: 10 },
      });
      return { ...state, format: { ...state.format, formats } };
    });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('data');

    const validationButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="dataValidation"]',
    );
    expect(validationButton).toBeTruthy();
    validationButton?.click();
    const formatDialog = document.body.querySelector<HTMLElement>('.fc-fmtdlg:not([hidden])');
    expect(formatDialog?.hidden).toBe(false);
    expect(formatDialog?.classList.contains('fc-fmtdlg--data-validation')).toBe(true);
    expect(formatDialog?.querySelector<HTMLElement>('.fc-fmtdlg__title')?.textContent).toBe(
      'Data validation',
    );
    expect(formatDialog?.querySelector<HTMLElement>('.fc-fmtdlg__tabs')?.hidden).toBe(true);
    expect(
      formatDialog?.querySelector<HTMLSelectElement>('select[aria-label="Kind"]'),
    ).toBeTruthy();
    formatDialog
      ?.querySelector<HTMLButtonElement>('.fc-fmtdlg__close')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(formatDialog?.hidden).toBe(true);

    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'dataValidation',
      menuId: 'menu-data-validation',
    });
    expect(host.querySelectorAll('#menu-data-validation .app__menu-item--iconic').length).toBe(4);
    const circleButton = host.querySelector<HTMLButtonElement>(
      '[data-validation-action="circle-invalid"]',
    );
    const initialClearCirclesButton = host.querySelector<HTMLButtonElement>(
      '[data-validation-action="clear-circles"]',
    );
    const initialClearRulesButton = host.querySelector<HTMLButtonElement>(
      '[data-validation-action="clear-rules"]',
    );
    expect(circleButton).toBeTruthy();
    expect(circleButton?.disabled).toBe(false);
    expect(initialClearCirclesButton?.disabled).toBe(true);
    expect(initialClearCirclesButton?.dataset.menuDisabledReason).toBe(
      'There are no validation circles to clear.',
    );
    expect(initialClearRulesButton?.disabled).toBe(false);
    const circleEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(circleEvent, 'target', { value: circleButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(circleEvent)).toBe(true);
    expect(sheet.instance.store.getState().errorIndicators.validationCircles).toEqual(
      new Set([addrKey({ sheet: 0, row: 1, col: 0 })]),
    );

    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'dataValidation',
      menuId: 'menu-data-validation',
    });
    const clearCirclesButton = host.querySelector<HTMLButtonElement>(
      '[data-validation-action="clear-circles"]',
    );
    expect(clearCirclesButton).toBeTruthy();
    expect(clearCirclesButton?.disabled).toBe(false);
    const clearCirclesEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(clearCirclesEvent, 'target', { value: clearCirclesButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(clearCirclesEvent)).toBe(true);
    expect(sheet.instance.store.getState().errorIndicators.validationCircles.size).toBe(0);

    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'dataValidation',
      menuId: 'menu-data-validation',
    });
    expect(clearCirclesButton?.disabled).toBe(true);
    expect(clearCirclesButton?.dataset.menuDisabledReason).toBe(
      'There are no validation circles to clear.',
    );

    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'dataValidation',
      menuId: 'menu-data-validation',
    });
    const clearRulesButton = host.querySelector<HTMLButtonElement>(
      '[data-validation-action="clear-rules"]',
    );
    expect(clearRulesButton).toBeTruthy();
    const clearRulesEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(clearRulesEvent, 'target', { value: clearRulesButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(clearRulesEvent)).toBe(true);
    expect(
      sheet.instance.store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))
        ?.validation,
    ).toBeUndefined();
    expect(
      sheet.instance.store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 0 }))
        ?.validation,
    ).toBeUndefined();

    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'dataValidation',
      menuId: 'menu-data-validation',
    });
    expect(circleButton?.disabled).toBe(true);
    expect(circleButton?.dataset.menuDisabledReason).toBe(
      'The selection does not contain data validation.',
    );
    expect(clearRulesButton?.disabled).toBe(true);
    expect(clearRulesButton?.dataset.menuDisabledReason).toBe(
      'The selection does not contain data validation.',
    );

    tb.dispose();
  });

  it('circles invalid validation cells in huge selections by scanning validation formats only', () => {
    seedNumber(sheet, 900_000, 0, 20);
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1048575, c1: 0 });
    sheet.instance.store.setState((state) => {
      const formats = new Map(state.format.formats);
      formats.set(addrKey({ sheet: 0, row: 900_000, col: 0 }), {
        validation: { kind: 'whole', op: '<=', a: 10 },
      });
      return { ...state, format: { ...state.format, formats } };
    });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'dataValidation',
      menuId: 'menu-data-validation',
    });
    const circleButton = host.querySelector<HTMLButtonElement>(
      '[data-validation-action="circle-invalid"]',
    );
    expect(circleButton?.disabled).toBe(false);

    const circleEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(circleEvent, 'target', { value: circleButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(circleEvent)).toBe(true);
    expect(sheet.instance.store.getState().errorIndicators.validationCircles).toEqual(
      new Set([addrKey({ sheet: 0, row: 900_000, col: 0 })]),
    );

    tb.dispose();
  });

  it('adds prompted conditional formatting rules through the Home dropdown', async () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const conditionalButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="conditional"]',
    );
    expect(conditionalButton).toBeTruthy();
    conditionalButton?.click();
    const greaterThanButton = host.querySelector<HTMLButtonElement>('[data-cf-action="cell-gt"]');
    expect(greaterThanButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: greaterThanButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);

    await Promise.resolve();
    const dialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(dialog?.textContent).toContain('Format cells that are GREATER THAN:');
    expect(dialog?.textContent).toContain('Light Red Fill with Dark Red Text');
    expect(dialog?.textContent).toContain('Yellow Fill with Dark Yellow Text');
    const formatSelect = dialog?.querySelector<HTMLSelectElement>('select.app__dlg__select');
    expect(formatSelect?.value).toBe('light-red-dark-red');
    const formatPreview = dialog?.querySelector<HTMLElement>('[data-conditional-format-preview]');
    expect(formatPreview?.style.background).toBe('#ffc7ce');
    if (!formatSelect) throw new Error('Expected conditional formatting style select.');
    formatSelect.value = 'yellow-dark-yellow';
    formatSelect.dispatchEvent(new Event('change', { bubbles: true }));
    expect(formatPreview?.style.background).toBe('#ffeb9c');
    const input = dialog?.querySelector<HTMLInputElement>('input[type="number"]');
    expect(input).toBeTruthy();
    if (!input) throw new Error('Expected conditional formatting number dialog.');
    input.value = '5';
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    await Promise.resolve();

    expect(sheet.instance.store.getState().conditional.rules).toHaveLength(1);
    expect(sheet.instance.store.getState().conditional.rules[0]).toMatchObject({
      kind: 'cell-value',
      range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 },
      op: '>',
      a: 5,
      apply: { fill: '#ffeb9c', color: '#9c6500' },
    });

    tb.dispose();
  });

  it('opens Conditional Formatting submenus on hover through shared dropdown wiring', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const conditionalButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="conditional"]',
    );
    expect(conditionalButton).toBeTruthy();
    conditionalButton?.click();

    const menu = host.querySelector<HTMLElement>('#menu-conditional');
    const highlight = menu?.querySelector<HTMLElement>('[data-cf-submenu="highlight"]');
    const topBottom = menu?.querySelector<HTMLElement>('[data-cf-submenu="topBottom"]');
    expect(menu?.hidden).toBe(false);
    expect(highlight?.textContent).toContain('Highlight Cells Rules');
    expect(topBottom?.textContent).toContain('Top/Bottom Rules');

    const hoverHighlight = new MouseEvent('mouseover', { bubbles: true });
    Object.defineProperty(hoverHighlight, 'target', { value: highlight });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownHover(hoverHighlight)).toBe(true);
    expect(menu?.querySelector<HTMLElement>('[data-cf-panel="highlight"]')?.hidden).toBe(false);
    expect(highlight?.getAttribute('aria-expanded')).toBe('true');

    const hoverTopBottom = new MouseEvent('mouseover', { bubbles: true });
    Object.defineProperty(hoverTopBottom, 'target', { value: topBottom });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownHover(hoverTopBottom)).toBe(true);
    expect(menu?.querySelector<HTMLElement>('[data-cf-panel="highlight"]')?.hidden).toBe(true);
    expect(highlight?.getAttribute('aria-expanded')).toBe('false');
    expect(menu?.querySelector<HTMLElement>('[data-cf-panel="topBottom"]')?.hidden).toBe(false);
    expect(topBottom?.classList.contains('app__menu-item--active')).toBe(true);
    expect(topBottom?.getAttribute('aria-expanded')).toBe('true');

    tb.dispose();
  });

  it('closes static fallback ribbon menus on outside mousedown', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });

    const conditional = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="conditional"]',
    );
    const find = host.querySelector<HTMLButtonElement>('[data-ribbon-command="findHome"]');
    conditional?.click();
    const conditionalMenu = host.querySelector<HTMLElement>('#menu-conditional');
    expect(conditionalMenu?.hidden).toBe(false);
    expect(conditional?.getAttribute('aria-expanded')).toBe('true');

    find?.click();
    const findMenu = host.querySelector<HTMLElement>('#menu-find-select');
    expect(conditionalMenu?.hidden).toBe(true);
    expect(conditional?.getAttribute('aria-expanded')).toBe('false');
    expect(findMenu?.hidden).toBe(false);
    expect(find?.getAttribute('aria-expanded')).toBe('true');

    document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    expect(findMenu?.hidden).toBe(true);
    expect(find?.getAttribute('aria-expanded')).toBe('false');

    tb.dispose();
  });

  it('closes static fallback ribbon menus on Escape and restores opener focus', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });

    const find = host.querySelector<HTMLButtonElement>('[data-ribbon-command="findHome"]');
    find?.click();
    const findMenu = host.querySelector<HTMLElement>('#menu-find-select');
    expect(findMenu?.hidden).toBe(false);
    expect(find?.getAttribute('aria-expanded')).toBe('true');

    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(findMenu?.hidden).toBe(true);
    expect(find?.getAttribute('aria-expanded')).toBe('false');
    expect(document.activeElement).toBe(find);

    tb.dispose();
  });

  it('closes dynamic ribbon dropdowns on Escape and restores opener focus', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const find = host.querySelector<HTMLButtonElement>('[data-ribbon-command="findHome"]');
    find?.click();
    const findMenu = host.querySelector<HTMLElement>('#menu-find-select');
    expect(findMenu?.hidden).toBe(false);
    expect(find?.getAttribute('aria-expanded')).toBe('true');

    const event = new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true });
    Object.defineProperty(event, 'target', { value: findMenu });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownKeydown(event)).toBe(true);
    expect(findMenu?.hidden).toBe(true);
    expect(find?.getAttribute('aria-expanded')).toBe('false');
    expect(document.activeElement).toBe(find);

    tb.dispose();
  });

  it('closes open dynamic ribbon dropdowns when Escape is handled at document scope', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const find = host.querySelector<HTMLButtonElement>('[data-ribbon-command="findHome"]');
    find?.click();
    const findMenu = host.querySelector<HTMLElement>('#menu-find-select');
    expect(findMenu?.hidden).toBe(false);
    expect(find?.getAttribute('aria-expanded')).toBe('true');

    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(findMenu?.hidden).toBe(true);
    expect(find?.getAttribute('aria-expanded')).toBe('false');
    expect(document.activeElement).toBe(find);

    tb.dispose();
  });

  it('closes open dynamic ribbon dropdowns when focus moves outside the menu and opener', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const find = host.querySelector<HTMLButtonElement>('[data-ribbon-command="findHome"]');
    find?.click();
    const findMenu = host.querySelector<HTMLElement>('#menu-find-select');
    expect(findMenu?.hidden).toBe(false);
    expect(find?.getAttribute('aria-expanded')).toBe('true');

    const firstItem = findMenu?.querySelector<HTMLButtonElement>('button');
    firstItem?.dispatchEvent(new FocusEvent('focusin', { bubbles: true }));
    expect(findMenu?.hidden).toBe(false);

    const outside = document.createElement('button');
    document.body.appendChild(outside);
    outside.dispatchEvent(new FocusEvent('focusin', { bubbles: true }));
    expect(findMenu?.hidden).toBe(true);
    expect(find?.getAttribute('aria-expanded')).toBe('false');

    outside.remove();
    tb.dispose();
  });

  it('opens and closes Conditional Formatting submenus from keyboard events', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const conditionalButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="conditional"]',
    );
    conditionalButton?.click();

    const menu = host.querySelector<HTMLElement>('#menu-conditional');
    const colorScale = menu?.querySelector<HTMLElement>('[data-cf-submenu="colorScale"]');
    const keyOpen = new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      key: 'ArrowRight',
    });
    Object.defineProperty(keyOpen, 'target', { value: colorScale });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownKeydown(keyOpen)).toBe(true);
    const panel = menu?.querySelector<HTMLElement>('[data-cf-panel="colorScale"]');
    expect(panel?.hidden).toBe(false);
    expect(colorScale?.getAttribute('aria-expanded')).toBe('true');
    expect(panel?.querySelector<HTMLButtonElement>('button')?.tabIndex).toBe(0);

    const firstPanelButton = panel?.querySelector<HTMLButtonElement>('button');
    const keyClose = new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      key: 'ArrowLeft',
    });
    Object.defineProperty(keyClose, 'target', { value: firstPanelButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownKeydown(keyClose)).toBe(true);
    expect(panel?.hidden).toBe(true);
    expect(colorScale?.getAttribute('aria-expanded')).toBe('false');
    expect(document.activeElement).toBe(colorScale);

    tb.dispose();
  });

  it('keeps the Conditional Formatting top-level menu order aligned with Excel 365', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="conditional"]')?.click();
    const menu = host.querySelector<HTMLElement>('#menu-conditional');
    const topLevelButtons = Array.from(menu?.children ?? []).filter(
      (child): child is HTMLButtonElement => child instanceof HTMLButtonElement,
    );

    expect(
      topLevelButtons.map((button) => button.dataset.cfSubmenu ?? button.dataset.cfAction),
    ).toEqual([
      'highlight',
      'topBottom',
      'dataBar',
      'colorScale',
      'iconSet',
      'new-rule',
      'clear',
      'manage',
    ]);
    expect(topLevelButtons.map((button) => button.textContent?.replace('▶', '').trim())).toEqual([
      'Highlight Cells Rules',
      'Top/Bottom Rules',
      'Data Bars',
      'Color Scales',
      'Icon Sets',
      'New Rule...',
      'Clear Rules',
      'Manage Rules...',
    ]);

    tb.dispose();
  });

  it('keeps the Conditional Formatting top-level menu icon and submenu affordances', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="conditional"]')?.click();
    const menu = host.querySelector<HTMLElement>('#menu-conditional');
    const topLevelButtons = Array.from(menu?.children ?? []).filter(
      (child): child is HTMLButtonElement => child instanceof HTMLButtonElement,
    );

    expect(topLevelButtons).toHaveLength(8);
    for (const button of topLevelButtons) {
      expect(button.classList.contains('app__menu-item--preset')).toBe(true);
      expect(button.querySelector('.app__cf-icon')).toBeTruthy();
      expect(button.querySelector('.app__menu-item__text')?.textContent?.trim()).not.toBe('');
    }
    const submenuButtons = topLevelButtons.filter((button) => button.dataset.cfSubmenu);
    expect(submenuButtons.map((button) => button.dataset.cfSubmenu)).toEqual([
      'highlight',
      'topBottom',
      'dataBar',
      'colorScale',
      'iconSet',
      'clear',
    ]);
    expect(submenuButtons.every((button) => !!button.querySelector('.app__menu-item__caret'))).toBe(
      true,
    );
    expect(submenuButtons.every((button) => button.getAttribute('aria-haspopup') === 'menu')).toBe(
      true,
    );
    expect(submenuButtons.every((button) => button.getAttribute('aria-expanded') === 'false')).toBe(
      true,
    );
    for (const button of submenuButtons) {
      const panelId = button.getAttribute('aria-controls');
      expect(panelId).toBe(`menu-conditional-${button.dataset.cfSubmenu}`);
      expect(menu?.querySelector<HTMLElement>(`#${panelId}`)?.dataset.cfPanel).toBe(
        button.dataset.cfSubmenu,
      );
    }
    expect(menu?.querySelectorAll<HTMLElement>('.app__submenu--cf')).toHaveLength(6);

    tb.dispose();
  });

  it('keeps Conditional Formatting submenus structured as Excel-style panels', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="conditional"]')?.click();
    const menu = host.querySelector<HTMLElement>('#menu-conditional');
    expect(menu).toBeTruthy();

    const openPanel = (key: string) => {
      const trigger = menu?.querySelector<HTMLElement>(`[data-cf-submenu="${key}"]`);
      const panel = menu?.querySelector<HTMLElement>(`[data-cf-panel="${key}"]`);
      expect(trigger).toBeTruthy();
      expect(panel).toBeTruthy();
      tb.dropdownsApi?.openDynamicConditionalSubmenu(
        menu as HTMLElement,
        key,
        trigger as HTMLElement,
      );
      expect(panel?.hidden).toBe(false);
      expect(panel?.classList.contains('app__submenu--cf')).toBe(true);
      return panel as HTMLElement;
    };

    const highlightPanel = openPanel('highlight');
    expect(
      Array.from(
        highlightPanel.querySelectorAll<HTMLButtonElement>(':scope > .app__menu-item--preset'),
      ).map((button) => button.dataset.cfAction),
    ).toEqual([
      'cell-gt',
      'cell-lt',
      'cell-between',
      'cell-eq',
      'text-contains',
      'date-occurring',
      'duplicates',
      'unique',
      'new-rule',
    ]);
    expect(
      highlightPanel.querySelectorAll(':scope > .app__menu-item--preset .app__cf-icon'),
    ).toHaveLength(9);

    const dataBarPanel = openPanel('dataBar');
    expect(
      Array.from(dataBarPanel.querySelectorAll<HTMLElement>(':scope > .app__menu-heading')).map(
        (heading) => heading.textContent,
      ),
    ).toEqual(['Gradient Fill', 'Solid Fill']);
    expect(
      Array.from(dataBarPanel.querySelectorAll<HTMLElement>(':scope > .app__cf-choice-row')).map(
        (row) => row.querySelectorAll('.app__cf-choice').length,
      ),
    ).toEqual([6, 6]);
    expect(
      dataBarPanel.querySelector<HTMLButtonElement>('[data-cf-action="new-rule"]'),
    ).toBeTruthy();

    const colorScalePanel = openPanel('colorScale');
    expect(colorScalePanel.querySelector('.app__cf-choice-grid-panel')).toBeTruthy();
    expect(colorScalePanel.querySelectorAll('.app__cf-choice')).toHaveLength(12);
    expect(
      Array.from(colorScalePanel.querySelectorAll<HTMLButtonElement>('.app__cf-choice')).every(
        (button) =>
          !!button.getAttribute('aria-label') && !!button.querySelector('.app__cf-choice-grid'),
      ),
    ).toBe(true);

    const iconSetPanel = openPanel('iconSet');
    expect(
      Array.from(iconSetPanel.querySelectorAll<HTMLElement>(':scope > .app__menu-heading')).map(
        (heading) => heading.textContent,
      ),
    ).toEqual(['Directional', 'Shapes', 'Indicators', 'Ratings']);
    expect(iconSetPanel.querySelectorAll(':scope > .app__cf-icon-panel')).toHaveLength(4);
    expect(iconSetPanel.querySelectorAll('.app__cf-icon-choice')).toHaveLength(13);

    const clearPanel = openPanel('clear');
    expect(
      Array.from(clearPanel.querySelectorAll<HTMLButtonElement>(':scope > .app__menu-item')).map(
        (button) => button.dataset.cfAction,
      ),
    ).toEqual(['clear-selection', 'clear-sheet']);
    expect(
      clearPanel.querySelectorAll(':scope > .app__menu-item .app__cf-icon--clear'),
    ).toHaveLength(2);

    tb.dispose();
  });

  it('does not dispatch disabled Conditional Formatting actions', () => {
    const applyConditionalMenuAction = vi.fn();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: () => ({
        ...dynamicDropdownNoopOverrides(),
        applyConditionalMenuAction,
      }),
      helpers: stubHelpers(),
    });

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="conditional"]')?.click();
    const greaterThanButton = host.querySelector<HTMLButtonElement>('[data-cf-action="cell-gt"]');
    expect(greaterThanButton).toBeTruthy();
    if (!greaterThanButton) throw new Error('Expected Conditional Formatting action.');
    greaterThanButton.disabled = true;
    greaterThanButton.setAttribute('aria-disabled', 'true');

    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: greaterThanButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(applyConditionalMenuAction).not.toHaveBeenCalled();
    expect(host.querySelector<HTMLElement>('#menu-conditional')?.hidden).toBe(false);

    tb.dispose();
  });

  it('does not dispatch aria-disabled Conditional Formatting actions', () => {
    const applyConditionalMenuAction = vi.fn();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: () => ({
        ...dynamicDropdownNoopOverrides(),
        applyConditionalMenuAction,
      }),
      helpers: stubHelpers(),
    });

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="conditional"]')?.click();
    const greaterThanButton = host.querySelector<HTMLButtonElement>('[data-cf-action="cell-gt"]');
    expect(greaterThanButton).toBeTruthy();
    if (!greaterThanButton) throw new Error('Expected Conditional Formatting action.');
    greaterThanButton.disabled = false;
    greaterThanButton.setAttribute('aria-disabled', 'true');

    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: greaterThanButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(applyConditionalMenuAction).not.toHaveBeenCalled();

    tb.dispose();
  });

  it('does not open disabled Conditional Formatting submenu triggers', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="conditional"]')?.click();
    const menu = host.querySelector<HTMLElement>('#menu-conditional');
    const trigger = menu?.querySelector<HTMLButtonElement>('[data-cf-submenu="iconSet"]');
    const panel = menu?.querySelector<HTMLElement>('[data-cf-panel="iconSet"]');
    expect(trigger).toBeTruthy();
    expect(panel?.hidden).toBe(true);
    if (!trigger) throw new Error('Expected Conditional Formatting submenu trigger.');
    trigger.disabled = true;
    trigger.setAttribute('aria-disabled', 'true');

    const clickEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(clickEvent, 'target', { value: trigger });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(clickEvent)).toBe(true);
    expect(panel?.hidden).toBe(true);

    const hoverEvent = new MouseEvent('mouseover', { bubbles: true });
    Object.defineProperty(hoverEvent, 'target', { value: trigger });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownHover(hoverEvent)).toBe(true);
    expect(panel?.hidden).toBe(true);

    const keyEvent = new KeyboardEvent('keydown', { bubbles: true, key: 'ArrowRight' });
    Object.defineProperty(keyEvent, 'target', { value: trigger });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownKeydown(keyEvent)).toBe(true);
    expect(panel?.hidden).toBe(true);

    tb.dispose();
  });

  it('keeps every rendered Conditional Formatting action backed by the shared dispatcher', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="conditional"]')?.click();
    const menu = host.querySelector<HTMLElement>('#menu-conditional');
    expect(menu).toBeTruthy();
    const renderedActions = Array.from(
      menu?.querySelectorAll<HTMLElement>('[data-cf-action]') ?? [],
    )
      .map((item) => item.dataset.cfAction)
      .filter((action): action is string => !!action && !action.startsWith('submenu-'));

    expect(renderedActions.length).toBeGreaterThan(0);
    expect([...new Set(renderedActions)].sort()).toEqual(
      [...SUPPORTED_CONDITIONAL_MENU_ACTIONS].sort(),
    );

    tb.dispose();
  });

  it('flips Conditional Formatting submenus left when the right viewport edge is tight', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="conditional"]')?.click();
    const menu = host.querySelector<HTMLElement>('#menu-conditional');
    const trigger = menu?.querySelector<HTMLElement>('[data-cf-submenu="iconSet"]');
    const panel = menu?.querySelector<HTMLElement>('[data-cf-panel="iconSet"]');
    expect(menu).toBeTruthy();
    expect(trigger).toBeTruthy();
    expect(panel).toBeTruthy();
    Object.defineProperty(window, 'innerWidth', { configurable: true, value: 900 });
    if (menu) {
      menu.getBoundingClientRect = vi.fn(
        () => ({ left: 560, right: 780, top: 20, bottom: 320, width: 220, height: 300 }) as DOMRect,
      );
    }
    if (trigger) {
      trigger.getBoundingClientRect = vi.fn(
        () => ({ left: 560, right: 780, top: 120, bottom: 146, width: 220, height: 26 }) as DOMRect,
      );
    }
    if (panel) {
      panel.getBoundingClientRect = vi.fn(
        () => ({ left: 0, right: 0, top: 0, bottom: 0, width: 260, height: 240 }) as DOMRect,
      );
    }

    tb.dropdownsApi?.openDynamicConditionalSubmenu(
      menu as HTMLElement,
      'iconSet',
      trigger as HTMLElement,
    );

    expect(panel?.hidden).toBe(false);
    expect(panel?.style.left).toBe('-259px');
    expect(panel?.style.top).toBe('96px');
    expect(menu?.style.overflowY).toBe('');

    tb.dispose();
  });

  it('clamps Conditional Formatting submenus vertically from the shared submenu opener', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="conditional"]')?.click();
    const menu = host.querySelector<HTMLElement>('#menu-conditional');
    const trigger = menu?.querySelector<HTMLElement>('[data-cf-submenu="iconSet"]');
    const panel = menu?.querySelector<HTMLElement>('[data-cf-panel="iconSet"]');
    expect(menu).toBeTruthy();
    expect(trigger).toBeTruthy();
    expect(panel).toBeTruthy();
    Object.defineProperty(window, 'innerWidth', { configurable: true, value: 1200 });
    Object.defineProperty(window, 'innerHeight', { configurable: true, value: 360 });
    if (menu) {
      menu.getBoundingClientRect = vi.fn(
        () => ({ left: 120, right: 340, top: 20, bottom: 320, width: 220, height: 300 }) as DOMRect,
      );
    }
    if (trigger) {
      trigger.getBoundingClientRect = vi.fn(
        () => ({ left: 120, right: 340, top: 300, bottom: 326, width: 220, height: 26 }) as DOMRect,
      );
    }
    if (panel) {
      panel.getBoundingClientRect = vi.fn(
        () => ({ left: 0, right: 0, top: 0, bottom: 0, width: 260, height: 420 }) as DOMRect,
      );
    }

    tb.dropdownsApi?.openDynamicConditionalSubmenu(
      menu as HTMLElement,
      'iconSet',
      trigger as HTMLElement,
    );

    expect(panel?.hidden).toBe(false);
    expect(panel?.style.left).toBe('219px');
    expect(panel?.style.top).toBe('0px');
    expect(panel?.style.maxHeight).toBe('332px');
    expect(panel?.style.overflowY).toBe('auto');
    expect(panel?.style.overscrollBehavior).toBe('contain');

    tb.dispose();
  });

  it('opens Create Table before applying a Home Format as Table style', async () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 2 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const tableButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="formatTableHome"]',
    );
    expect(tableButton).toBeTruthy();
    tableButton?.click();
    const menu = host.querySelector<HTMLElement>('#menu-table-style-home');
    expect(menu?.classList.contains('app__tablestyle-menu')).toBe(true);
    const scrollBody = menu?.querySelector<HTMLElement>(':scope > .app__tablestyle-scroll');
    expect(scrollBody?.getAttribute('role')).toBe('group');
    expect(scrollBody?.getAttribute('aria-label')).toBe('Format as Table');
    const headings = Array.from(
      scrollBody?.querySelectorAll<HTMLElement>(':scope > .app__tablestyle-heading') ?? [],
    ).map((heading) => heading.textContent);
    expect(headings).toEqual(['Light', 'Medium', 'Dark']);
    const grids = Array.from(
      scrollBody?.querySelectorAll<HTMLElement>(':scope > .app__tablestyle-grid') ?? [],
    );
    expect(grids).toHaveLength(3);
    expect(grids.map((grid) => grid.getAttribute('aria-label'))).toEqual([
      'Light',
      'Medium',
      'Dark',
    ]);
    expect(grids.map((grid) => grid.querySelectorAll('.app__tablestyle-swatch').length)).toEqual([
      28, 28, 7,
    ]);
    const footerActions = Array.from(
      menu?.querySelectorAll<HTMLButtonElement>(':scope > .app__tablestyle-footer') ?? [],
    );
    expect(footerActions.map((button) => button.dataset.tableStyleFooter)).toEqual([
      'new-table-style',
      'new-pivot-style',
    ]);
    expect(
      footerActions.every((button) => button.classList.contains('app__menu-item--iconic')),
    ).toBe(true);
    expect(menu?.querySelector('.app__menu-icon--table-style-new')).toBeTruthy();
    expect(menu?.querySelector('.app__menu-icon--pivot-style-new')).toBeTruthy();
    const styleButton = host.querySelector<HTMLButtonElement>(
      '[data-table-style="dark"][data-table-color="#4472c4"][data-table-variant="banded"]',
    );
    expect(styleButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: styleButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    await Promise.resolve();

    const dialog = document.body.querySelector<HTMLElement>('.app__dlg');
    const rangeInput = dialog?.querySelector<HTMLInputElement>('input[type="text"]');
    expect(document.body.textContent).toContain('Create Table');
    expect(rangeInput?.value).toBe('A1:C4');
    expect(sheet.instance.store.getState().tables.tables).toEqual([]);
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();

    expect(sheet.instance.store.getState().tables.tables).toMatchObject([
      {
        range: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 2 },
        style: 'dark',
        color: '#4472c4',
        banded: true,
        firstCol: false,
      },
    ]);
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'formatTableHome', menuId: 'menu-table-style-home' },
      tableButton,
    );
    const activeStyleButton = host.querySelector<HTMLButtonElement>(
      '[data-table-style="dark"][data-table-color="#4472c4"][data-table-variant="banded"]',
    );
    expect(activeStyleButton?.getAttribute('role')).toBe('menuitemradio');
    expect(activeStyleButton?.getAttribute('aria-checked')).toBe('true');
    expect(activeStyleButton?.classList.contains('app__menu-item--active')).toBe(true);

    tb.dispose();
  });

  it('opens Create Table from the Insert ribbon primary button', async () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 2 });
    seedText(sheet, 0, 0, 'Region');
    seedText(sheet, 0, 1, 'Sales');
    seedText(sheet, 0, 2, 'Qty');
    seedText(sheet, 1, 0, 'East');
    seedNumber(sheet, 1, 1, 10);
    seedNumber(sheet, 1, 2, 2);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('insert');

    const tableButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="formatTableInsert"]',
    );
    expect(tableButton).toBeTruthy();
    tableButton?.click();
    await Promise.resolve();

    expect(document.body.textContent).toContain('Create Table');
    const dialog = document.body.querySelector<HTMLElement>('.app__dlg');
    const rangeInput = dialog?.querySelector<HTMLInputElement>('input[type="text"]');
    expect(rangeInput?.value).toBe('A1:C4');
    await waitFor(() => document.activeElement === rangeInput);
    expect(rangeInput?.closest('.fc-range-picker')).toBeTruthy();
    const rangePicker = dialog?.querySelector<HTMLButtonElement>(
      '[data-range-picker="table-range"]',
    );
    expect(rangePicker).toBeTruthy();
    expect(rangePicker?.getAttribute('aria-label')).toBe('Select range');
    expect(dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.textContent).toBe(
      'OK',
    );
    expect(
      Array.from(dialog?.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__btn') ?? []).some(
        (button) => button.textContent === 'Cancel',
      ),
    ).toBe(true);
    rangePicker?.click();
    expect(rangePicker?.dataset.rangePickerActive).toBe('true');
    expect(rangePicker?.getAttribute('aria-pressed')).toBe('true');
    expect(
      rangeInput?.closest('.fc-range-picker')?.classList.contains('fc-range-picker--picking'),
    ).toBe(true);
    expect(dialog?.closest('.fc-fmtdlg')?.classList.contains('fc-fmtdlg--range-picking')).toBe(
      true,
    );
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 1, r1: 4, c1: 3 });
    expect(rangeInput?.value).toBe('B2:D5');
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(rangePicker?.dataset.rangePickerActive).toBe('false');
    expect(rangePicker?.getAttribute('aria-pressed')).toBe('false');
    expect(dialog?.closest('.fc-fmtdlg')?.classList.contains('fc-fmtdlg--range-picking')).toBe(
      false,
    );
    const headersCheckbox = dialog?.querySelector<HTMLInputElement>('input[type="checkbox"]');
    expect(headersCheckbox?.checked).toBe(true);
    if (headersCheckbox) headersCheckbox.checked = false;
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();

    expect(sheet.instance.store.getState().tables.tables).toMatchObject([
      {
        range: { sheet: 0, r0: 1, c0: 1, r1: 4, c1: 3 },
        style: 'medium',
        showHeader: false,
        banded: true,
      },
    ]);

    tb.dispose();
  });

  it('leaves Create Table headers unchecked for a blank single-cell selection', async () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 4, col: 4 });
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 4, c0: 4, r1: 4, c1: 4 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('insert');

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="formatTableInsert"]')?.click();
    await Promise.resolve();

    const dialog = document.body.querySelector<HTMLElement>('.app__dlg');
    const rangeInput = dialog?.querySelector<HTMLInputElement>('input[type="text"]');
    await waitFor(() => document.activeElement === rangeInput);
    expect(rangeInput?.value).toBe('E5');
    expect(dialog?.querySelector<HTMLInputElement>('input[type="checkbox"]')?.checked).toBe(false);

    tb.dispose();
  });

  it('applies cell styles through the Home Cell Styles dropdown', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 1, col: 1 });
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const stylesButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="cellStyles"]',
    );
    expect(stylesButton).toBeTruthy();
    stylesButton?.click();
    const menu = host.querySelector<HTMLElement>('#menu-cell-styles-home');
    const scrollBody = menu?.querySelector<HTMLElement>(':scope > .app__cellstyle-scroll');
    expect(scrollBody?.getAttribute('role')).toBe('group');
    expect(scrollBody?.getAttribute('aria-label')).toBe('Cell styles');
    expect(
      Array.from(
        scrollBody?.querySelectorAll<HTMLElement>(':scope > .app__cellstyle-heading') ?? [],
      ).map((heading) => heading.textContent),
    ).toEqual([
      'Good, Bad and Neutral',
      'Data and Model',
      'Titles and Headings',
      'Themed Cell Styles',
      'Number Format',
    ]);
    expect(menu?.querySelector<HTMLElement>(':scope > .app__cellstyle-footer')?.parentElement).toBe(
      menu,
    );
    const goodButton = host.querySelector<HTMLButtonElement>('[data-cell-style="good"]');
    expect(goodButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: goodButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);

    expect(
      sheet.instance.store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 1 }))
        ?.cellStyle,
    ).toBe('good');
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'cellStyles', menuId: 'menu-cell-styles-home' },
      stylesButton,
    );
    expect(goodButton?.getAttribute('role')).toBe('menuitemradio');
    expect(goodButton?.getAttribute('aria-checked')).toBe('true');
    expect(goodButton?.classList.contains('app__menu-item--active')).toBe(true);

    tb.dispose();
  });

  it('creates named table and cell styles from the style footer actions', async () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 1, col: 1 });
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 });
    mutators.upsertTableOverlay(sheet.instance.store, {
      id: 'source-table',
      source: 'session',
      range: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 2 },
      style: 'dark',
      color: '#4472c4',
      showHeader: true,
      showTotal: false,
      banded: true,
      firstCol: true,
    });
    mutators.setCellFormat(
      sheet.instance.store,
      { sheet: 0, row: 1, col: 1 },
      {
        bold: true,
        fill: '#c6efce',
        color: '#006100',
      },
    );
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const tableButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="formatTableHome"]',
    );
    expect(tableButton).toBeTruthy();
    tableButton?.click();
    const newTableStyleButton = host.querySelector<HTMLButtonElement>(
      '[data-table-style-footer="new-table-style"]',
    );
    expect(newTableStyleButton).toBeTruthy();
    const tableEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(tableEvent, 'target', { value: newTableStyleButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(tableEvent)).toBe(true);
    await Promise.resolve();
    const tableStyleDialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(tableStyleDialog?.textContent).toContain('New Table Style');
    const tableStyleName = tableStyleDialog?.querySelector<HTMLInputElement>('input');
    expect(tableStyleName).toBeTruthy();
    if (!tableStyleName) throw new Error('Expected new table style name input.');
    tableStyleName.value = 'Review Table';
    const tableStyleType = tableStyleDialog?.querySelector<HTMLSelectElement>(
      '[data-dialog-field="style"]',
    );
    const tableStyleColor = tableStyleDialog?.querySelector<HTMLSelectElement>(
      '[data-dialog-field="color"]',
    );
    const tableStyleBanded = tableStyleDialog?.querySelector<HTMLInputElement>(
      '[data-dialog-field="bandedRows"]',
    );
    const tableStyleFirstCol = tableStyleDialog?.querySelector<HTMLInputElement>(
      '[data-dialog-field="firstColumn"]',
    );
    expect(tableStyleType?.value).toBe('dark');
    expect(tableStyleColor?.value).toBe('#4472c4');
    expect(tableStyleBanded?.checked).toBe(true);
    expect(tableStyleFirstCol?.checked).toBe(true);
    if (!tableStyleType || !tableStyleColor || !tableStyleBanded || !tableStyleFirstCol) {
      throw new Error('Expected table style editor controls.');
    }
    tableStyleType.value = 'light';
    tableStyleColor.value = '#ed7d31';
    tableStyleBanded.checked = false;
    tableStyleFirstCol.checked = true;
    tableStyleDialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();

    expect(sheet.instance.store.getState().tables.customTableStyles).toContainEqual({
      id: customTableStyleId('Review Table'),
      label: 'Review Table',
      style: 'light',
      color: '#ed7d31',
      variant: 'firstCol',
    });

    tb.rerender();
    const refreshedTableButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="formatTableHome"]',
    );
    refreshedTableButton?.click();
    const customTableStyleButton = host.querySelector<HTMLButtonElement>(
      `[data-table-style="${customTableStyleId('Review Table')}"]`,
    );
    expect(customTableStyleButton?.title).toBe('Review Table');

    vi.spyOn(sheet.workbook, 'getPivotTables').mockReturnValue([
      {
        sheetIndex: 0,
        pivotIndex: 2,
        top: 1,
        left: 1,
        rows: 4,
        cols: 3,
        cells: 12,
        fields: ['Region', 'Sales'],
        fieldItems: { Region: ['East'], Sales: ['10'] },
      },
    ]);
    refreshedTableButton?.click();
    const newPivotStyleButton = host.querySelector<HTMLButtonElement>(
      '[data-table-style-footer="new-pivot-style"]',
    );
    expect(newPivotStyleButton).toBeTruthy();
    const pivotEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(pivotEvent, 'target', { value: newPivotStyleButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(pivotEvent)).toBe(true);
    await Promise.resolve();
    const pivotStyleDialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(pivotStyleDialog?.textContent).toContain('New PivotTable Style');
    const pivotStyleName = pivotStyleDialog?.querySelector<HTMLInputElement>('input');
    expect(pivotStyleName).toBeTruthy();
    if (!pivotStyleName) throw new Error('Expected new PivotTable style name input.');
    pivotStyleName.value = 'Review Pivot';
    const pivotStyleType = pivotStyleDialog?.querySelector<HTMLSelectElement>(
      '[data-dialog-field="style"]',
    );
    const pivotStyleColor = pivotStyleDialog?.querySelector<HTMLSelectElement>(
      '[data-dialog-field="color"]',
    );
    expect(pivotStyleType?.value).toBe('dark');
    expect(pivotStyleColor?.value).toBe('#4472c4');
    if (!pivotStyleType || !pivotStyleColor) {
      throw new Error('Expected PivotTable style editor controls.');
    }
    pivotStyleType.value = 'medium';
    pivotStyleColor.value = '#70ad47';
    pivotStyleDialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();

    expect(sheet.instance.store.getState().tables.customPivotTableStyles).toContainEqual({
      id: customPivotTableStyleId('Review Pivot'),
      label: 'Review Pivot',
      style: 'medium',
      color: '#70ad47',
      variant: 'bandedFirstCol',
    });
    expect(sheet.instance.store.getState().tables.pivotTableStyles).toContainEqual({
      sheetIndex: 0,
      pivotIndex: 2,
      styleId: customPivotTableStyleId('Review Pivot'),
    });
    tb.rerender();
    const pivotGalleryButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="formatTableHome"]',
    );
    pivotGalleryButton?.click();
    const customPivotStyleButton = host.querySelector<HTMLButtonElement>(
      `[data-pivot-table-style="${customPivotTableStyleId('Review Pivot')}"]`,
    );
    expect(customPivotStyleButton?.title).toBe('Review Pivot');
    expect(customPivotStyleButton?.getAttribute('role')).toBe('menuitemradio');
    expect(customPivotStyleButton?.getAttribute('aria-checked')).toBe('true');
    expect(customPivotStyleButton?.classList.contains('app__menu-item--active')).toBe(true);

    const stylesButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="cellStyles"]',
    );
    expect(stylesButton).toBeTruthy();
    stylesButton?.click();
    expect(
      host.querySelector('#menu-cell-styles-home .app__menu-icon--cell-style-new'),
    ).toBeTruthy();
    expect(
      host.querySelector('#menu-cell-styles-home .app__menu-icon--cell-style-merge'),
    ).toBeTruthy();
    const newCellStyleButton = host.querySelector<HTMLButtonElement>(
      '[data-cell-style-footer="new-cell-style"]',
    );
    expect(newCellStyleButton).toBeTruthy();
    const cellEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(cellEvent, 'target', { value: newCellStyleButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(cellEvent)).toBe(true);
    await Promise.resolve();
    const cellStyleDialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(cellStyleDialog?.textContent).toContain('New Cell Style');
    const styleName = cellStyleDialog?.querySelector<HTMLInputElement>('input');
    expect(styleName).toBeTruthy();
    if (!styleName) throw new Error('Expected new cell style name input.');
    styleName.value = 'Review OK';
    const includeFill = cellStyleDialog?.querySelector<HTMLInputElement>(
      '[data-dialog-field="fill"]',
    );
    expect(includeFill?.checked).toBe(true);
    if (!includeFill) throw new Error('Expected cell style include fill checkbox.');
    includeFill.checked = false;
    cellStyleDialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();

    expect(
      sheet.instance.store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 1 })),
    ).toMatchObject({
      cellStyle: 'Review OK',
      bold: true,
      fill: '#c6efce',
      color: '#006100',
    });
    const customCellStyle = sheet.instance.store
      .getState()
      .format.customCellStyles?.find((style) => style.id === customCellStyleId('Review OK'));
    expect(customCellStyle?.format).toMatchObject({ bold: true, color: '#006100' });
    expect(customCellStyle?.format.fill).toBeUndefined();
    tb.rerender();
    const refreshedStylesButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="cellStyles"]',
    );
    refreshedStylesButton?.click();
    const customStyleButton = host.querySelector<HTMLButtonElement>(
      `[data-cell-style="${customCellStyleId('Review OK')}"]`,
    );
    expect(customStyleButton?.textContent).toBe('Review OK');

    tb.dispose();
  });

  it('applies currency presets through the Home Currency dropdown', () => {
    seedNumber(sheet, 2, 2, 123);
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 2, col: 2 });
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const currencyButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="currency"]',
    );
    expect(currencyButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'currency', menuId: 'menu-currency-home' },
      currencyButton,
    );
    expect(host.querySelectorAll('#menu-currency-home [data-currency-preset]').length).toBe(5);
    expect(host.querySelector('#menu-currency-home .app__menu-icon--svg')).toBeFalsy();
    expect(host.querySelectorAll('#menu-currency-home .app__menu-item__icon-spacer').length).toBe(
      6,
    );
    const eurButton = host.querySelector<HTMLButtonElement>('[data-currency-preset="€"]');
    expect(eurButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: eurButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);

    expect(
      sheet.instance.store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 2 }))
        ?.numFmt,
    ).toEqual({ kind: 'currency', decimals: 2, symbol: '€' });
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'currency', menuId: 'menu-currency-home' },
      currencyButton,
    );
    const activeEurButton = host.querySelector<HTMLButtonElement>(
      '#menu-currency-home [data-currency-preset="€"]',
    );
    const inactiveUsdButton = host.querySelector<HTMLButtonElement>(
      '#menu-currency-home [data-currency-preset="$"]',
    );
    expect(activeEurButton?.getAttribute('role')).toBe('menuitemradio');
    expect(activeEurButton?.getAttribute('aria-checked')).toBe('true');
    expect(activeEurButton?.classList.contains('app__menu-item--active')).toBe(true);
    expect(inactiveUsdButton?.getAttribute('aria-checked')).toBe('false');

    tb.dispose();
  });

  it('opens custom script from primary click and keeps built-in script actions secondary', async () => {
    seedText(sheet, 0, 0, ' alpha ');
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('automate');

    const scriptButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="script"]');
    expect(scriptButton).toBeTruthy();
    expect(scriptButton?.dataset.ribbonActivation).toBe('splitPrimary');
    scriptButton?.click();
    expect(host.querySelector<HTMLDivElement>('#menu-script')?.hidden).toBe(true);
    const scriptDialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(scriptDialog?.textContent).toContain('Script');
    expect(scriptDialog?.textContent).toContain('Trim whitespace');
    const select = scriptDialog?.querySelector<HTMLSelectElement>('[data-script-command-select]');
    expect(select).toBeTruthy();
    if (!select) throw new Error('Expected script command select.');
    select.value = 'trim';
    scriptDialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    await Promise.resolve();

    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'alpha',
    });
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

    seedText(sheet, 1, 0, ' BETA ');
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 0 });
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'script', menuId: 'menu-script' },
      scriptButton as HTMLButtonElement,
    );
    expect(host.querySelectorAll('#menu-script .app__menu-item--iconic').length).toBe(5);
    const trimButton = host.querySelector<HTMLButtonElement>('[data-script-action="trim"]');
    expect(trimButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: trimButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    await Promise.resolve();

    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'text',
      value: 'BETA',
    });
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      '1 cell(s) changed',
    );
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dispose();
  });

  it('runs PDF and Add-ins primary actions while keeping secondary menus available', async () => {
    const print = vi.spyOn(sheet.instance, 'print').mockImplementation(() => undefined);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('acrobat');

    const pdfButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="pdf"]');
    expect(pdfButton).toBeTruthy();
    expect(pdfButton?.dataset.ribbonActivation).toBe('splitPrimary');
    pdfButton?.click();
    await Promise.resolve();
    expect(host.querySelector<HTMLDivElement>('#menu-pdf')?.hidden).toBe(true);
    expect(print).toHaveBeenCalledWith('pdf');
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'PDF export has been sent',
    );
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'pdf', menuId: 'menu-pdf' },
      pdfButton as HTMLButtonElement,
    );
    expect(host.querySelectorAll('#menu-pdf .app__menu-item--iconic').length).toBe(3);
    const createPdfButton = host.querySelector<HTMLButtonElement>('[data-pdf-action="create"]');
    expect(createPdfButton).toBeTruthy();
    const createPdfEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(createPdfEvent, 'target', { value: createPdfButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(createPdfEvent)).toBe(true);
    await Promise.resolve();
    expect(print).toHaveBeenCalledWith('pdf');
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'PDF export has been sent',
    );
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'pdf', menuId: 'menu-pdf' },
      pdfButton as HTMLButtonElement,
    );
    const shareButton = host.querySelector<HTMLButtonElement>('[data-pdf-action="share"]');
    expect(shareButton).toBeTruthy();
    const shareEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(shareEvent, 'target', { value: shareButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(shareEvent)).toBe(true);
    await Promise.resolve();
    expect(print).toHaveBeenCalledWith('pdf');
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'PDF export is ready',
    );
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'pdf', menuId: 'menu-pdf' },
      pdfButton as HTMLButtonElement,
    );
    const preferencesButton = host.querySelector<HTMLButtonElement>(
      '[data-pdf-action="preferences"]',
    );
    expect(preferencesButton).toBeTruthy();
    const pdfEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(pdfEvent, 'target', { value: preferencesButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(pdfEvent)).toBe(true);
    expect(document.body.querySelector<HTMLElement>('.fc-pgsetup')?.hidden).toBe(false);
    document.body
      .querySelector<HTMLButtonElement>('.fc-pgsetup .fc-fmtdlg__close')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const addInButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="addIn"]');
    expect(addInButton).toBeTruthy();
    expect(addInButton?.dataset.ribbonActivation).toBe('splitPrimary');
    addInButton?.click();
    await Promise.resolve();
    expect(host.querySelector<HTMLDivElement>('#menu-add-ins')?.hidden).toBe(true);
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'Add-in management',
    );
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'addIn', menuId: 'menu-add-ins' },
      addInButton as HTMLButtonElement,
    );
    expect(host.querySelectorAll('#menu-add-ins .app__menu-item--iconic').length).toBe(3);
    const getButton = host.querySelector<HTMLButtonElement>('[data-add-in-action="get"]');
    expect(getButton).toBeTruthy();
    const getAddInEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(getAddInEvent, 'target', { value: getButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(getAddInEvent)).toBe(true);
    await Promise.resolve();
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'Get Add-ins',
    );
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'Office Add-ins',
    );
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'addIn', menuId: 'menu-add-ins' },
      addInButton as HTMLButtonElement,
    );
    const myButton = host.querySelector<HTMLButtonElement>('[data-add-in-action="my"]');
    expect(myButton).toBeTruthy();
    const addInEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(addInEvent, 'target', { value: myButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(addInEvent)).toBe(true);
    await Promise.resolve();
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'External add-ins',
    );
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'addIn', menuId: 'menu-add-ins' },
      addInButton as HTMLButtonElement,
    );
    const manageButton = host.querySelector<HTMLButtonElement>('[data-add-in-action="manage"]');
    expect(manageButton).toBeTruthy();
    const manageEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(manageEvent, 'target', { value: manageButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(manageEvent)).toBe(true);
    await Promise.resolve();
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'Add-in management',
    );
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dispose();
    print.mockRestore();
  });

  it('surfaces Insert media reports and creates session shapes', async () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('insert');

    const pictureButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="pictureInsert"]',
    );
    expect(pictureButton).toBeTruthy();
    pictureButton?.click();
    const pictureMenu = host.querySelector<HTMLElement>('#menu-picture-insert');
    expect(pictureMenu?.classList.contains('app__menu--visual')).toBe(true);
    expect(pictureMenu?.querySelectorAll('.app__visual-tile')).toHaveLength(3);
    const stockButton = host.querySelector<HTMLButtonElement>('[data-picture-insert="stock"]');
    expect(stockButton).toBeTruthy();
    const stockEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(stockEvent, 'target', { value: stockButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(stockEvent)).toBe(true);
    await Promise.resolve();
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'Stock Images',
    );
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'host-provided media picker',
    );
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

    pictureButton?.click();
    const onlineButton = host.querySelector<HTMLButtonElement>('[data-picture-insert="online"]');
    expect(onlineButton).toBeTruthy();
    const pictureEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(pictureEvent, 'target', { value: onlineButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(pictureEvent)).toBe(true);
    await Promise.resolve();
    const pictureDialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(pictureDialog?.textContent).toContain('Online Pictures');
    expect(pictureDialog?.textContent).toContain('host-provided media picker');
    expect(pictureDialog?.querySelector('input')).toBeNull();
    pictureDialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    expect(sheet.instance.store.getState().illustrations.illustrations).toEqual([]);

    const toDataUrl = vi
      .spyOn(HTMLCanvasElement.prototype, 'toDataURL')
      .mockReturnValue('data:image/png;base64,current-view');
    const screenshotButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="screenshotInsert"]',
    );
    expect(screenshotButton).toBeTruthy();
    screenshotButton?.click();
    const screenshotMenu = host.querySelector<HTMLElement>('#menu-screenshot-insert');
    expect(screenshotMenu?.classList.contains('app__menu--visual')).toBe(true);
    expect(screenshotMenu?.querySelectorAll('.app__visual-tile')).toHaveLength(2);
    expect(screenshotMenu?.querySelector('.app__menu-heading')?.textContent).toBe(
      'Available Windows',
    );
    const currentViewButton = host.querySelector<HTMLButtonElement>(
      '[data-screenshot-insert="current-view"]',
    );
    expect(currentViewButton).toBeTruthy();
    expect(currentViewButton?.classList.contains('app__visual-tile--screenshot-preview')).toBe(
      true,
    );
    const screenshotEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(screenshotEvent, 'target', { value: currentViewButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(screenshotEvent)).toBe(true);
    expect(sheet.instance.store.getState().illustrations.illustrations).toMatchObject([
      {
        kind: 'image',
        src: 'data:image/png;base64,current-view',
        sheet: 0,
      },
    ]);
    toDataUrl.mockRestore();

    const OriginalFileReader = globalThis.FileReader;
    vi.stubGlobal(
      'FileReader',
      class {
        result: string | ArrayBuffer | null = null;
        private listeners = new Map<string, EventListener[]>();
        addEventListener(type: string, listener: EventListener): void {
          this.listeners.set(type, [...(this.listeners.get(type) ?? []), listener]);
        }
        readAsDataURL(file: File): void {
          this.result = `data:${file.type};base64,from-device`;
          for (const listener of this.listeners.get('load') ?? []) {
            listener.call(this, new Event('load'));
          }
        }
      } as unknown as typeof FileReader,
    );
    const inputClick = vi.spyOn(HTMLInputElement.prototype, 'click').mockImplementation(() => {});
    const createElement = vi.spyOn(document, 'createElement');
    const originalCreateElement = createElement.getMockImplementation();
    let fileInput: HTMLInputElement | null = null;
    createElement.mockImplementation(((tagName: string, options?: ElementCreationOptions) => {
      const el = originalCreateElement
        ? originalCreateElement.call(document, tagName, options)
        : Document.prototype.createElement.call(document, tagName, options);
      if (tagName.toLowerCase() === 'input') fileInput = el as HTMLInputElement;
      return el;
    }) as typeof document.createElement);
    pictureButton?.click();
    const deviceButton = host.querySelector<HTMLButtonElement>('[data-picture-insert="device"]');
    expect(deviceButton).toBeTruthy();
    const deviceEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(deviceEvent, 'target', { value: deviceButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(deviceEvent)).toBe(true);
    await Promise.resolve();
    expect(inputClick).toHaveBeenCalled();
    expect(fileInput).toBeTruthy();
    const selectedFileInput = fileInput as HTMLInputElement | null;
    if (!selectedFileInput) throw new Error('Expected device image file input.');
    Object.defineProperty(selectedFileInput, 'files', {
      value: [new File(['device'], 'device.png', { type: 'image/png' })],
      configurable: true,
    });
    selectedFileInput.dispatchEvent(new Event('change'));
    await Promise.resolve();
    expect(sheet.instance.store.getState().illustrations.illustrations).toMatchObject([
      {
        kind: 'image',
        src: 'data:image/png;base64,current-view',
      },
      {
        kind: 'image',
        src: 'data:image/png;base64,from-device',
        alt: 'device.png',
        sheet: 0,
      },
    ]);
    createElement.mockRestore();
    inputClick.mockRestore();
    vi.stubGlobal('FileReader', OriginalFileReader);

    const shapeButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="shapesInsert"]',
    );
    expect(shapeButton).toBeTruthy();
    shapeButton?.click();
    const shapesMenu = host.querySelector<HTMLElement>('#menu-shapes-insert');
    expect(shapesMenu?.classList.contains('app__menu--visual')).toBe(true);
    expect(shapesMenu?.querySelectorAll('.app__visual-tile')).toHaveLength(7);
    expect(
      Array.from(shapesMenu?.querySelectorAll<HTMLElement>('.app__menu-heading') ?? []).map(
        (heading) => heading.textContent,
      ),
    ).toEqual(['Lines', 'Rectangles', 'Basic Shapes']);
    const arrowButton = host.querySelector<HTMLButtonElement>('[data-shape-insert="arrow"]');
    const diamondButton = host.querySelector<HTMLButtonElement>('[data-shape-insert="diamond"]');
    expect(arrowButton).toBeTruthy();
    expect(diamondButton).toBeTruthy();
    const shapeEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(shapeEvent, 'target', { value: diamondButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(shapeEvent)).toBe(true);
    expect(sheet.instance.store.getState().illustrations.illustrations).toMatchObject([
      {
        kind: 'image',
        src: 'data:image/png;base64,current-view',
      },
      {
        kind: 'image',
        src: 'data:image/png;base64,from-device',
      },
      {
        kind: 'shape',
        shape: 'diamond',
        sheet: 0,
        w: 160,
        h: 96,
      },
    ]);

    tb.dispose();
  });

  it('inserts host-provided screen clippings from the Screenshot menu', async () => {
    const screenSheet = await mountStubSheet({
      locale: 'en',
      captureScreenClip: () => ({
        src: 'data:image/png;base64,screen-clip',
        alt: 'Screen clipping',
      }),
    });
    const screenHost = document.createElement('div');
    document.body.appendChild(screenHost);
    try {
      const tb = Spreadsheet.mountToolbar(screenHost, screenSheet.instance, {
        dynamicDropdowns: true,
        helpers: stubHelpers(),
      });
      tb.setActiveTab('insert');
      const screenshotButton = screenHost.querySelector<HTMLButtonElement>(
        '[data-ribbon-command="screenshotInsert"]',
      );
      expect(screenshotButton).toBeTruthy();
      screenshotButton?.click();
      const screenClippingButton = screenHost.querySelector<HTMLButtonElement>(
        '[data-screenshot-insert="screen-clipping"]',
      );
      expect(screenClippingButton).toBeTruthy();
      const event = new MouseEvent('click', { bubbles: true });
      Object.defineProperty(event, 'target', { value: screenClippingButton });
      expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
      await Promise.resolve();
      await Promise.resolve();
      expect(screenSheet.instance.store.getState().illustrations.illustrations).toMatchObject([
        {
          alt: 'Screen clipping',
          kind: 'image',
          sheet: 0,
          src: 'data:image/png;base64,screen-clip',
          w: 240,
          h: 160,
        },
      ]);
      tb.dispose();
    } finally {
      screenSheet.dispose();
      screenHost.remove();
    }
  });

  it('reports Screen Clipping as host-provided when no capture hook is available', async () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('insert');
    const screenshotButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="screenshotInsert"]',
    );
    expect(screenshotButton).toBeTruthy();
    screenshotButton?.click();
    const screenClippingButton = host.querySelector<HTMLButtonElement>(
      '[data-screenshot-insert="screen-clipping"]',
    );
    expect(screenClippingButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: screenClippingButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    await Promise.resolve();
    await Promise.resolve();

    const dialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(dialog?.textContent).toContain('Screen Clipping');
    expect(dialog?.textContent).toContain('captureScreenClip');
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();

    tb.dispose();
  });

  it('reports Current Sheet View screenshot export failures with a screenshot-specific detail', async () => {
    const toDataUrl = vi.spyOn(HTMLCanvasElement.prototype, 'toDataURL').mockReturnValue('');
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('insert');
    const screenshotButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="screenshotInsert"]',
    );
    expect(screenshotButton).toBeTruthy();
    screenshotButton?.click();
    const currentViewButton = host.querySelector<HTMLButtonElement>(
      '[data-screenshot-insert="current-view"]',
    );
    expect(currentViewButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: currentViewButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    await Promise.resolve();
    await Promise.resolve();

    const dialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(dialog?.textContent).toContain('Current Sheet View');
    expect(dialog?.textContent).toContain('mounted grid canvas');
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();

    toDataUrl.mockRestore();
    tb.dispose();
  });

  it('disables Page Layout Arrange ordering actions until an illustration exists', () => {
    const openWorkbookObjects = vi.spyOn(sheet.instance, 'openWorkbookObjects');
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('pageLayout');

    const arrangeButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="arrangeObjectsPageLayout"]',
    );
    expect(arrangeButton).toBeTruthy();
    arrangeButton?.click();
    const bringForward = host.querySelector<HTMLButtonElement>(
      '#menu-arrange-objects [data-arrange-action="bring-forward"]',
    );
    const selectionPane = host.querySelector<HTMLButtonElement>(
      '#menu-arrange-objects [data-arrange-action="selection-pane"]',
    );
    expect(bringForward?.disabled).toBe(true);
    expect(bringForward?.getAttribute('aria-disabled')).toBe('true');
    expect(bringForward?.dataset.menuDisabledReason).toBe('Select an object to arrange.');
    expect(selectionPane?.disabled).toBe(false);

    const disabledEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(disabledEvent, 'target', { value: bringForward });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(disabledEvent)).toBe(true);
    expect(openWorkbookObjects).not.toHaveBeenCalled();

    if (!bringForward) throw new Error('Expected Bring Forward item.');
    bringForward.disabled = false;
    bringForward.setAttribute('aria-disabled', 'true');
    const ariaDisabledEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(ariaDisabledEvent, 'target', { value: bringForward });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(ariaDisabledEvent)).toBe(true);
    expect(openWorkbookObjects).not.toHaveBeenCalled();

    mutators.upsertIllustration(sheet.instance.store, {
      id: 'shape-back',
      kind: 'shape',
      shape: 'rectangle',
      sheet: 0,
    });
    mutators.upsertIllustration(sheet.instance.store, {
      id: 'shape-front',
      kind: 'shape',
      shape: 'arrow',
      sheet: 0,
    });

    tb.setActiveTab('pageLayout');
    host
      .querySelector<HTMLButtonElement>('[data-ribbon-command="arrangeObjectsPageLayout"]')
      ?.click();
    const frontBringForward = host.querySelector<HTMLButtonElement>(
      '#menu-arrange-objects [data-arrange-action="bring-forward"]',
    );
    const frontSendBackward = host.querySelector<HTMLButtonElement>(
      '#menu-arrange-objects [data-arrange-action="send-backward"]',
    );
    expect(frontBringForward?.disabled).toBe(true);
    expect(frontBringForward?.dataset.menuDisabledReason).toBe(
      'The selected object is already at the front.',
    );
    expect(frontSendBackward?.disabled).toBe(false);

    const selectedBack = document.createElement('div');
    selectedBack.className = 'fc-illustration';
    selectedBack.dataset.illustrationId = 'shape-back';
    selectedBack.setAttribute('aria-selected', 'true');
    sheet.instance.host.appendChild(selectedBack);
    host
      .querySelector<HTMLButtonElement>('[data-ribbon-command="arrangeObjectsPageLayout"]')
      ?.click();
    const backBringForward = host.querySelector<HTMLButtonElement>(
      '#menu-arrange-objects [data-arrange-action="bring-forward"]',
    );
    const backSendBackward = host.querySelector<HTMLButtonElement>(
      '#menu-arrange-objects [data-arrange-action="send-backward"]',
    );
    expect(backBringForward?.disabled).toBe(false);
    expect(backBringForward?.getAttribute('aria-disabled')).toBe('false');
    expect(backBringForward?.dataset.menuDisabledReason).toBeUndefined();
    expect(backSendBackward?.disabled).toBe(true);
    expect(backSendBackward?.getAttribute('aria-disabled')).toBe('true');
    expect(backSendBackward?.dataset.menuDisabledReason).toBe(
      'The selected object is already at the back.',
    );

    tb.dispose();
    openWorkbookObjects.mockRestore();
  });

  it('opens Protect Sheet from primary click and keeps protection options secondary', async () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('review');

    const protectButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="protectReview"]',
    );
    expect(protectButton).toBeTruthy();
    expect(protectButton?.dataset.ribbonActivation).toBe('splitPrimary');
    protectButton?.click();
    expect(host.querySelector<HTMLDivElement>('#menu-protect-review')?.hidden).toBe(true);
    const protectDialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(protectDialog?.textContent).toContain('Protect Sheet');
    expect(protectDialog?.textContent).toContain('Allow all users of this worksheet to:');
    for (const label of [
      'Format columns',
      'Format rows',
      'Insert columns',
      'Insert hyperlinks',
      'Delete columns',
      'Delete rows',
      'Use PivotTable reports',
      'Edit objects',
      'Edit scenarios',
    ]) {
      expect(protectDialog?.textContent).toContain(label);
    }
    const protectInputs = Array.from(
      protectDialog?.querySelectorAll<HTMLInputElement>('input') ?? [],
    );
    const passwordInput = protectInputs[0];
    const confirmInput = protectInputs[1];
    expect(passwordInput).toBeTruthy();
    expect(confirmInput).toBeTruthy();
    if (!passwordInput) throw new Error('Expected Protect Sheet password input.');
    if (!confirmInput) throw new Error('Expected Protect Sheet confirmation input.');
    passwordInput.value = 'pw';
    confirmInput.value = 'pw';
    const formatCellsOption = protectInputs.find((input) =>
      input.closest('label')?.textContent?.includes('Format cells'),
    );
    expect(formatCellsOption).toBeTruthy();
    if (!formatCellsOption) throw new Error('Expected Format cells permission option.');
    formatCellsOption.checked = true;
    protectDialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    expect(sheet.instance.store.getState().protection.protectedSheets.has(0)).toBe(true);
    expect(sheet.instance.store.getState().protection.protectedSheets.get(0)?.password).toBe('pw');
    expect(
      sheet.instance.store.getState().protection.protectedSheets.get(0)?.permissions,
    ).toMatchObject({
      formatCells: true,
      formatColumns: false,
      insertColumns: false,
      insertHyperlinks: false,
      deleteRows: false,
      selectLockedCells: true,
      selectUnlockedCells: true,
      pivotTables: false,
      objects: false,
      scenarios: false,
    });

    protectButton?.click();
    const unprotectDialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(unprotectDialog?.textContent).toContain('Unprotect Sheet');
    const unprotectInput = unprotectDialog?.querySelector<HTMLInputElement>('input');
    expect(unprotectInput).toBeTruthy();
    if (!unprotectInput) throw new Error('Expected Unprotect Sheet password input.');
    unprotectInput.value = 'bad';
    unprotectDialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'The password is incorrect.',
    );
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    expect(sheet.instance.store.getState().protection.protectedSheets.has(0)).toBe(true);

    protectButton?.click();
    const unprotectRetryDialog = document.body.querySelector<HTMLElement>('.app__dlg');
    const unprotectRetryInput = unprotectRetryDialog?.querySelector<HTMLInputElement>('input');
    expect(unprotectRetryInput).toBeTruthy();
    if (!unprotectRetryInput) throw new Error('Expected retry password input.');
    unprotectRetryInput.value = 'pw';
    unprotectRetryDialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    expect(sheet.instance.store.getState().protection.protectedSheets.has(0)).toBe(false);

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'protectReview', menuId: 'menu-protect-review' },
      protectButton as HTMLButtonElement,
    );
    const reviewProtectMenu = host.querySelector<HTMLElement>('#menu-protect-review');
    expect(reviewProtectMenu?.querySelectorAll('.app__menu-item--iconic').length).toBe(8);
    expect(
      reviewProtectMenu?.querySelector<HTMLButtonElement>('[data-protect-action="protect-sheet"]')
        ?.disabled,
    ).toBe(false);
    expect(
      reviewProtectMenu?.querySelector<HTMLButtonElement>('[data-protect-action="unprotect-sheet"]')
        ?.disabled,
    ).toBe(true);
    const workbookButton = reviewProtectMenu?.querySelector<HTMLButtonElement>(
      '[data-protect-action="protect-workbook"]',
    );
    const unprotectWorkbookButton = reviewProtectMenu?.querySelector<HTMLButtonElement>(
      '[data-protect-action="unprotect-workbook"]',
    );
    const clearAllowedRangesButton = reviewProtectMenu?.querySelector<HTMLButtonElement>(
      '[data-protect-action="clear-allowed-edit-ranges"]',
    );
    expect(workbookButton).toBeTruthy();
    expect(workbookButton?.disabled).toBe(false);
    expect(unprotectWorkbookButton?.disabled).toBe(true);
    expect(unprotectWorkbookButton?.dataset.menuDisabledReason).toBe(
      'Workbook structure is not protected.',
    );
    expect(clearAllowedRangesButton?.disabled).toBe(true);
    expect(clearAllowedRangesButton?.dataset.menuDisabledReason).toBe(
      'There are no allowed edit ranges to clear.',
    );
    const workbookEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(workbookEvent, 'target', { value: workbookButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(workbookEvent)).toBe(true);
    expect(sheet.instance.store.getState().protection.workbookStructure).toEqual({});

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'protectReview', menuId: 'menu-protect-review' },
      protectButton as HTMLButtonElement,
    );
    expect(workbookButton?.disabled).toBe(true);
    expect(workbookButton?.dataset.menuDisabledReason).toBe(
      'Workbook structure is already protected.',
    );
    expect(unprotectWorkbookButton?.disabled).toBe(false);
    expect(unprotectWorkbookButton?.dataset.menuDisabledReason).toBeUndefined();
    const allowRangeButton = reviewProtectMenu?.querySelector<HTMLButtonElement>(
      '[data-protect-action="allow-edit-ranges"]',
    );
    expect(allowRangeButton).toBeTruthy();
    const allowEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(allowEvent, 'target', { value: allowRangeButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(allowEvent)).toBe(true);
    await waitFor(() => Boolean(document.body.querySelector<HTMLElement>('.app__dlg')));
    const allowDialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(allowDialog?.textContent).toContain('Range');
    const allowRangeInput = allowDialog?.querySelector<HTMLInputElement>('input');
    expect(allowRangeInput?.value).toBeTruthy();
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();
    await waitFor(() => sheet.instance.store.getState().protection.allowedEditRanges.length === 1);
    await waitFor(() =>
      Boolean(
        document.body.querySelector<HTMLElement>('.app__dlg')?.textContent?.includes('Allowed'),
      ),
    );
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'protectReview', menuId: 'menu-protect-review' },
      protectButton as HTMLButtonElement,
    );
    expect(clearAllowedRangesButton?.disabled).toBe(false);
    expect(clearAllowedRangesButton?.dataset.menuDisabledReason).toBeUndefined();
    const clearAllowedEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(clearAllowedEvent, 'target', { value: clearAllowedRangesButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(clearAllowedEvent)).toBe(true);
    expect(sheet.instance.store.getState().protection.allowedEditRanges).toEqual([]);
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'Cleared allowed edit ranges',
    );
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dispose();
  });

  it('opens Protect Sheet from the View Protect primary button', async () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('view');

    const protectButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="protect"]');
    expect(protectButton).toBeTruthy();
    expect(protectButton?.dataset.ribbonActivation).toBe('splitPrimary');
    protectButton?.click();

    expect(host.querySelector<HTMLDivElement>('#menu-protect-view')?.hidden).toBe(true);
    const protectDialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(protectDialog?.textContent).toContain('Protect Sheet');
    protectDialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    expect(sheet.instance.store.getState().protection.protectedSheets.has(0)).toBe(true);

    tb.dispose();
  });

  it('unprotects sheets protected with OOXML strong password hash metadata', async () => {
    mutators.setSheetProtected(sheet.instance.store, 0, true, {
      passwordHash: {
        algorithmName: 'SHA-512',
        hashValue:
          'auSsy34WkQnOLam+2c6zPtbZkU+N88uiaOoz1sI5q58RL3NaIY7C18AbhRNGb15+UxRtDZCp10kT11782Plh7A==',
        saltValue: 'Furur6jnDIFaQBhHQBXzFA==',
        spinCount: 0,
      },
    });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('review');

    const protectButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="protectReview"]',
    );
    expect(protectButton).toBeTruthy();
    protectButton?.click();

    let unprotectDialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(unprotectDialog?.textContent).toContain('Unprotect Sheet');
    let input = unprotectDialog?.querySelector<HTMLInputElement>('input');
    expect(input).toBeTruthy();
    if (!input) throw new Error('Expected Unprotect Sheet password input.');
    input.value = 'wrong';
    unprotectDialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await waitFor(
      () =>
        document.body
          .querySelector<HTMLElement>('.app__dlg')
          ?.textContent?.includes('The password is incorrect.') === true,
    );
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'The password is incorrect.',
    );
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();
    expect(sheet.instance.store.getState().protection.protectedSheets.has(0)).toBe(true);

    protectButton?.click();
    unprotectDialog = document.body.querySelector<HTMLElement>('.app__dlg');
    input = unprotectDialog?.querySelector<HTMLInputElement>('input');
    expect(input).toBeTruthy();
    if (!input) throw new Error('Expected retry password input.');
    input.value = 'password';
    unprotectDialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await waitFor(() => !sheet.instance.store.getState().protection.protectedSheets.has(0));

    tb.dispose();
  });

  it('keeps View Protect secondary menu scoped to sheet protection actions', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('view');

    const protectButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="protect"]');
    expect(protectButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'protect', menuId: 'menu-protect-view' },
      protectButton as HTMLButtonElement,
    );
    const menu = host.querySelector<HTMLDivElement>('#menu-protect-view');
    expect(menu?.querySelectorAll('.app__menu-item--iconic').length).toBe(2);
    const protectSheet = menu?.querySelector<HTMLButtonElement>(
      '[data-protect-action="protect-sheet"]',
    );
    const unprotectSheet = menu?.querySelector<HTMLButtonElement>(
      '[data-protect-action="unprotect-sheet"]',
    );
    expect(protectSheet).toBeTruthy();
    expect(unprotectSheet).toBeTruthy();
    expect(protectSheet?.disabled).toBe(false);
    expect(unprotectSheet?.disabled).toBe(true);
    expect(menu?.querySelector('[data-protect-action="protect-workbook"]')).toBeNull();
    expect(menu?.querySelector('[data-protect-action="allow-edit-ranges"]')).toBeNull();
    expect(menu?.querySelector('[data-protect-action="lock-cell"]')).toBeNull();

    mutators.setSheetProtected(sheet.instance.store, 0, true);
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'protect', menuId: 'menu-protect-view' },
      protectButton as HTMLButtonElement,
    );
    expect(protectSheet?.disabled).toBe(true);
    expect(unprotectSheet?.disabled).toBe(false);

    tb.dispose();
  });

  it('deletes the active comment from primary click and keeps delete-all secondary', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    setComment(sheet.instance.store, { sheet: 0, row: 0, col: 0 }, 'note', sheet.workbook);
    setComment(sheet.instance.store, { sheet: 0, row: 1, col: 0 }, 'other note', sheet.workbook);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('review');

    const commentsButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="deleteCommentReview"]',
    );
    expect(commentsButton).toBeTruthy();
    expect(commentsButton?.dataset.ribbonActivation).toBe('splitPrimary');
    commentsButton?.click();

    expect(host.querySelector<HTMLDivElement>('#menu-review-comments')?.hidden).toBe(true);
    expect(commentAt(sheet.instance.store.getState(), { sheet: 0, row: 0, col: 0 })).toBeNull();
    expect(commentAt(sheet.instance.store.getState(), { sheet: 0, row: 1, col: 0 })).toBe(
      'other note',
    );

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'deleteCommentReview', menuId: 'menu-review-comments' },
      commentsButton as HTMLButtonElement,
    );
    expect(host.querySelectorAll('#menu-review-comments .app__menu-item--iconic').length).toBe(2);
    const deleteActiveButton = host.querySelector<HTMLButtonElement>(
      '[data-comment-action="delete-active"]',
    );
    expect(deleteActiveButton?.disabled).toBe(true);
    const deleteAllButton = host.querySelector<HTMLButtonElement>(
      '[data-comment-action="delete-all"]',
    );
    expect(deleteAllButton).toBeTruthy();
    expect(deleteAllButton?.disabled).toBe(false);
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: deleteAllButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);

    expect(commentAt(sheet.instance.store.getState(), { sheet: 0, row: 1, col: 0 })).toBeNull();

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'deleteCommentReview', menuId: 'menu-review-comments' },
      commentsButton as HTMLButtonElement,
    );
    expect(deleteActiveButton?.disabled).toBe(true);
    expect(deleteAllButton?.disabled).toBe(true);

    setComment(sheet.instance.store, { sheet: 0, row: 0, col: 0 }, 'fresh note', sheet.workbook);
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'deleteCommentReview', menuId: 'menu-review-comments' },
      commentsButton as HTMLButtonElement,
    );
    expect(deleteActiveButton?.disabled).toBe(false);
    expect(deleteAllButton?.disabled).toBe(false);
    const deleteActiveEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(deleteActiveEvent, 'target', { value: deleteActiveButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(deleteActiveEvent)).toBe(true);
    expect(commentAt(sheet.instance.store.getState(), { sheet: 0, row: 0, col: 0 })).toBeNull();

    tb.dispose();
  });

  it('opens Recommended Charts report from primary click and keeps chart types secondary', async () => {
    seedNumber(sheet, 0, 0, 1);
    seedNumber(sheet, 1, 0, 2);
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('insert');

    const chartButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="chartInsert"]',
    );
    expect(chartButton).toBeTruthy();
    chartButton?.click();
    await Promise.resolve();
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'Recommended Charts',
    );
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'Persisted chart creation and editing',
    );
    expect(sheet.instance.store.getState().charts.charts).toEqual([]);
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 });
    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'chartInsert',
      menuId: 'menu-chart-insert',
    });
    const chartMenu = host.querySelector<HTMLDivElement>('#menu-chart-insert');
    expect(chartMenu?.classList.contains('app__menu--visual')).toBe(true);
    expect(chartMenu?.querySelectorAll('.app__visual-tile[data-chart-insert]').length).toBe(7);
    const recommendedButton = host.querySelector<HTMLButtonElement>(
      '[data-chart-insert="recommended"]',
    );
    expect(recommendedButton).toBeTruthy();
    const recommendedEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(recommendedEvent, 'target', { value: recommendedButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(recommendedEvent)).toBe(true);
    await Promise.resolve();
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'Recommended Charts',
    );
    expect(sheet.instance.store.getState().charts.charts).toEqual([]);
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'chartInsert',
      menuId: 'menu-chart-insert',
    });
    const barButton = host.querySelector<HTMLButtonElement>('[data-chart-insert="bar"]');
    expect(barButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: barButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);

    expect(sheet.instance.store.getState().charts.charts).toMatchObject([
      { kind: 'bar', source: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 } },
    ]);

    tb.dispose();
  });

  it('applies row, protection, and sheet-tab formatting through the Home Format dropdown', async () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const formatButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="formatCellsHome"]',
    );
    expect(formatButton).toBeTruthy();
    formatButton?.click();
    expect(host.querySelectorAll('#menu-format-cells .app__menu-item--iconic').length).toBe(18);
    expect(host.querySelectorAll('#menu-format-cells .app__color-swatch').length).toBe(14);
    const visibilityTrigger = host.querySelector<HTMLButtonElement>(
      '#menu-format-cells [data-format-submenu="visibility"]',
    );
    const tabColorTrigger = host.querySelector<HTMLButtonElement>(
      '#menu-format-cells [data-format-submenu="tabColor"]',
    );
    const visibilityPanel = host.querySelector<HTMLElement>('#menu-format-cells-visibility');
    const tabColorPanel = host.querySelector<HTMLElement>('#menu-format-cells-tabColor');
    expect(visibilityTrigger?.getAttribute('aria-controls')).toBe('menu-format-cells-visibility');
    expect(tabColorTrigger?.getAttribute('aria-controls')).toBe('menu-format-cells-tabColor');
    expect(visibilityPanel?.hidden).toBe(true);
    expect(tabColorPanel?.hidden).toBe(true);
    visibilityTrigger?.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }));
    expect(visibilityPanel?.hidden).toBe(false);
    expect(visibilityTrigger?.getAttribute('aria-expanded')).toBe('true');
    expect(
      host
        .querySelector<HTMLButtonElement>('#menu-format-cells [data-cell-format="tab-color-red"]')
        ?.classList.contains('app__color-swatch'),
    ).toBe(true);
    expect(
      host
        .querySelector<HTMLButtonElement>('#menu-format-cells [data-cell-format="tab-color-none"]')
        ?.getAttribute('aria-checked'),
    ).toBe('true');
    expect(
      host
        .querySelector<HTMLButtonElement>('#menu-format-cells [data-cell-format="tab-color-red"]')
        ?.getAttribute('aria-checked'),
    ).toBe('false');
    tabColorTrigger?.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }));
    expect(tabColorPanel?.hidden).toBe(false);
    expect(visibilityPanel?.hidden).toBe(true);
    const lockCellButton = host.querySelector<HTMLButtonElement>(
      '#menu-format-cells [data-cell-format="lock-cell"]',
    );
    expect(lockCellButton?.getAttribute('role')).toBe('menuitemcheckbox');
    expect(lockCellButton?.getAttribute('aria-checked')).toBe('true');
    expect(lockCellButton?.classList.contains('app__menu-item--checked')).toBe(true);
    const showRowsBeforeHide = host.querySelector<HTMLButtonElement>(
      '[data-cell-format="show-rows"]',
    );
    const showColsBeforeHide = host.querySelector<HTMLButtonElement>(
      '[data-cell-format="show-cols"]',
    );
    const moveCopyButton = host.querySelector<HTMLButtonElement>(
      '[data-cell-format="move-sheet-copy"]',
    );
    const renameSheetButton = host.querySelector<HTMLButtonElement>(
      '[data-cell-format="rename-sheet"]',
    );
    const unhideSheetButton = host.querySelector<HTMLButtonElement>(
      '[data-cell-format="unhide-sheet"]',
    );
    expect(showRowsBeforeHide?.getAttribute('aria-disabled')).toBe('true');
    expect(showRowsBeforeHide?.dataset.menuDisabledReason).toBe('No hidden rows are selected.');
    expect(showColsBeforeHide?.getAttribute('aria-disabled')).toBe('true');
    expect(showColsBeforeHide?.dataset.menuDisabledReason).toBe('No hidden columns are selected.');
    expect(renameSheetButton?.getAttribute('aria-disabled')).toBe('true');
    expect(renameSheetButton?.dataset.menuDisabledReason).toBe(
      'This workbook engine cannot rename, move, hide, or unhide sheets.',
    );
    expect(moveCopyButton?.getAttribute('aria-disabled')).toBe('true');
    expect(moveCopyButton?.dataset.menuDisabledReason).toBe(
      'This workbook engine cannot rename, move, hide, or unhide sheets.',
    );
    expect(unhideSheetButton?.getAttribute('aria-disabled')).toBe('true');
    expect(unhideSheetButton?.dataset.menuDisabledReason).toBe(
      'This workbook engine cannot rename, move, hide, or unhide sheets.',
    );
    const hideRowsButton = host.querySelector<HTMLButtonElement>('[data-cell-format="hide-rows"]');
    expect(hideRowsButton).toBeTruthy();
    const hideEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(hideEvent, 'target', { value: hideRowsButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(hideEvent)).toBe(true);
    expect(sheet.instance.store.getState().layout.hiddenRows.has(1)).toBe(true);

    formatButton?.click();
    const showRowsButton = host.querySelector<HTMLButtonElement>('[data-cell-format="show-rows"]');
    expect(showRowsButton).toBeTruthy();
    expect(showRowsButton?.getAttribute('aria-disabled')).toBe('false');
    expect(showRowsButton?.dataset.menuDisabledReason).toBeUndefined();
    const showEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(showEvent, 'target', { value: showRowsButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(showEvent)).toBe(true);
    expect(sheet.instance.store.getState().layout.hiddenRows.has(1)).toBe(false);

    formatButton?.click();
    const unlockButton = host.querySelector<HTMLButtonElement>('[data-cell-format="unlock-cell"]');
    expect(unlockButton).toBeTruthy();
    const unlockEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(unlockEvent, 'target', { value: unlockButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(unlockEvent)).toBe(true);
    expect(
      sheet.instance.store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 1 }))
        ?.locked,
    ).toBe(false);
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'formatCellsHome', menuId: 'menu-format-cells' },
      formatButton as HTMLButtonElement,
    );
    const unlockedLockCellButton = host.querySelector<HTMLButtonElement>(
      '#menu-format-cells [data-cell-format="lock-cell"]',
    );
    expect(unlockedLockCellButton?.getAttribute('aria-checked')).toBe('false');
    expect(unlockedLockCellButton?.classList.contains('app__menu-item--checked')).toBe(false);

    formatButton?.click();
    const tabColorButton = host.querySelector<HTMLButtonElement>(
      '[data-cell-format="tab-color-red"]',
    );
    expect(tabColorButton).toBeTruthy();
    const tabColorEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(tabColorEvent, 'target', { value: tabColorButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(tabColorEvent)).toBe(true);
    expect(sheet.instance.store.getState().layout.sheetTabColors.get(0)).toBe('#c00000');
    formatButton?.click();
    expect(
      host
        .querySelector<HTMLButtonElement>('#menu-format-cells [data-cell-format="tab-color-red"]')
        ?.getAttribute('aria-checked'),
    ).toBe('true');
    expect(
      host
        .querySelector<HTMLButtonElement>('#menu-format-cells [data-cell-format="tab-color-red"]')
        ?.classList.contains('app__color-swatch--active'),
    ).toBe(true);

    formatButton?.click();
    const rowHeightButton = host.querySelector<HTMLButtonElement>(
      '[data-cell-format="row-height"]',
    );
    expect(rowHeightButton).toBeTruthy();
    const rowHeightEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(rowHeightEvent, 'target', { value: rowHeightButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(rowHeightEvent)).toBe(true);
    const dialog = document.body.querySelector<HTMLElement>('.app__dlg');
    const input = dialog?.querySelector<HTMLInputElement>('input');
    expect(input).toBeTruthy();
    if (!input) throw new Error('Row height prompt input was not rendered');
    input.value = '48';
    dialog?.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter' }));
    await Promise.resolve();
    await Promise.resolve();
    expect(sheet.instance.store.getState().layout.rowHeights.get(1)).toBe(48);

    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 });
    formatButton?.click();
    const hideColsButton = host.querySelector<HTMLButtonElement>('[data-cell-format="hide-cols"]');
    expect(hideColsButton).toBeTruthy();
    const hideColsEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(hideColsEvent, 'target', { value: hideColsButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(hideColsEvent)).toBe(true);
    expect(sheet.instance.store.getState().layout.hiddenCols.has(1)).toBe(true);

    formatButton?.click();
    const showColsButton = host.querySelector<HTMLButtonElement>('[data-cell-format="show-cols"]');
    expect(showColsButton).toBeTruthy();
    expect(showColsButton?.getAttribute('aria-disabled')).toBe('false');
    expect(showColsButton?.dataset.menuDisabledReason).toBeUndefined();
    const showColsEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(showColsEvent, 'target', { value: showColsButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(showColsEvent)).toBe(true);
    expect(sheet.instance.store.getState().layout.hiddenCols.has(1)).toBe(false);

    formatButton?.click();
    const colWidthButton = host.querySelector<HTMLButtonElement>('[data-cell-format="col-width"]');
    expect(colWidthButton).toBeTruthy();
    const colWidthEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(colWidthEvent, 'target', { value: colWidthButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(colWidthEvent)).toBe(true);
    const colDialog = document.body.querySelector<HTMLElement>('.app__dlg');
    const colInput = colDialog?.querySelector<HTMLInputElement>('input');
    expect(colInput).toBeTruthy();
    if (!colInput) throw new Error('Column width prompt input was not rendered');
    colInput.value = '96';
    colDialog?.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter' }));
    await Promise.resolve();
    await Promise.resolve();
    expect(sheet.instance.store.getState().layout.colWidths.get(1)).toBe(96);

    tb.dispose();
  });

  it('disables the sheet move-or-copy entry when the host cannot reorder sheets', () => {
    const added = sheet.workbook.addSheet();
    expect(added).toBeGreaterThan(0);
    mutators.setSheetIndex(sheet.instance.store, added);
    mutators.setRange(sheet.instance.store, { sheet: added, r0: 0, c0: 0, r1: 0, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const formatButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="formatCellsHome"]',
    );
    expect(formatButton).toBeTruthy();
    formatButton?.click();
    const moveCopyButton = host.querySelector<HTMLButtonElement>(
      '[data-cell-format="move-sheet-copy"]',
    );
    expect(moveCopyButton).toBeTruthy();
    expect(moveCopyButton?.disabled).toBe(true);
    expect(moveCopyButton?.getAttribute('aria-disabled')).toBe('true');

    tb.dispose();
  });

  it('updates manual page breaks through the Excel-style Page Layout Breaks dropdown', () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 4, c0: 2, r1: 4, c1: 2 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('pageLayout');

    const breaksButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="pageBreaks"]',
    );
    expect(breaksButton).toBeTruthy();
    breaksButton?.click();
    expect(host.querySelectorAll('#menu-page-breaks .app__menu-item--iconic')).toHaveLength(3);
    const insertButton = host.querySelector<HTMLButtonElement>('[data-page-break-action="insert"]');
    const removeButton = host.querySelector<HTMLButtonElement>('[data-page-break-action="remove"]');
    const resetButton = host.querySelector<HTMLButtonElement>(
      '[data-page-break-action="reset-all"]',
    );
    expect(insertButton).toBeTruthy();
    expect(removeButton).toBeTruthy();
    expect(
      host.querySelector<HTMLElement>('#menu-page-breaks .app__menu-icon--break-page'),
    ).toBeTruthy();
    expect(insertButton?.textContent).toContain('Insert Page Break');
    expect(removeButton?.textContent).toContain('Remove Page Break');
    expect(insertButton?.disabled).toBe(false);
    expect(removeButton?.disabled).toBe(true);
    expect(removeButton?.getAttribute('aria-disabled')).toBe('true');
    expect(removeButton?.dataset.menuDisabledReason).toBe(
      'There is no manual page break at the selection.',
    );
    expect(resetButton?.disabled).toBe(true);
    expect(resetButton?.dataset.menuDisabledReason).toBe(
      'There are no manual page breaks to reset.',
    );

    const insertEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(insertEvent, 'target', { value: insertButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(insertEvent)).toBe(true);

    expect(getPageSetup(sheet.instance.store.getState(), 0).manualPageBreakRows).toEqual([4]);
    expect(getPageSetup(sheet.instance.store.getState(), 0).manualPageBreakCols).toEqual([2]);

    breaksButton?.click();
    const enabledRemoveButton = host.querySelector<HTMLButtonElement>(
      '[data-page-break-action="remove"]',
    );
    const enabledResetButton = host.querySelector<HTMLButtonElement>(
      '[data-page-break-action="reset-all"]',
    );
    expect(enabledRemoveButton?.disabled).toBe(false);
    expect(enabledRemoveButton?.getAttribute('aria-disabled')).toBe('false');
    expect(enabledRemoveButton?.dataset.menuDisabledReason).toBeUndefined();
    expect(enabledResetButton?.disabled).toBe(false);
    expect(enabledResetButton?.dataset.menuDisabledReason).toBeUndefined();
    const removeEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(removeEvent, 'target', { value: enabledRemoveButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(removeEvent)).toBe(true);
    expect(getPageSetup(sheet.instance.store.getState(), 0).manualPageBreakRows).toBeUndefined();
    expect(getPageSetup(sheet.instance.store.getState(), 0).manualPageBreakCols).toBeUndefined();

    breaksButton?.click();
    expect(enabledRemoveButton?.disabled).toBe(true);
    expect(enabledRemoveButton?.dataset.menuDisabledReason).toBe(
      'There is no manual page break at the selection.',
    );
    expect(enabledResetButton?.disabled).toBe(true);

    const insertAgainEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(insertAgainEvent, 'target', { value: insertButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(insertAgainEvent)).toBe(true);
    breaksButton?.click();
    const resetAfterRemoveColButton = host.querySelector<HTMLButtonElement>(
      '[data-page-break-action="reset-all"]',
    );
    expect(resetAfterRemoveColButton?.disabled).toBe(false);
    const resetEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(resetEvent, 'target', { value: resetAfterRemoveColButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(resetEvent)).toBe(true);

    expect(getPageSetup(sheet.instance.store.getState(), 0).manualPageBreakRows).toBeUndefined();
    expect(getPageSetup(sheet.instance.store.getState(), 0).manualPageBreakCols).toBeUndefined();

    tb.dispose();
  });

  it('projects the active rotation inside the Text Orientation dropdown', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 1, col: 1 });
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const orientationButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="textOrientation"]',
    );
    expect(orientationButton).toBeTruthy();
    orientationButton?.click();
    const ccwButton = host.querySelector<HTMLButtonElement>(
      '#menu-text-orientation [data-text-orientation="ccw"]',
    );
    expect(ccwButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: ccwButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(sheet.instance.store.getState().ui.pendingFormat).toEqual({
      addr: { sheet: 0, row: 1, col: 1 },
      format: { rotation: 45 },
    });

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'textOrientation', menuId: 'menu-text-orientation' },
      orientationButton,
    );
    expect(ccwButton?.getAttribute('role')).toBe('menuitemradio');
    expect(ccwButton?.getAttribute('aria-checked')).toBe('true');
    expect(ccwButton?.classList.contains('app__menu-item--active')).toBe(true);

    tb.dispose();
  });

  it('clears formula audit arrows by kind through the Clear Arrows dropdown', () => {
    const precedent = {
      kind: 'precedent' as const,
      from: { sheet: 0, row: 0, col: 0 },
      to: { sheet: 0, row: 0, col: 2 },
    };
    const dependent = {
      kind: 'dependent' as const,
      from: { sheet: 0, row: 0, col: 1 },
      to: { sheet: 0, row: 0, col: 2 },
    };
    mutators.addTrace(sheet.instance.store, precedent);
    mutators.addTrace(sheet.instance.store, dependent);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('formulas');

    const clearArrowsButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="clearArrows"]',
    );
    expect(clearArrowsButton).toBeTruthy();
    clearArrowsButton?.click();
    expect(host.querySelectorAll('#menu-clear-arrows .app__menu-item--iconic').length).toBe(3);
    const clearAllButton = host.querySelector<HTMLButtonElement>(
      '[data-formula-audit-action="clear-all"]',
    );
    const clearPrecedentsButton = host.querySelector<HTMLButtonElement>(
      '[data-formula-audit-action="clear-precedents"]',
    );
    const clearDependentsButton = host.querySelector<HTMLButtonElement>(
      '[data-formula-audit-action="clear-dependents"]',
    );
    expect(clearAllButton?.disabled).toBe(false);
    expect(clearPrecedentsButton).toBeTruthy();
    expect(clearPrecedentsButton?.disabled).toBe(false);
    expect(clearDependentsButton?.disabled).toBe(false);
    const clearPrecedentsEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(clearPrecedentsEvent, 'target', { value: clearPrecedentsButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(clearPrecedentsEvent)).toBe(true);
    expect(sheet.instance.store.getState().traces.items).toEqual([dependent]);

    clearArrowsButton?.click();
    expect(clearAllButton?.disabled).toBe(false);
    expect(clearPrecedentsButton?.disabled).toBe(true);
    expect(clearDependentsButton?.disabled).toBe(false);
    expect(clearDependentsButton).toBeTruthy();
    const clearDependentsEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(clearDependentsEvent, 'target', { value: clearDependentsButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(clearDependentsEvent)).toBe(true);
    expect(sheet.instance.store.getState().traces.items).toEqual([]);

    clearArrowsButton?.click();
    expect(clearAllButton?.disabled).toBe(true);
    expect(clearPrecedentsButton?.disabled).toBe(true);
    expect(clearDependentsButton?.disabled).toBe(true);

    tb.dispose();
  });

  it('sets and deletes sheet background from the Excel-style primary button', async () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('pageLayout');

    const backgroundButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="sheetBackground"]',
    );
    expect(backgroundButton).toBeTruthy();
    expect(backgroundButton?.dataset.ribbonActivation).toBe('primaryAction');
    expect(backgroundButton?.dataset.ribbonMenuId).toBeUndefined();
    expect(backgroundButton?.getAttribute('aria-haspopup')).toBeNull();
    expect(backgroundButton?.textContent).toContain('Background');
    expect(backgroundButton?.getAttribute('aria-label')).toBe('Background');
    backgroundButton?.click();

    const input = document.body.querySelector<HTMLInputElement>('input[type="file"]');
    expect(input).toBeTruthy();
    if (!input) throw new Error('Sheet background file picker input was not rendered');
    const file = new File(['background'], 'background.png', { type: 'image/png' });
    Object.defineProperty(input, 'files', { value: [file], configurable: true });
    input.dispatchEvent(new Event('change', { bubbles: true }));

    await waitFor(() =>
      Boolean(sheet.instance.store.getState().ui.sheetBackgroundImages.get(0)?.startsWith('data:')),
    );
    expect(document.body.querySelector('input[type="file"]')).toBeNull();
    expect(backgroundButton?.textContent).toContain('Delete Background');
    expect(backgroundButton?.getAttribute('aria-label')).toBe('Delete Background');
    expect(host.querySelector('#menu-sheet-background')).toBeNull();

    mutators.setSheetBackgroundImage(
      sheet.instance.store,
      0,
      'https://example.test/second-background.png',
    );
    backgroundButton?.click();
    expect(sheet.instance.store.getState().ui.sheetBackgroundImages.has(0)).toBe(false);
    expect(backgroundButton?.textContent).toContain('Background');
    expect(backgroundButton?.getAttribute('aria-label')).toBe('Background');
    expect(document.body.querySelector('.app__dlg')).toBeNull();

    tb.dispose();
  });

  it('opens Page Setup Sheet tab from Print Titles primary click and keeps secondary actions', () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 2, r1: 3, c1: 4 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('pageLayout');

    const printTitlesButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="printTitles"]',
    );
    expect(printTitlesButton).toBeTruthy();
    expect(printTitlesButton?.dataset.ribbonActivation).toBe('dialog');
    expect(printTitlesButton?.dataset.ribbonMenuId).toBeUndefined();
    expect(printTitlesButton?.getAttribute('aria-haspopup')).toBeNull();
    printTitlesButton?.click();
    const pageSetupDialog = document.body.querySelector<HTMLElement>('.fc-pgsetup');
    expect(pageSetupDialog?.hidden).toBe(false);
    expect(
      pageSetupDialog
        ?.querySelector<HTMLButtonElement>('[data-pgsetup-tab="sheet"]')
        ?.getAttribute('aria-selected'),
    ).toBe('true');
    pageSetupDialog
      ?.querySelector<HTMLButtonElement>('.fc-fmtdlg__close')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(pageSetupDialog?.hidden).toBe(true);
    expect(host.querySelector('#menu-print-titles')).toBeNull();

    tb.dispose();
  });

  it('wires every registered dynamic ribbon dropdown to a rendered button and menu', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    expect(Array.from(tb.dropdownsApi?.DYNAMIC_RIBBON_DROPDOWN_IDS ?? []).sort()).toEqual(
      Object.values(RIBBON_DROPDOWN_MENU_FOR_COMMAND).sort(),
    );
    expect(new Set(Object.values(RIBBON_DROPDOWN_MENU_FOR_COMMAND)).size).toBe(
      Object.values(RIBBON_DROPDOWN_MENU_FOR_COMMAND).length,
    );

    for (const [command, menuId] of Object.entries(RIBBON_DROPDOWN_MENU_FOR_COMMAND)) {
      const button = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`);
      const menu = host.querySelector<HTMLDivElement>(`#${menuId}`);
      expect(button, `${command} button`).toBeTruthy();
      expect(menu, `${command} menu`).toBeTruthy();
      expect(button?.dataset.ribbonMenuId, `${command} menu id metadata`).toBe(menuId);
      expect(button?.getAttribute('aria-haspopup'), `${command} aria-haspopup`).toBe('menu');
      expect(
        button?.querySelector('.demo__rb-split-chevron'),
        `${command} renders dropdown affordance`,
      ).toBeTruthy();
      expect(tb.dropdownsApi?.dynamicDropdownSpecForButton(button as HTMLButtonElement)).toEqual({
        command,
        menuId,
      });
      expect(tb.dropdownsApi?.dynamicDropdownSpecForMenu(menu as HTMLDivElement)).toEqual({
        command,
        menuId,
      });

      tb.dropdownsApi?.openDynamicRibbonDropdown({ command, menuId }, button as HTMLButtonElement);
      expect(menu?.hidden, `${command} opens ${menuId}`).toBe(false);
      expect(button?.getAttribute('aria-expanded'), `${command} aria-expanded`).toBe('true');
      tb.dropdownsApi?.closeDynamicRibbonDropdown({ command, menuId });
      expect(menu?.hidden, `${command} closes ${menuId}`).toBe(true);
    }

    tb.dispose();
  }, 30_000);

  it('opens every dropdown and gallery command menu from primary click', () => {
    const onCommand = vi.fn();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
      onCommand,
    });
    const misses: string[] = [];

    for (const command of RIBBON_DYNAMIC_MENU_FIRST_COMMANDS) {
      const menuId = RIBBON_DROPDOWN_MENU_FOR_COMMAND[command];
      if (!menuId) {
        misses.push(`${command}:menu-map`);
        continue;
      }
      const button = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`);
      const menu = host.querySelector<HTMLDivElement>(`#${menuId}`);
      if (!button || !menu) {
        misses.push(`${command}:missing`);
        continue;
      }
      onCommand.mockClear();
      tb.dropdownsApi?.closeDynamicRibbonDropdown({ command, menuId });
      button.click();
      if (menu.hidden) misses.push(`${command}:closed`);
      if (button.getAttribute('aria-expanded') !== 'true') misses.push(`${command}:aria`);
      if (onCommand.mock.calls.length > 0) misses.push(`${command}:command`);
      tb.dropdownsApi?.closeDynamicRibbonDropdown({ command, menuId });
    }

    expect(misses).toEqual([]);
    tb.dispose();
  });

  it('opens every externally owned ribbon menu from primary click', () => {
    const onCommand = vi.fn();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
      onCommand,
    });
    tb.rerender();
    const misses: string[] = [];

    for (const command of RIBBON_EXTERNAL_MENU_FIRST_COMMANDS) {
      const menuId = RIBBON_EXTERNAL_MENU_FOR_COMMAND[command];
      const button = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`);
      const menu = host.querySelector<HTMLDivElement>(`#${menuId}`);
      if (!button || !menu) {
        misses.push(`${command}:missing`);
        continue;
      }
      onCommand.mockClear();
      if (!menu.hidden) button.click();
      button.click();
      if (menu.hidden) misses.push(`${command}:closed`);
      if (button.getAttribute('aria-expanded') !== 'true') misses.push(`${command}:aria`);
      if (onCommand.mock.calls.length > 0) misses.push(`${command}:command`);
    }

    expect(misses).toEqual([]);
    tb.dispose();
  });

  it('keeps the shared ribbon activation model aligned with dropdown menus', () => {
    for (const [command, menuId] of Object.entries(RIBBON_MENU_FOR_COMMAND)) {
      expect(RIBBON_SPLIT_BUTTON_COMMANDS.has(command), `${command} renders as menu button`).toBe(
        true,
      );
      expect(ribbonActivationForCommand(command).menuId, `${command} activation menu`).toBe(menuId);
    }

    for (const command of RIBBON_PRIMARY_ACTION_SPLIT_COMMANDS) {
      expect(RIBBON_SPLIT_BUTTON_COMMANDS.has(command), `${command} renders as split`).toBe(true);
      expect(
        RIBBON_DROPDOWN_MENU_FOR_COMMAND[command],
        `${command} has a secondary menu`,
      ).toBeTruthy();
      expect(ribbonActivationForCommand(command).kind, `${command} activation kind`).toBe(
        'splitPrimary',
      );
    }
    for (const command of RIBBON_SPLIT_TOGGLE_COMMANDS) {
      expect(RIBBON_SPLIT_BUTTON_COMMANDS.has(command), `${command} renders as split toggle`).toBe(
        true,
      );
      expect(
        RIBBON_DROPDOWN_MENU_FOR_COMMAND[command],
        `${command} has a secondary menu`,
      ).toBeTruthy();
      expect(ribbonActivationForCommand(command).kind, `${command} activation kind`).toBe(
        'splitToggle',
      );
    }

    expect(ribbonActivationForCommand('formatTableHome').kind).toBe('gallery');
    expect(ribbonActivationForCommand('conditional').kind).toBe('gallery');
    expect(ribbonActivationForCommand('dataValidation').kind).toBe('splitPrimary');
    expect(ribbonActivationForCommand('deleteCommentReview').kind).toBe('splitPrimary');
    expect(ribbonActivationForCommand('errorChecking').kind).toBe('splitPrimary');
    expect(ribbonActivationForCommand('protect').kind).toBe('splitPrimary');
    expect(ribbonActivationForCommand('protectReview').kind).toBe('splitPrimary');
    expect(ribbonActivationForCommand('script').kind).toBe('splitPrimary');
    expect(ribbonActivationForCommand('addIn').kind).toBe('splitPrimary');
    expect(ribbonActivationForCommand('pdf').kind).toBe('splitPrimary');
    expect(ribbonActivationForCommand('watch').kind).toBe('splitPrimary');
    expect(ribbonActivationForCommand('watchView').kind).toBe('splitPrimary');
    expect(ribbonActivationForCommand('borders').kind).toBe('dropdown');
    expect(ribbonActivationForCommand('borders').menuId).toBe(RIBBON_BORDERS_MENU_ID);
    expect(RIBBON_EXTERNAL_MENU_FOR_COMMAND.borders).toBe(RIBBON_BORDERS_MENU_ID);
    expect(ribbonActivationForCommand('pageSetup').kind).toBe('dialog');
    expect(ribbonActivationForCommand('printTitles').kind).toBe('dialog');
    expect(ribbonActivationForCommand('sum').kind).toBe('dialog');
    expect(ribbonActivationForCommand('bold').kind).toBe('toggle');
    expect(ribbonActivationForCommand('underline').kind).toBe('splitToggle');
    expect(ribbonActivationForCommand('helpSearch').kind).toBe('disabled');
    expect(RIBBON_PRIMARY_ACTION_COMMANDS.has('print')).toBe(true);
    expect(RIBBON_PRIMARY_ACTION_COMMANDS.has('sheetBackground')).toBe(true);
    expect(RIBBON_PRIMARY_ACTION_SPLIT_COMMANDS.has('pivotTableInsert')).toBe(true);
    expect(ribbonActivationForCommand('formatTableInsert').kind).toBe('dialog');
  });

  it('fixtures the audited menu-backed ribbon activation categories', () => {
    expect(Array.from(RIBBON_PRIMARY_ACTION_SPLIT_COMMANDS).sort()).toEqual(
      [...RIBBON_AUDITED_PRIMARY_ACTION_SPLIT_COMMANDS].sort(),
    );
    expect(Array.from(RIBBON_SPLIT_TOGGLE_COMMANDS).sort()).toEqual(
      [...RIBBON_AUDITED_SPLIT_TOGGLE_COMMANDS].sort(),
    );
    expect(Array.from(RIBBON_GALLERY_COMMANDS).sort()).toEqual(
      [...RIBBON_AUDITED_GALLERY_COMMANDS].sort(),
    );
    expect(Array.from(RIBBON_DROPDOWN_COMMANDS).sort()).toEqual(
      [...RIBBON_AUDITED_DROPDOWN_COMMANDS].sort(),
    );
  });

  it('keeps non-menu ribbon activation sets mutually exclusive', () => {
    const nonMenuSets = ribbonActivationCategories().filter(
      ([kind]) =>
        kind === 'primaryAction' || kind === 'dialog' || kind === 'toggle' || kind === 'disabled',
    );
    const overlaps: string[] = [];
    for (const [index, [leftName, left]] of nonMenuSets.entries()) {
      for (const [rightName, right] of nonMenuSets.slice(index + 1)) {
        for (const command of left) {
          if (right.has(command)) overlaps.push(`${command}:${leftName}/${rightName}`);
        }
      }
    }

    const menuOverlaps = Array.from(RIBBON_PRIMARY_ACTION_COMMANDS)
      .filter((command) => RIBBON_MENU_FOR_COMMAND[command])
      .map((command) => `${command}:primary/menu`);

    expect([...overlaps, ...menuOverlaps]).toEqual([]);
  });

  it('only marks menu-backed commands as gallery or split-primary activations', () => {
    for (const command of RIBBON_GALLERY_COMMANDS) {
      expect(RIBBON_MENU_FOR_COMMAND[command], `${command} gallery has menu`).toBeTruthy();
      expect(ribbonActivationForCommand(command).kind, `${command} gallery activation`).toBe(
        'gallery',
      );
    }
    for (const command of RIBBON_PRIMARY_ACTION_SPLIT_COMMANDS) {
      expect(RIBBON_MENU_FOR_COMMAND[command], `${command} split has menu`).toBeTruthy();
      expect(ribbonActivationForCommand(command).kind, `${command} split activation`).toBe(
        'splitPrimary',
      );
    }
    for (const command of RIBBON_SPLIT_TOGGLE_COMMANDS) {
      expect(RIBBON_MENU_FOR_COMMAND[command], `${command} split toggle has menu`).toBeTruthy();
      expect(ribbonActivationForCommand(command).kind, `${command} split toggle activation`).toBe(
        'splitToggle',
      );
    }
  });

  it('does not leave rendered primary-action commands implicit in the activation model', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    const implicit = Array.from(host.querySelectorAll<HTMLElement>('[data-ribbon-command]'))
      .filter((el): el is HTMLButtonElement => el instanceof HTMLButtonElement)
      .map((button) => button.dataset.ribbonCommand ?? '')
      .filter((command) => {
        const activation = ribbonActivationForCommand(command);
        return activation.kind === 'primaryAction' && !RIBBON_PRIMARY_ACTION_COMMANDS.has(command);
      })
      .sort();

    expect(implicit).toEqual([]);
    tb.dispose();
  });

  it('projects rendered ribbon button activation metadata from the shared resolver', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    const mismatches = Array.from(
      host.querySelectorAll<HTMLButtonElement>('button[data-ribbon-command]'),
    )
      .map((button) => {
        const command = button.dataset.ribbonCommand ?? '';
        const activation = ribbonActivationForCommand(command);
        const actualMenuId = button.dataset.ribbonMenuId;
        if (button.dataset.ribbonActivation !== activation.kind) {
          return `${command}:kind:${button.dataset.ribbonActivation}->${activation.kind}`;
        }
        if (actualMenuId !== activation.menuId) {
          return `${command}:menu:${actualMenuId ?? 'none'}->${activation.menuId ?? 'none'}`;
        }
        return null;
      })
      .filter((mismatch): mismatch is string => mismatch !== null)
      .sort();

    expect(mismatches).toEqual([]);
    tb.dispose();
  });

  it('renders exactly the shared activatable ribbon command surface as buttons', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    const renderedButtonCommands = Array.from(
      host.querySelectorAll<HTMLButtonElement>('button[data-ribbon-command]'),
    )
      .map((button) => button.dataset.ribbonCommand ?? '')
      .sort();
    const expectedButtonCommands = ribbonActivatableSurfaceCommandIds().sort();

    expect(renderedButtonCommands).toEqual(expectedButtonCommands);
    tb.dispose();
  });

  it('renders disabled activation commands with disabled accessibility state', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });

    for (const command of RIBBON_DISABLED_COMMANDS) {
      const button = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`);
      expect(button, `${command} button`).toBeTruthy();
      expect(button?.disabled, `${command} disabled`).toBe(true);
      expect(button?.getAttribute('aria-disabled'), `${command} aria-disabled`).toBe('true');
      expect(button?.getAttribute('aria-description'), `${command} aria-description`).toBe(
        'Coming soon',
      );
      expect(button?.dataset.ribbonDisabledReason, `${command} disabled reason`).toBe(
        'Coming soon',
      );
      expect(button?.title, `${command} title`).toContain('Coming soon');
      expect(button?.dataset.ribbonActivation, `${command} activation`).toBe('disabled');
      expect(button?.dataset.ribbonMenuId, `${command} menu`).toBeUndefined();
    }

    tb.dispose();
  });

  it('localizes disabled activation command reasons through shared ribbon text', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
      lang: 'ja',
    });
    const button = host.querySelector<HTMLButtonElement>('[data-ribbon-command="helpSearch"]');

    expect(button?.disabled).toBe(true);
    expect(button?.getAttribute('aria-description')).toBe('未実装');
    expect(button?.dataset.ribbonDisabledReason).toBe('未実装');
    expect(button?.title).toContain('未実装');

    tb.dispose();
  });

  it('does not classify ribbon layout row breaks as primary actions', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    const rowBreaks = Array.from(host.querySelectorAll<HTMLElement>('.demo__rb-break')).map(
      (el) => el.dataset.ribbonCommand,
    );

    expect(rowBreaks).toEqual(['font-row-2', 'alignment-row-2', 'number-row-2']);
    for (const command of rowBreaks) {
      expect(command).toBeTruthy();
      if (command) expect(RIBBON_PRIMARY_ACTION_COMMANDS.has(command)).toBe(false);
    }

    tb.dispose();
  });

  it('keeps every explicit primary-action command consumable by the shared dispatcher', () => {
    vi.spyOn(sheet.instance, 'openFindReplace').mockImplementation(() => undefined);
    vi.spyOn(sheet.instance, 'print').mockImplementation(() => undefined);
    vi.spyOn(sheet.instance, 'tracePrecedents').mockReturnValue(1);
    vi.spyOn(sheet.instance, 'traceDependents').mockReturnValue(1);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
      hooks: {
        automation: {
          allScripts: vi.fn(),
          recordActions: vi.fn(),
        },
        drawing: {
          setInkMode: vi.fn(),
        },
        page: {
          inspect: vi.fn(),
          outline: vi.fn(),
          sheetBackground: vi.fn(),
        },
        protection: {
          allowEditRanges: vi.fn(),
          runWorkbook: vi.fn(),
        },
        review: {
          accessibility: vi.fn(),
          selectComment: vi.fn(),
          spelling: vi.fn(),
          translate: vi.fn(),
        },
        sheetView: {
          deleteActive: vi.fn(),
          save: vi.fn(),
        },
        sortFilter: {
          customSort: vi.fn(),
          removeDuplicates: vi.fn(),
          sort: vi.fn(),
        },
      },
    });
    const missed = Array.from(RIBBON_PRIMARY_ACTION_COMMANDS)
      .filter((command) => tb.applyCommand(command) !== true)
      .sort();

    expect(missed).toEqual([]);
    tb.dispose();
  });

  it('keeps every primary-face menu command consumable from its primary face', () => {
    vi.spyOn(sheet.instance, 'openPivotTableDialog').mockImplementation(() => undefined);
    vi.spyOn(sheet.instance, 'openDataValidationDialog').mockImplementation(() => undefined);
    vi.spyOn(sheet.instance, 'openExternalLinksDialog').mockImplementation(() => undefined);
    vi.spyOn(sheet.instance, 'openNamedRangeDialog').mockImplementation(() => undefined);
    vi.spyOn(sheet.instance, 'openPageSetup').mockImplementation(() => undefined);
    vi.spyOn(sheet.instance, 'toggleWatchWindow').mockImplementation(() => undefined);
    vi.spyOn(sheet.instance, 'print').mockImplementation(() => undefined);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
      hooks: {
        automation: {
          addInManager: vi.fn(),
          runScript: vi.fn(),
        },
        clipboard: {
          paste: vi.fn(),
        },
        formula: {
          autoSum: vi.fn(),
          errorChecking: vi.fn(),
        },
        insert: {
          createRecommendedChart: vi.fn(),
          insertSymbol: vi.fn(),
        },
        page: {
          pdf: vi.fn(),
          sheetBackground: vi.fn(),
        },
        protection: {
          runSheet: vi.fn(),
        },
        review: {
          deleteComment: vi.fn(),
        },
        sortFilter: {
          splitTextToColumnsCustom: vi.fn(),
        },
      },
    });
    const missed = Array.from(RIBBON_PRIMARY_FACE_MENU_COMMANDS)
      .filter((command) => tb.applyCommand(command) !== true)
      .sort();

    expect(missed).toEqual([]);
    tb.dispose();
  });

  it('keeps every toggle-classified command consumable by the shared dispatcher', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
    });
    const missed = Array.from(RIBBON_TOGGLE_COMMANDS)
      .filter((command) => tb.applyCommand(command) !== true)
      .sort();

    expect(missed).toEqual([]);
    tb.dispose();
  });

  it('keeps the Paste secondary menu iconified under the shared shell', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const pasteButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="paste"]');
    expect(pasteButton).toBeTruthy();

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'paste', menuId: 'menu-paste' },
      pasteButton as HTMLButtonElement,
    );

    expect(host.querySelectorAll('#menu-paste .app__menu-item--iconic').length).toBe(8);
    expect(
      Array.from(host.querySelectorAll<HTMLButtonElement>('#menu-paste [data-paste-action]'))
        .filter((button) => !button.hidden)
        .map((button) => button.dataset.pasteAction),
    ).toEqual(['all', 'dialog']);
    expect(host.querySelector<HTMLElement>('#menu-paste .app__menu-sep')?.hidden).toBe(true);
    const pasteSpecial = host.querySelector<HTMLButtonElement>('[data-paste-action="dialog"]');
    const pasteAll = host.querySelector<HTMLButtonElement>('[data-paste-action="all"]');
    expect(pasteSpecial).toBeTruthy();
    expect(pasteAll?.getAttribute('aria-disabled')).toBe('false');
    expect(pasteSpecial?.disabled).toBe(true);
    expect(pasteSpecial?.getAttribute('aria-disabled')).toBe('true');
    expect(pasteSpecial?.getAttribute('aria-description')).toBe(
      'Copy or cut cells before using this paste option.',
    );
    expect(pasteSpecial?.dataset.menuDisabledReason).toBe(
      'Copy or cut cells before using this paste option.',
    );
    expect(pasteSpecial?.title).toContain('Copy or cut cells before using this paste option.');

    tb.dispose();
  });

  it('routes enabled Paste secondary actions through shared clipboard hooks', () => {
    seedNumber(sheet, 0, 0, 12);
    const snapshot = captureSnapshot(sheet.instance.store.getState(), {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 0,
      c1: 0,
    });
    expect(snapshot).toBeTruthy();
    if (!snapshot) throw new Error('Expected clipboard snapshot.');
    Object.defineProperty(sheet.instance, 'clipboard', {
      configurable: true,
      value: {
        detach: vi.fn(),
        getSnapshot: () => snapshot,
        runShortcut: vi.fn(),
      },
    });
    const openPasteSpecial = vi
      .spyOn(sheet.instance, 'openPasteSpecial')
      .mockImplementation(() => undefined);
    const pasteSpecial = vi.spyOn(sheet.instance, 'pasteSpecial').mockReturnValue(true);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const pasteButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="paste"]');
    expect(pasteButton).toBeTruthy();

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'paste', menuId: 'menu-paste' },
      pasteButton as HTMLButtonElement,
    );
    expect(
      Array.from(host.querySelectorAll<HTMLButtonElement>('#menu-paste [data-paste-action]')).some(
        (button) => button.hidden,
      ),
    ).toBe(false);
    const pasteSpecialButton = host.querySelector<HTMLButtonElement>(
      '[data-paste-action="dialog"]',
    );
    const pasteValuesButton = host.querySelector<HTMLButtonElement>('[data-paste-action="values"]');
    expect(pasteSpecialButton?.disabled).toBe(false);
    expect(pasteSpecialButton?.getAttribute('aria-description')).toBeNull();
    expect(pasteSpecialButton?.dataset.menuDisabledReason).toBeUndefined();
    expect(pasteValuesButton?.disabled).toBe(false);
    const dialogEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(dialogEvent, 'target', { value: pasteSpecialButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(dialogEvent)).toBe(true);
    expect(openPasteSpecial).toHaveBeenCalledTimes(1);

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'paste', menuId: 'menu-paste' },
      pasteButton as HTMLButtonElement,
    );
    const valuesEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(valuesEvent, 'target', { value: pasteValuesButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(valuesEvent)).toBe(true);
    expect(pasteSpecial).toHaveBeenCalledWith({
      operation: 'none',
      skipBlanks: false,
      transpose: false,
      what: 'values',
    });

    tb.dispose();
  });

  it('keeps primary dialog commands out of menu-first activation unless explicitly split', () => {
    const overlaps = Object.keys(RIBBON_DIALOG_OPENERS)
      .filter((command) => RIBBON_DROPDOWN_MENU_FOR_COMMAND[command])
      .sort();

    expect(overlaps).toEqual(Array.from(RIBBON_PRIMARY_SPLIT_DIALOG_COMMANDS).sort());
    for (const command of overlaps) {
      expect(RIBBON_PRIMARY_ACTION_SPLIT_COMMANDS.has(command), `${command} is primary split`).toBe(
        true,
      );
      expect(ribbonActivationForCommand(command).kind, `${command} activation`).toBe(
        'splitPrimary',
      );
    }
  });

  it('keeps dialog-classified ribbon commands backed by a shared dispatcher path', () => {
    const backedDialogs = new Set([
      ...Object.keys(RIBBON_DIALOG_OPENERS),
      ...Object.keys(RIBBON_FUNCTION_ARG_OPENERS),
      ...RIBBON_HOOK_DIALOG_COMMANDS,
    ]);

    const missing = Array.from(RIBBON_DIALOG_COMMANDS)
      .filter((command) => !backedDialogs.has(command))
      .sort();
    expect(missing).toEqual([]);
  });

  it('runs Error Checking from the primary button and keeps the menu secondary', async () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('formulas');

    const errorButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="errorChecking"]',
    );
    expect(errorButton).toBeTruthy();
    expect(errorButton?.dataset.ribbonActivation).toBe('splitPrimary');
    errorButton?.click();
    await Promise.resolve();
    await Promise.resolve();

    expect(host.querySelector<HTMLDivElement>('#menu-error-checking')?.hidden).toBe(true);
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'Error Checking',
    );
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'errorChecking', menuId: 'menu-error-checking' },
      errorButton as HTMLButtonElement,
    );
    expect(host.querySelector<HTMLDivElement>('#menu-error-checking')?.hidden).toBe(false);
    expect(host.querySelectorAll('#menu-error-checking .app__menu-item--iconic').length).toBe(3);
    expect(
      host.querySelector<HTMLButtonElement>('[data-formula-audit-action="trace-error"]'),
    ).toBeTruthy();
    expect(
      host
        .querySelector<HTMLButtonElement>('[data-formula-audit-action="trace-error"]')
        ?.getAttribute('aria-disabled'),
    ).toBe('true');
    expect(
      host.querySelector<HTMLButtonElement>('[data-formula-audit-action="trace-error"]')?.dataset
        .menuDisabledReason,
    ).toBe('The active cell does not contain a formula error.');
    expect(
      host
        .querySelector<HTMLButtonElement>('[data-formula-audit-action="ignore-error"]')
        ?.getAttribute('aria-disabled'),
    ).toBe('true');

    const errorAddr = { sheet: 0, row: 1, col: 1 };
    sheet.workbook.setFormula(errorAddr, '=1/0');
    sheet.instance.store.setState((state) => {
      const cells = new Map(state.data.cells);
      cells.set(addrKey(errorAddr), {
        value: { kind: 'error', code: 7, text: '#DIV/0!' },
        formula: '=1/0',
      });
      return { ...state, data: { ...state.data, cells } };
    });
    mutators.setActive(sheet.instance.store, errorAddr);
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 });
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'errorChecking', menuId: 'menu-error-checking' },
      errorButton as HTMLButtonElement,
    );
    expect(
      host
        .querySelector<HTMLButtonElement>('[data-formula-audit-action="trace-error"]')
        ?.getAttribute('aria-disabled'),
    ).toBe('false');
    expect(
      host.querySelector<HTMLButtonElement>('[data-formula-audit-action="trace-error"]')?.dataset
        .menuDisabledReason,
    ).toBeUndefined();
    expect(
      host
        .querySelector<HTMLButtonElement>('[data-formula-audit-action="ignore-error"]')
        ?.getAttribute('aria-disabled'),
    ).toBe('false');

    tb.dispose();
  });

  it('opens Watch Window from the primary button and keeps Add/Delete secondary', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('formulas');

    const watchButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="watch"]');
    expect(watchButton).toBeTruthy();
    expect(watchButton?.dataset.ribbonActivation).toBe('splitPrimary');
    watchButton?.click();

    expect(host.querySelector<HTMLDivElement>('#menu-watch-formulas')?.hidden).toBe(true);
    expect(sheet.instance.store.getState().ui.watchPanelOpen).toBe(true);

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'watch', menuId: 'menu-watch-formulas' },
      watchButton as HTMLButtonElement,
    );
    expect(host.querySelector<HTMLDivElement>('#menu-watch-formulas')?.hidden).toBe(false);
    expect(host.querySelectorAll('#menu-watch-formulas .app__menu-item--iconic').length).toBe(4);
    const addButton = host.querySelector<HTMLButtonElement>('[data-watch-action="add"]');
    expect(addButton).toBeTruthy();
    const deleteButton = host.querySelector<HTMLButtonElement>('[data-watch-action="delete"]');
    const deleteAllButton = host.querySelector<HTMLButtonElement>(
      '[data-watch-action="delete-all"]',
    );
    expect(deleteButton).toBeTruthy();
    expect(deleteButton?.disabled).toBe(true);
    expect(deleteButton?.dataset.menuDisabledReason).toBe('The active cell is not being watched.');
    expect(deleteAllButton?.disabled).toBe(true);
    expect(deleteAllButton?.dataset.menuDisabledReason).toBe('There are no watches to delete.');
    const addEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(addEvent, 'target', { value: addButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(addEvent)).toBe(true);
    expect(sheet.instance.store.getState().watch.watches).toContainEqual({
      sheet: 0,
      row: 0,
      col: 0,
    });
    expect(sheet.instance.store.getState().ui.watchPanelOpen).toBe(true);

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'watch', menuId: 'menu-watch-formulas' },
      watchButton as HTMLButtonElement,
    );
    expect(deleteButton?.disabled).toBe(false);
    expect(deleteButton?.getAttribute('aria-disabled')).toBe('false');
    expect(deleteButton?.dataset.menuDisabledReason).toBeUndefined();
    expect(deleteAllButton?.disabled).toBe(false);
    expect(deleteAllButton?.dataset.menuDisabledReason).toBeUndefined();
    const deleteEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(deleteEvent, 'target', { value: deleteButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(deleteEvent)).toBe(true);
    expect(sheet.instance.store.getState().watch.watches).toEqual([]);

    mutators.addWatch(sheet.instance.store, { sheet: 0, row: 0, col: 0 });

    mutators.setActive(sheet.instance.store, { sheet: 0, row: 1, col: 0 });
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'watch', menuId: 'menu-watch-formulas' },
      watchButton as HTMLButtonElement,
    );
    expect(deleteButton?.disabled).toBe(true);
    expect(deleteButton?.dataset.menuDisabledReason).toBe('The active cell is not being watched.');
    expect(deleteAllButton?.disabled).toBe(false);
    const deleteAllEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(deleteAllEvent, 'target', { value: deleteAllButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(deleteAllEvent)).toBe(true);
    expect(sheet.instance.store.getState().watch.watches).toEqual([]);

    tb.dispose();
  });

  it('classifies every rendered ribbon menu under a dispatcher or explicit external owner', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const externallyWiredMenuIds = new Set(Object.values(RIBBON_EXTERNAL_MENU_FOR_COMMAND));
    const unowned = Array.from(host.querySelectorAll<HTMLDivElement>('.app__menu'))
      .map((menu) => menu.id)
      .filter(
        (menuId) =>
          !(tb.dropdownsApi?.DYNAMIC_RIBBON_DROPDOWN_IDS.has(menuId) ?? false) &&
          !externallyWiredMenuIds.has(menuId),
      );

    expect(unowned).toEqual([]);
    tb.dispose();
  });

  it('keeps every rendered dynamic dropdown menu item dispatchable', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const missing: string[] = [];

    for (const menuId of tb.dropdownsApi?.DYNAMIC_RIBBON_DROPDOWN_IDS ?? []) {
      const menu = host.querySelector<HTMLElement>(`#${menuId}`);
      if (!menu) continue;
      for (const button of menu.querySelectorAll<HTMLButtonElement>('button')) {
        const keys = Object.keys(button.dataset);
        if (keys.length === 0) {
          missing.push(`${menuId}:${button.textContent ?? ''}`);
          continue;
        }
        if (!keys.some((key) => DYNAMIC_RIBBON_DROPDOWN_HANDLER_DATASET_KEYS.has(key))) {
          missing.push(`${menuId}:${button.textContent ?? ''}:${keys.join(',')}`);
        }
      }
    }

    expect(missing).toEqual([]);
    tb.dispose();
  });

  it('keeps every registered dynamic dropdown handler key represented by default menus', () => {
    mutators.upsertCustomPivotTableStyle(sheet.instance.store, {
      id: customPivotTableStyleId('Dispatch Coverage Pivot'),
      label: 'Dispatch Coverage Pivot',
      style: 'medium',
      color: '#70ad47',
      variant: 'bandedFirstCol',
    });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const renderedKeys = new Set<string>();

    for (const menuId of tb.dropdownsApi?.DYNAMIC_RIBBON_DROPDOWN_IDS ?? []) {
      const menu = host.querySelector<HTMLElement>(`#${menuId}`);
      if (!menu) continue;
      for (const element of menu.querySelectorAll<HTMLElement>('*')) {
        for (const key of Object.keys(element.dataset)) renderedKeys.add(key);
      }
    }

    const missing = Array.from(DYNAMIC_RIBBON_DROPDOWN_HANDLER_DATASET_KEYS)
      .filter((key) => !renderedKeys.has(key))
      .sort();
    expect(missing).toEqual([]);
    tb.dispose();
  });

  it('keeps rendered ribbon menus out of plain text-only fallback items', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const plainItems: string[] = [];

    const isStructuredMenuButton = (item: HTMLButtonElement): boolean =>
      item.classList.contains('app__menu-item--iconic') ||
      item.classList.contains('app__menu-item--preset') ||
      item.classList.contains('app__cellstyle-chip') ||
      item.classList.contains('app__tablestyle-swatch') ||
      item.classList.contains('app__visual-tile') ||
      item.classList.contains('app__symbol-tile') ||
      item.classList.contains('app__color-swatch') ||
      item.classList.contains('app__cf-choice') ||
      item.classList.contains('app__cf-icon-choice') ||
      item.classList.contains('app__submenu-item') ||
      item.classList.contains('fc-colorpalette__swatch') ||
      item.classList.contains('fc-colorpalette__action') ||
      !!item.querySelector(
        '.app__border-preview, .app__cf-icon, .app__menu-item__icon-spacer, .app__text-orientation-preview',
      );

    for (const menu of host.querySelectorAll<HTMLElement>('.app__menu')) {
      for (const item of menu.querySelectorAll<HTMLButtonElement>('button')) {
        if (!isStructuredMenuButton(item)) {
          plainItems.push(`${menu.id}:${item.textContent?.trim() ?? ''}`);
        }
      }
    }

    expect(plainItems).toEqual([]);
    tb.dispose();
  });

  it('consumes every rendered dynamic dropdown menu item click through the shared dispatcher', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: dynamicDropdownNoopOverrides(),
      helpers: stubHelpers(),
    });
    const missed: string[] = [];

    for (const menuId of tb.dropdownsApi?.DYNAMIC_RIBBON_DROPDOWN_IDS ?? []) {
      const menu = host.querySelector<HTMLElement>(`#${menuId}`);
      if (!menu) continue;
      for (const button of menu.querySelectorAll<HTMLButtonElement>('button')) {
        const event = new MouseEvent('click', { bubbles: true, cancelable: true });
        Object.defineProperty(event, 'target', { value: button });
        if (tb.dropdownsApi?.dynamicRibbonDropdownClick(event) !== true) {
          missed.push(`${menuId}:${button.textContent ?? ''}:${Object.keys(button.dataset)}`);
        }
      }
    }

    expect(missed).toEqual([]);
    tb.dispose();
  });

  it('applies Borders dropdown presets through the built-in menu owner after rerender', () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.rerender();

    const borderButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="borders"]');
    expect(borderButton).toBeTruthy();
    expect(borderButton?.dataset.ribbonActivation).toBe('dropdown');
    expect(borderButton?.dataset.ribbonMenuId).toBe(RIBBON_BORDERS_MENU_ID);
    expect(borderButton?.getAttribute('aria-haspopup')).toBe('menu');
    borderButton?.click();
    const menu = host.querySelector<HTMLDivElement>(`#${RIBBON_BORDERS_MENU_ID}`);
    expect(menu?.hidden).toBe(false);
    const lineColorTrigger = menu?.querySelector<HTMLButtonElement>(
      '[data-border-submenu="lineColor"]',
    );
    const lineStyleTrigger = menu?.querySelector<HTMLButtonElement>(
      '[data-border-submenu="lineStyle"]',
    );
    expect(lineColorTrigger?.getAttribute('aria-controls')).toBe('menu-borders-line-color');
    expect(lineStyleTrigger?.getAttribute('aria-controls')).toBe('menu-borders-line-style');
    expect(menu?.querySelector('#menu-borders-line-color')).toBeTruthy();
    expect(menu?.querySelector('#menu-borders-line-style')).toBeTruthy();
    host
      .querySelector<HTMLButtonElement>(`#${RIBBON_BORDERS_MENU_ID} [data-border-preset="bottom"]`)
      ?.click();

    expect(
      sheet.instance.store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 })),
    ).toBeUndefined();
    expect(sheet.instance.store.getState().ui.pendingFormat).toEqual({
      addr: { sheet: 0, row: 0, col: 0 },
      format: { borders: { bottom: { style: 'thin' } } },
    });

    tb.dispose();
  });

  it('closes an open Borders menu when a dynamic dropdown opens', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const borderButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="borders"]');
    expect(borderButton).toBeTruthy();
    borderButton?.click();
    const borderMenu = host.querySelector<HTMLDivElement>(`#${RIBBON_BORDERS_MENU_ID}`);
    expect(borderMenu?.hidden).toBe(false);

    const fillButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="fillHome"]');
    expect(fillButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'fillHome', menuId: 'menu-fill' },
      fillButton,
    );

    expect(borderMenu?.hidden).toBe(true);
    expect(host.querySelector<HTMLDivElement>('#menu-fill')?.hidden).toBe(false);

    tb.dispose();
  });

  it('opens the PivotTable dialog from the Insert ribbon primary button', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('insert');

    const pivotButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="pivotTableInsert"]',
    );
    expect(pivotButton).toBeTruthy();
    pivotButton?.click();

    expect(document.body.textContent).toContain('Create PivotTable');
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn')?.click();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'pivotTableInsert', menuId: 'menu-pivot-table' },
      pivotButton as HTMLButtonElement,
    );
    expect(host.querySelectorAll('#menu-pivot-table .app__menu-item--iconic').length).toBe(5);
    expect(
      host.querySelector<HTMLButtonElement>('[data-pivot-table-action="dialog"]'),
    ).toBeTruthy();
    expect(
      host.querySelector<HTMLButtonElement>('[data-pivot-table-action="refresh"]'),
    ).toBeTruthy();
    const recommendedButton = host.querySelector<HTMLButtonElement>(
      '[data-pivot-table-action="recommended"]',
    );
    expect(recommendedButton).toBeTruthy();
    const recommendedEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(recommendedEvent, 'target', { value: recommendedButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(recommendedEvent)).toBe(true);
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'Recommended PivotTables',
    );
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'Creating or editing PivotTable definitions',
    );
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'pivotTableInsert', menuId: 'menu-pivot-table' },
      pivotButton as HTMLButtonElement,
    );
    const newSheetButton = host.querySelector<HTMLButtonElement>(
      '[data-pivot-table-action="new-sheet"]',
    );
    expect(newSheetButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: newSheetButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(
      document.querySelector<HTMLInputElement>('input[name="fc-pivotdlg-destination"]:checked')
        ?.value,
    ).toBe('new');

    tb.dispose();
  });

  it('opens More Symbols from the Insert Symbol primary button', async () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('insert');

    const symbolButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="symbolInsert"]',
    );
    expect(symbolButton).toBeTruthy();
    expect(symbolButton?.dataset.ribbonActivation).toBe('splitPrimary');
    symbolButton?.click();

    await Promise.resolve();
    const dialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(dialog?.textContent).toContain('More Symbols');
    const input = dialog?.querySelector<HTMLInputElement>('input');
    expect(input).toBeTruthy();
    if (!input) throw new Error('Expected More Symbols input.');
    input.value = 'Ω';
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    await Promise.resolve();

    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'Ω',
    });

    tb.dispose();
  });

  it('opens the Insert Symbol secondary menu and inserts the selected symbol', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('insert');

    const symbolButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="symbolInsert"]',
    );
    expect(symbolButton).toBeTruthy();
    expect(
      tb.dropdownsApi?.dynamicDropdownSpecForButton(symbolButton as HTMLButtonElement),
    ).toEqual({
      command: 'symbolInsert',
      menuId: 'menu-symbol',
    });
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'symbolInsert', menuId: 'menu-symbol' },
      symbolButton as HTMLButtonElement,
    );
    const menu = host.querySelector<HTMLDivElement>('#menu-symbol');
    expect(menu?.hidden).toBe(false);
    expect(menu?.classList.contains('app__menu--symbols')).toBe(true);
    expect(menu?.querySelectorAll('.app__symbol-grid').length).toBeGreaterThan(0);
    expect(menu?.querySelectorAll('button.app__menu-item[data-symbol]').length).toBe(0);
    expect(menu?.querySelectorAll('.app__menu-item--iconic').length).toBe(1);
    const piButton = Array.from(
      menu?.querySelectorAll<HTMLButtonElement>('[data-symbol]') ?? [],
    ).find((button) => button.dataset.symbol === 'π');
    expect(piButton).toBeTruthy();
    expect(piButton?.classList.contains('app__symbol-tile')).toBe(true);
    expect(piButton?.querySelector('.app__symbol-tile__glyph')?.textContent).toBe('π');
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: piButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'π',
    });

    tb.dispose();
  });

  it('opens More Symbols from the Insert Symbol secondary menu and inserts custom text', async () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('insert');

    const symbolButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="symbolInsert"]',
    );
    expect(symbolButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'symbolInsert', menuId: 'menu-symbol' },
      symbolButton as HTMLButtonElement,
    );
    const moreButton = host.querySelector<HTMLButtonElement>('[data-symbol-action="more"]');
    expect(moreButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: moreButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);

    await Promise.resolve();
    const dialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(dialog?.textContent).toContain('More Symbols');
    const input = dialog?.querySelector<HTMLInputElement>('input');
    expect(input).toBeTruthy();
    if (!input) throw new Error('Expected More Symbols input.');
    input.value = 'Ω';
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    await Promise.resolve();

    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'Ω',
    });

    tb.dispose();
  });

  it('keeps Merge Cells as an icon split button and routes secondary merge actions', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    const helpers = stubHelpers();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: {
        ...helpers,
        createIcon: (name) => {
          const icon = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
          icon.setAttribute('class', 'demo__rb-icon');
          icon.dataset.icon = name;
          return icon;
        },
      },
    });

    const mergeButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="merge"]');
    expect(mergeButton).toBeTruthy();
    expect(mergeButton?.dataset.ribbonActivation).toBe('splitPrimary');
    expect(mergeButton?.dataset.ribbonMenuId).toBe('menu-merge');
    expect(mergeButton?.getAttribute('aria-haspopup')).toBe('menu');
    expect(mergeButton?.querySelector('.demo__rb-icon')).toBeTruthy();
    const textLabels = Array.from(mergeButton?.querySelectorAll('span') ?? []).filter(
      (span) =>
        !span.classList.contains('demo__rb-icon') &&
        !span.classList.contains('demo__rb-split-chevron'),
    );
    expect(textLabels).toEqual([]);

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'merge', menuId: 'menu-merge' },
      mergeButton as HTMLButtonElement,
    );
    const mergeItems = Array.from(
      host.querySelectorAll<HTMLButtonElement>('#menu-merge .app__menu-item--iconic'),
    );
    expect(mergeItems.map((item) => item.textContent)).toEqual([
      'Merge & Center',
      'Merge Across',
      'Merge cells',
      'Unmerge Cells',
    ]);

    const mergeCenter = host.querySelector<HTMLButtonElement>(
      '#menu-merge [data-merge-action="mergeCenter"]',
    );
    const mergeEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(mergeEvent, 'target', { value: mergeCenter });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(mergeEvent)).toBe(true);
    expect(
      sheet.instance.store.getState().merges.byAnchor.get(addrKey({ sheet: 0, row: 0, col: 0 })),
    ).toEqual({
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 1,
      c1: 1,
    });

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'merge', menuId: 'menu-merge' },
      mergeButton as HTMLButtonElement,
    );
    const unmerge = host.querySelector<HTMLButtonElement>(
      '#menu-merge [data-merge-action="unmergeCells"]',
    );
    const unmergeEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(unmergeEvent, 'target', { value: unmerge });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(unmergeEvent)).toBe(true);
    expect(sheet.instance.store.getState().merges.byAnchor.size).toBe(0);

    tb.dispose();
  });

  it('skips huge Merge Across selections before iterating each row', () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1048575, c1: 1 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const mergeButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="merge"]');
    expect(mergeButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'merge', menuId: 'menu-merge' },
      mergeButton as HTMLButtonElement,
    );
    const mergeAcross = host.querySelector<HTMLButtonElement>(
      '#menu-merge [data-merge-action="mergeAcross"]',
    );
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: mergeAcross });

    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(sheet.instance.store.getState().merges.byAnchor.size).toBe(0);

    tb.dispose();
  });

  it('keeps Home Merge and Find rendering in the audited Excel 365 ribbon layout', () => {
    const helpers = stubHelpers();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: {
        ...helpers,
        createIcon: (name) => {
          const icon = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
          icon.setAttribute('class', 'demo__rb-icon');
          icon.dataset.icon = name;
          return icon;
        },
      },
    });

    const mergeButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="merge"]');
    expect(mergeButton).toBeTruthy();
    expect(mergeButton?.closest('.demo__ribbon-group')?.classList).toContain(
      'demo__ribbon-group--alignment',
    );
    expect(mergeButton?.classList.contains('demo__rb--wide')).toBe(false);
    expect(mergeButton?.dataset.ribbonActivation).toBe('splitPrimary');
    expect(mergeButton?.querySelector('.demo__rb-icon')?.getAttribute('data-icon')).toBe('merge');
    expect(
      Array.from(mergeButton?.querySelectorAll('span') ?? []).filter(
        (span) => !span.classList.contains('demo__rb-split-chevron'),
      ),
    ).toEqual([]);

    const findButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="findHome"]');
    expect(findButton).toBeTruthy();
    expect(findButton?.closest('.demo__ribbon-group')?.classList).toContain(
      'demo__ribbon-group--editing',
    );
    expect(findButton?.classList.contains('demo__rb--wide')).toBe(true);
    expect(findButton?.dataset.ribbonActivation).toBe('dropdown');
    expect(findButton?.dataset.ribbonMenuId).toBe('menu-find-select');
    expect(findButton?.querySelector('.demo__rb-icon')?.getAttribute('data-icon')).toBe('find');
    expect(findButton?.querySelector('span')?.textContent).toBe('Find & Select');
    expect(findButton?.querySelector('.demo__rb-split-chevron')).toBeTruthy();

    tb.dispose();
  });

  it('projects Home dense groups through shared layout classes', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });

    for (const variant of HOME_TILE_LAYOUT_GROUP_VARIANTS) {
      const group = host.querySelector<HTMLElement>(`.demo__ribbon-group--${variant}`);
      expect(group, variant).toBeTruthy();
      expect(group?.classList.contains('demo__ribbon-group--tiles'), variant).toBe(true);
      const nonWideCommands = Array.from(
        group?.querySelectorAll<HTMLButtonElement>('[data-ribbon-command]') ?? [],
      )
        .filter((button) => !button.classList.contains('demo__rb--wide'))
        .map((button) => button.dataset.ribbonCommand);
      expect(nonWideCommands, variant).toEqual([]);
    }
    for (const variant of HOME_STACKED_LAYOUT_GROUP_VARIANTS) {
      const group = host.querySelector<HTMLElement>(`.demo__ribbon-group--${variant}`);
      expect(group, variant).toBeTruthy();
      expect(group?.classList.contains('demo__ribbon-group--stacked'), variant).toBe(true);
      expect(
        Array.from(group?.querySelectorAll<HTMLButtonElement>('[data-ribbon-command]') ?? []).map(
          (button) => button.classList.contains('demo__rb--stacked'),
        ),
        variant,
      ).toEqual([true, true, true]);
    }
    for (const variant of HOME_MIXED_LAYOUT_GROUP_VARIANTS) {
      const group = host.querySelector<HTMLElement>(`.demo__ribbon-group--${variant}`);
      expect(group, variant).toBeTruthy();
      expect(group?.classList.contains('demo__ribbon-group--mixed'), variant).toBe(true);
      expect(
        Array.from(group?.querySelectorAll<HTMLButtonElement>('[data-ribbon-command]') ?? []).map(
          (button) => button.classList.contains('demo__rb--stacked'),
        ),
        variant,
      ).toEqual([true, true, true, false, false]);
    }

    tb.dispose();
  });

  it('applies Freeze Panes dropdown actions through default dynamic wiring', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 3, col: 2 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('view');

    const freezeButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="freeze"]');
    expect(freezeButton).toBeTruthy();
    if (freezeButton) {
      Object.defineProperty(freezeButton, 'getBoundingClientRect', {
        value: () =>
          ({
            x: 120,
            y: 32,
            width: 40,
            height: 28,
            top: 32,
            right: 160,
            bottom: 60,
            left: 120,
            toJSON: () => ({}),
          }) as DOMRect,
      });
    }
    freezeButton?.click();
    const freezeMenu = host.querySelector<HTMLElement>('#menu-freeze');
    expect(freezeMenu?.style.position).toBe('fixed');
    expect(freezeMenu?.style.left).toBe('120px');
    expect(freezeMenu?.style.top).toBe('63px');
    document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    expect(freezeMenu?.hidden).toBe(true);
    freezeButton?.click();
    const freezeItems = Array.from(
      host.querySelectorAll<HTMLButtonElement>('#menu-freeze .app__menu-item--iconic'),
    );
    expect(freezeItems).toHaveLength(3);
    expect(freezeItems.map((item) => item.textContent)).toEqual([
      'Freeze Panes',
      'Freeze Top Row',
      'Freeze First Column',
    ]);
    expect(host.querySelector('#menu-freeze [data-freeze="off"]')).toBeNull();
    const selectionFreeze = host.querySelector<HTMLButtonElement>(
      '#menu-freeze [data-freeze="selection"]',
    );
    expect(selectionFreeze).toBeTruthy();
    const freezeEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(freezeEvent, 'target', { value: selectionFreeze });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(freezeEvent)).toBe(true);
    expect(sheet.instance.store.getState().layout.freezeRows).toBe(3);
    expect(sheet.instance.store.getState().layout.freezeCols).toBe(2);

    freezeButton?.click();
    const unfreeze = host.querySelector<HTMLButtonElement>('#menu-freeze [data-freeze="off"]');
    expect(unfreeze).toBeTruthy();
    expect(host.querySelectorAll('#menu-freeze .app__menu-item--iconic').length).toBe(3);
    expect(unfreeze?.textContent).toBe('Unfreeze Panes');
    expect(unfreeze?.disabled).toBe(false);
    expect(unfreeze?.getAttribute('aria-disabled')).toBe('false');
    const unfreezeEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(unfreezeEvent, 'target', { value: unfreeze });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(unfreezeEvent)).toBe(true);
    expect(sheet.instance.store.getState().layout.freezeRows).toBe(0);
    expect(sheet.instance.store.getState().layout.freezeCols).toBe(0);

    tb.dispose();
  });

  it('clamps dynamic dropdowns to the viewport from the shared opener positioning', () => {
    Object.defineProperty(window, 'innerWidth', { configurable: true, value: 640 });
    Object.defineProperty(window, 'innerHeight', { configurable: true, value: 360 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('view');

    const freezeButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="freeze"]');
    const freezeMenu = host.querySelector<HTMLElement>('#menu-freeze');
    expect(freezeButton).toBeTruthy();
    expect(freezeMenu).toBeTruthy();
    if (freezeButton) {
      Object.defineProperty(freezeButton, 'getBoundingClientRect', {
        value: () =>
          ({
            x: 580,
            y: 310,
            width: 40,
            height: 28,
            top: 310,
            right: 620,
            bottom: 338,
            left: 580,
            toJSON: () => ({}),
          }) as DOMRect,
      });
    }
    if (freezeMenu) {
      Object.defineProperty(freezeMenu, 'offsetWidth', { configurable: true, value: 216 });
      Object.defineProperty(freezeMenu, 'offsetHeight', { configurable: true, value: 420 });
    }

    freezeButton?.click();

    expect(freezeMenu?.style.position).toBe('fixed');
    expect(freezeMenu?.style.left).toBe('416px');
    expect(freezeMenu?.style.top).toBe('8px');
    expect(freezeMenu?.style.maxHeight).toBe('299px');
    expect(freezeMenu?.style.overflowY).toBe('auto');
    expect(freezeMenu?.style.overscrollBehavior).toBe('contain');

    tb.dispose();
  });

  it('opens External Links from primary click and keeps hyperlink actions secondary', () => {
    const originalOpen = window.open;
    const open = vi.fn();
    window.open = open;
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setHyperlink(
      sheet.instance.store,
      { sheet: 0, row: 0, col: 0 },
      'https://example.test',
      sheet.workbook,
    );
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('data');

    const linksButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="linksData"]');
    expect(linksButton).toBeTruthy();
    linksButton?.click();
    const linksDialog = document.body.querySelector<HTMLElement>('.fc-extlinkdlg');
    expect(linksDialog?.hidden).toBe(false);
    linksDialog
      ?.querySelector<HTMLButtonElement>('.fc-extlinkdlg__close')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(linksDialog?.hidden).toBe(true);

    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'linksData',
      menuId: 'menu-links-data',
    });
    const linksMenu = host.querySelector<HTMLElement>('#menu-links-data');
    expect(linksMenu?.querySelectorAll('.app__menu-item--iconic').length).toBe(4);
    const openButton = linksMenu?.querySelector<HTMLButtonElement>('[data-link-action="open"]');
    const clearButton = linksMenu?.querySelector<HTMLButtonElement>('[data-link-action="clear"]');
    expect(openButton?.disabled).toBe(false);
    expect(clearButton).toBeTruthy();
    expect(clearButton?.disabled).toBe(false);
    const openEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(openEvent, 'target', { value: openButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(openEvent)).toBe(true);
    expect(open).toHaveBeenCalledWith('https://example.test', '_blank', 'noopener,noreferrer');

    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'linksData',
      menuId: 'menu-links-data',
    });
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: clearButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);

    expect(hyperlinkAt(sheet.instance.store.getState(), { sheet: 0, row: 0, col: 0 })).toBeNull();

    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'linksData',
      menuId: 'menu-links-data',
    });
    expect(openButton?.disabled).toBe(true);
    expect(openButton?.dataset.menuDisabledReason).toBe(
      'The active cell does not contain a hyperlink.',
    );
    expect(clearButton?.disabled).toBe(true);
    expect(clearButton?.dataset.menuDisabledReason).toBe(
      'The active cell does not contain a hyperlink.',
    );

    tb.dispose();
    window.open = originalOpen;
  });

  it('routes Find & Select dropdown actions through shared dialogs, reports, and selection', async () => {
    seedNumber(sheet, 0, 0, 1);
    seedText(sheet, 1, 1, 'plain');
    sheet.workbook.setFormula({ sheet: 0, row: 2, col: 2 }, '=A1+1');
    sheet.instance.store.setState((state) => {
      const cells = new Map(state.data.cells);
      cells.set(addrKey({ sheet: 0, row: 2, col: 2 }), {
        value: { kind: 'number', value: 2 },
        formula: '=A1+1',
      });
      return { ...state, data: { ...state.data, cells } };
    });
    const openFindReplace = vi
      .spyOn(sheet.instance, 'openFindReplace')
      .mockImplementation(() => undefined);
    const openGoTo = vi.spyOn(sheet.instance, 'openGoTo').mockImplementation(() => undefined);
    const openGoToSpecial = vi
      .spyOn(sheet.instance, 'openGoToSpecial')
      .mockImplementation(() => undefined);
    const openWorkbookObjects = vi
      .spyOn(sheet.instance, 'openWorkbookObjects')
      .mockImplementation(() => undefined);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const findButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="findHome"]');
    expect(findButton).toBeTruthy();
    const clickFindSelect = async (action: string): Promise<void> => {
      tb.dropdownsApi?.openDynamicRibbonDropdown(
        { command: 'findHome', menuId: 'menu-find-select' },
        findButton as HTMLButtonElement,
      );
      const button = host.querySelector<HTMLButtonElement>(`[data-find-select="${action}"]`);
      expect(button).toBeTruthy();
      const event = new MouseEvent('click', { bubbles: true });
      Object.defineProperty(event, 'target', { value: button });
      expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
      await Promise.resolve();
    };

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'findHome', menuId: 'menu-find-select' },
      findButton as HTMLButtonElement,
    );
    expect(host.querySelectorAll('#menu-find-select .app__menu-item--iconic').length).toBe(11);

    await clickFindSelect('find');
    await clickFindSelect('replace');
    await clickFindSelect('go-to');
    await clickFindSelect('go-to-special');
    expect(openFindReplace).toHaveBeenNthCalledWith(1, 'find');
    expect(openFindReplace).toHaveBeenNthCalledWith(2, 'replace');
    expect(openGoTo).toHaveBeenCalledTimes(1);
    expect(openGoToSpecial).toHaveBeenCalledTimes(1);

    await clickFindSelect('object-select');
    await clickFindSelect('selection-pane');
    expect(openWorkbookObjects).toHaveBeenCalledTimes(2);

    await clickFindSelect('conditional-format');
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'No matching cells were found.',
    );
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

    await clickFindSelect('comments');
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'No comments or notes were found.',
    );
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

    await clickFindSelect('formulas');

    expect(sheet.instance.store.getState().selection.active).toEqual({
      sheet: 0,
      row: 2,
      col: 2,
    });

    tb.dispose();
  });

  it('opens Name Manager from primary click and keeps Define Name secondary', async () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('formulas');

    const namesButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="namedRanges"]',
    );
    expect(namesButton).toBeTruthy();
    namesButton?.click();
    await new Promise((resolve) => requestAnimationFrame(resolve));
    const manager = document.body.querySelector<HTMLElement>('.fc-namedlg');
    expect(manager?.hidden).toBe(false);
    expect(manager?.querySelector<HTMLElement>('.fc-namedlg__list')).toBeTruthy();

    Array.from(manager?.querySelectorAll<HTMLButtonElement>('button') ?? [])
      .find((button) => button.textContent === 'Close')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(manager?.hidden).toBe(true);

    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'namedRanges',
      menuId: 'menu-defined-names',
    });
    const namesMenu = host.querySelector<HTMLElement>('#menu-defined-names');
    expect(namesMenu?.querySelectorAll('.app__menu-item--iconic').length).toBe(7);
    expect(
      namesMenu?.querySelector<HTMLButtonElement>('[data-defined-name-action="use-formula"]')
        ?.disabled,
    ).toBe(true);
    for (const createButton of namesMenu?.querySelectorAll<HTMLButtonElement>(
      '[data-defined-name-action^="create-"]',
    ) ?? []) {
      expect(createButton.disabled).toBe(!sheet.workbook.capabilities.definedNameMutate);
    }
    vi.spyOn(sheet.workbook, 'definedNames').mockImplementation(function* () {
      yield { name: 'TaxRate', formula: '=Sheet1!$A$1', localSheetId: -1 };
      yield { name: 'NetSales', formula: '=Sheet1!$B$1', localSheetId: -1 };
    });
    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'namedRanges',
      menuId: 'menu-defined-names',
    });
    const useFormulaButton = namesMenu?.querySelector<HTMLButtonElement>(
      '[data-defined-name-action="use-formula"]',
    );
    expect(useFormulaButton?.disabled).toBe(false);
    const useFormulaEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(useFormulaEvent, 'target', { value: useFormulaButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(useFormulaEvent)).toBe(true);
    await Promise.resolve();
    const useFormulaDialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(useFormulaDialog?.textContent).toContain('Use in Formula');
    expect(useFormulaDialog?.textContent).toContain('TaxRate');
    expect(useFormulaDialog?.textContent).toContain('NetSales');
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBeNull();
    const netSalesRadio =
      useFormulaDialog?.querySelector<HTMLInputElement>('input[value="NetSales"]');
    expect(netSalesRadio).toBeTruthy();
    netSalesRadio?.click();
    useFormulaDialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=NetSales');

    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'namedRanges',
      menuId: 'menu-defined-names',
    });
    const defineButton = namesMenu?.querySelector<HTMLButtonElement>(
      '[data-defined-name-action="define"]',
    );
    expect(defineButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: defineButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    await new Promise((resolve) => requestAnimationFrame(resolve));

    expect(document.body.querySelector<HTMLElement>('.fc-namedlg')?.hidden).toBe(false);

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
});
