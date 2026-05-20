import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { customCellStyleId } from '../../../src/commands/cell-styles.js';
import { commentAt, setComment } from '../../../src/commands/comment.js';
import { customTableStyleId } from '../../../src/commands/format-as-table.js';
import { hyperlinkAt, setHyperlink } from '../../../src/commands/hyperlinks.js';
import { addrKey } from '../../../src/engine/address.js';
import { Spreadsheet } from '../../../src/mount.js';
import { getPageSetup, mutators } from '../../../src/store/store.js';
import {
  DYNAMIC_RIBBON_DROPDOWN_HANDLER_DATASET_KEYS,
  type DynamicDropdownsCtx,
  RIBBON_DROPDOWN_MENU_FOR_COMMAND,
} from '../../../src/toolbar/ribbon/dynamic-dropdowns.js';
import type { RibbonRenderHelpers } from '../../../src/toolbar/ribbon/render-ribbon.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

// Minimal helpers stub: enough for the renderer to emit a shell, no real
// dropdown DOM. The toolbar still needs `createSelect/Color/Icon/makeSvg`
// because every command path may reach them.
const stubHelpers = (): RibbonRenderHelpers => ({
  createSelect: () => document.createElement('div'),
  createColor: () => document.createElement('div'),
  createIcon: () => null,
  makeSvg: () => document.createElementNS('http://www.w3.org/2000/svg', 'svg'),
  chevronPath: 'M0 0',
});

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
  applyPivotTableAction: vi.fn(),
  applyDefinedNameAction: vi.fn(),
  applyLinksAction: vi.fn(),
  applyFillSeries: vi.fn(),
  applyFillDirection: vi.fn(),
  applyClearAction: vi.fn(),
  applyFreezeAction: vi.fn(),
  applyTextOrientationAction: vi.fn(),
  applyCellInsertAction: vi.fn(),
  applyCellDeleteAction: vi.fn(),
  applyCellFormatAction: vi.fn(),
  applyPageBreakAction: vi.fn(),
  applySheetBackgroundAction: vi.fn(),
  applyPrintAreaAction: vi.fn(),
  applyPrintTitlesAction: vi.fn(),
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
});

const EXTERNALLY_WIRED_MENU_IDS = new Set(['menu-borders']);

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
    expect(setPrintAreaButton).toBeTruthy();
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
    const shiftDownButton = host.querySelector<HTMLButtonElement>(
      '[data-cell-insert="shift-down"]',
    );
    expect(shiftDownButton).toBeTruthy();
    const insertEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(insertEvent, 'target', { value: shiftDownButton });
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
    const shiftUpButton = host.querySelector<HTMLButtonElement>('[data-cell-delete="shift-up"]');
    expect(shiftUpButton).toBeTruthy();
    const deleteEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(deleteEvent, 'target', { value: shiftUpButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(deleteEvent)).toBe(true);
    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({
      kind: 'number',
      value: 10,
    });

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
    dedupeDialog?.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter' }));
    await Promise.resolve();
    await Promise.resolve();

    expect(sheet.workbook.getValue({ sheet: 0, row: 2, col: 0 }).kind).toBe('blank');

    tb.dispose();
  });

  it('splits text through the Data Text to Columns dropdown', () => {
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
    const commaButton = host.querySelector<HTMLButtonElement>(
      '[data-text-to-columns-delimiter=","]',
    );
    expect(commaButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: commaButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);

    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'alpha',
    });
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
      kind: 'text',
      value: 'beta',
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
    clearButton?.click();
    const clearCommentsButton = host.querySelector<HTMLButtonElement>('[data-clear="comments"]');
    expect(clearCommentsButton).toBeTruthy();
    const commentsEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(commentsEvent, 'target', { value: clearCommentsButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(commentsEvent)).toBe(true);
    expect(commentAt(sheet.instance.store.getState(), { sheet: 0, row: 0, col: 0 })).toBeNull();
    expect(hyperlinkAt(sheet.instance.store.getState(), { sheet: 0, row: 0, col: 0 })).toBe(
      'https://example.test',
    );

    clearButton?.click();
    const clearHyperlinksButton = host.querySelector<HTMLButtonElement>(
      '[data-clear="remove-hyperlinks"]',
    );
    expect(clearHyperlinksButton).toBeTruthy();
    const hyperlinksEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(hyperlinksEvent, 'target', { value: clearHyperlinksButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(hyperlinksEvent)).toBe(true);
    expect(hyperlinkAt(sheet.instance.store.getState(), { sheet: 0, row: 0, col: 0 })).toBeNull();

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

  it('circles and clears data validation state through the Data Validation dropdown', () => {
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
    const circleButton = host.querySelector<HTMLButtonElement>(
      '[data-validation-action="circle-invalid"]',
    );
    expect(circleButton).toBeTruthy();
    const circleEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(circleEvent, 'target', { value: circleButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(circleEvent)).toBe(true);
    expect(sheet.instance.store.getState().errorIndicators.validationCircles).toEqual(
      new Set([addrKey({ sheet: 0, row: 1, col: 0 })]),
    );

    validationButton?.click();
    const clearCirclesButton = host.querySelector<HTMLButtonElement>(
      '[data-validation-action="clear-circles"]',
    );
    expect(clearCirclesButton).toBeTruthy();
    const clearCirclesEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(clearCirclesEvent, 'target', { value: clearCirclesButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(clearCirclesEvent)).toBe(true);
    expect(sheet.instance.store.getState().errorIndicators.validationCircles.size).toBe(0);

    validationButton?.click();
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
    const input = dialog?.querySelector<HTMLInputElement>('input[type="number"]');
    expect(input).toBeTruthy();
    if (!input) throw new Error('Expected conditional formatting number prompt.');
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
      apply: { fill: '#ffc7ce', color: '#9c0006' },
    });

    tb.dispose();
  });

  it('creates a session table through the Home Format as Table dropdown', () => {
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
    const styleButton = host.querySelector<HTMLButtonElement>(
      '[data-table-style="dark"][data-table-color="#4472c4"][data-table-variant="banded"]',
    );
    expect(styleButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: styleButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);

    expect(sheet.instance.store.getState().tables.tables).toMatchObject([
      {
        range: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 2 },
        style: 'dark',
        color: '#4472c4',
        banded: true,
        firstCol: false,
      },
    ]);

    tb.dispose();
  });

  it('applies cell styles through the Home Cell Styles dropdown', () => {
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
    const goodButton = host.querySelector<HTMLButtonElement>('[data-cell-style="good"]');
    expect(goodButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: goodButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);

    expect(
      sheet.instance.store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 1 }))
        ?.cellStyle,
    ).toBe('good');

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

    const stylesButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="cellStyles"]',
    );
    expect(stylesButton).toBeTruthy();
    stylesButton?.click();
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
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const currencyButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="currency"]',
    );
    expect(currencyButton).toBeTruthy();
    currencyButton?.click();
    const eurButton = host.querySelector<HTMLButtonElement>('[data-currency-preset="€"]');
    expect(eurButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: eurButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);

    expect(
      sheet.instance.store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 2 }))
        ?.numFmt,
    ).toEqual({ kind: 'currency', decimals: 2, symbol: '€' });

    tb.dispose();
  });

  it('runs text scripts through the Automate Script dropdown', async () => {
    seedText(sheet, 0, 0, ' alpha ');
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('automate');

    const scriptButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="script"]');
    expect(scriptButton).toBeTruthy();
    scriptButton?.click();
    const trimButton = host.querySelector<HTMLButtonElement>('[data-script-action="trim"]');
    expect(trimButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: trimButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    await Promise.resolve();

    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'alpha',
    });
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      '1 cell(s) changed',
    );
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dispose();
  });

  it('routes PDF and Add-ins dropdowns through their shared reports', async () => {
    const print = vi.spyOn(sheet.instance, 'print').mockImplementation(() => undefined);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('acrobat');

    const pdfButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="pdf"]');
    expect(pdfButton).toBeTruthy();
    pdfButton?.click();
    const createPdfButton = host.querySelector<HTMLButtonElement>('[data-pdf-action="create"]');
    expect(createPdfButton).toBeTruthy();
    const pdfEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(pdfEvent, 'target', { value: createPdfButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(pdfEvent)).toBe(true);
    await Promise.resolve();
    expect(print).toHaveBeenCalledWith('pdf');
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'PDF export has been sent',
    );
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

    const addInButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="addIn"]');
    expect(addInButton).toBeTruthy();
    addInButton?.click();
    const manageButton = host.querySelector<HTMLButtonElement>('[data-add-in-action="manage"]');
    expect(manageButton).toBeTruthy();
    const addInEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(addInEvent, 'target', { value: manageButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(addInEvent)).toBe(true);
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
    const onlineButton = host.querySelector<HTMLButtonElement>('[data-picture-insert="online"]');
    expect(onlineButton).toBeTruthy();
    const pictureEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(pictureEvent, 'target', { value: onlineButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(pictureEvent)).toBe(true);
    await Promise.resolve();
    const pictureDialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(pictureDialog?.textContent).toContain('Image URL');
    const input = pictureDialog?.querySelector<HTMLInputElement>('input');
    expect(input).toBeTruthy();
    if (!input) throw new Error('Expected picture URL input.');
    input.value = 'https://example.test/picture.png';
    pictureDialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    await Promise.resolve();
    expect(sheet.instance.store.getState().illustrations.illustrations).toMatchObject([
      {
        kind: 'image',
        src: 'https://example.test/picture.png',
        sheet: 0,
        w: 240,
        h: 160,
      },
    ]);

    const toDataUrl = vi
      .spyOn(HTMLCanvasElement.prototype, 'toDataURL')
      .mockReturnValue('data:image/png;base64,current-view');
    const screenshotButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="screenshotInsert"]',
    );
    expect(screenshotButton).toBeTruthy();
    screenshotButton?.click();
    const currentViewButton = host.querySelector<HTMLButtonElement>(
      '[data-screenshot-insert="current-view"]',
    );
    expect(currentViewButton).toBeTruthy();
    const screenshotEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(screenshotEvent, 'target', { value: currentViewButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(screenshotEvent)).toBe(true);
    expect(sheet.instance.store.getState().illustrations.illustrations).toMatchObject([
      {
        kind: 'image',
        src: 'https://example.test/picture.png',
      },
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
        src: 'https://example.test/picture.png',
      },
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
    const arrowButton = host.querySelector<HTMLButtonElement>('[data-shape-insert="arrow"]');
    expect(arrowButton).toBeTruthy();
    const shapeEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(shapeEvent, 'target', { value: arrowButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(shapeEvent)).toBe(true);
    expect(sheet.instance.store.getState().illustrations.illustrations).toMatchObject([
      {
        kind: 'image',
        src: 'https://example.test/picture.png',
      },
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
        shape: 'arrow',
        sheet: 0,
        w: 180,
        h: 80,
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

  it('updates protection through the Review Protect dropdown', () => {
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
    protectButton?.click();
    const protectSheetButton = host.querySelector<HTMLButtonElement>(
      '[data-protect-action="protect-sheet"]',
    );
    expect(protectSheetButton).toBeTruthy();
    const protectEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(protectEvent, 'target', { value: protectSheetButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(protectEvent)).toBe(true);
    expect(sheet.instance.store.getState().protection.protectedSheets.has(0)).toBe(true);

    protectButton?.click();
    const workbookButton = host.querySelector<HTMLButtonElement>(
      '[data-protect-action="protect-workbook"]',
    );
    expect(workbookButton).toBeTruthy();
    const workbookEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(workbookEvent, 'target', { value: workbookButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(workbookEvent)).toBe(true);
    expect(sheet.instance.store.getState().protection.workbookStructure).toEqual({});

    protectButton?.click();
    const allowRangeButton = host.querySelector<HTMLButtonElement>(
      '[data-protect-action="allow-edit-ranges"]',
    );
    expect(allowRangeButton).toBeTruthy();
    const allowEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(allowEvent, 'target', { value: allowRangeButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(allowEvent)).toBe(true);
    expect(sheet.instance.store.getState().protection.allowedEditRanges).toHaveLength(1);
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dispose();
  });

  it('deletes comments through the Review Comments dropdown', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    setComment(sheet.instance.store, { sheet: 0, row: 0, col: 0 }, 'note', sheet.workbook);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('review');

    const commentsButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="deleteCommentReview"]',
    );
    expect(commentsButton).toBeTruthy();
    commentsButton?.click();
    const deleteButton = host.querySelector<HTMLButtonElement>(
      '[data-comment-action="delete-active"]',
    );
    expect(deleteButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: deleteButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);

    expect(commentAt(sheet.instance.store.getState(), { sheet: 0, row: 0, col: 0 })).toBeNull();

    tb.dispose();
  });

  it('creates session charts through the Insert Chart dropdown', () => {
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
    const hideRowsButton = host.querySelector<HTMLButtonElement>('[data-cell-format="hide-rows"]');
    expect(hideRowsButton).toBeTruthy();
    const hideEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(hideEvent, 'target', { value: hideRowsButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(hideEvent)).toBe(true);
    expect(sheet.instance.store.getState().layout.hiddenRows.has(1)).toBe(true);

    formatButton?.click();
    const showRowsButton = host.querySelector<HTMLButtonElement>('[data-cell-format="show-rows"]');
    expect(showRowsButton).toBeTruthy();
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

    tb.dispose();
  });

  it('updates manual page breaks through the Page Layout Breaks dropdown', () => {
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
    const insertRowButton = host.querySelector<HTMLButtonElement>(
      '[data-page-break-action="insert-row"]',
    );
    const insertColButton = host.querySelector<HTMLButtonElement>(
      '[data-page-break-action="insert-col"]',
    );
    expect(insertRowButton).toBeTruthy();
    expect(insertColButton).toBeTruthy();

    const rowEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(rowEvent, 'target', { value: insertRowButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(rowEvent)).toBe(true);

    breaksButton?.click();
    const colEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(colEvent, 'target', { value: insertColButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(colEvent)).toBe(true);

    expect(getPageSetup(sheet.instance.store.getState(), 0).manualPageBreakRows).toEqual([4]);
    expect(getPageSetup(sheet.instance.store.getState(), 0).manualPageBreakCols).toEqual([2]);

    breaksButton?.click();
    const resetButton = host.querySelector<HTMLButtonElement>(
      '[data-page-break-action="reset-all"]',
    );
    expect(resetButton).toBeTruthy();
    const resetEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(resetEvent, 'target', { value: resetButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(resetEvent)).toBe(true);

    expect(getPageSetup(sheet.instance.store.getState(), 0).manualPageBreakRows).toBeUndefined();
    expect(getPageSetup(sheet.instance.store.getState(), 0).manualPageBreakCols).toBeUndefined();

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
    const clearPrecedentsButton = host.querySelector<HTMLButtonElement>(
      '[data-formula-audit-action="clear-precedents"]',
    );
    expect(clearPrecedentsButton).toBeTruthy();
    const clearPrecedentsEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(clearPrecedentsEvent, 'target', { value: clearPrecedentsButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(clearPrecedentsEvent)).toBe(true);
    expect(sheet.instance.store.getState().traces.items).toEqual([dependent]);

    clearArrowsButton?.click();
    const clearDependentsButton = host.querySelector<HTMLButtonElement>(
      '[data-formula-audit-action="clear-dependents"]',
    );
    expect(clearDependentsButton).toBeTruthy();
    const clearDependentsEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(clearDependentsEvent, 'target', { value: clearDependentsButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(clearDependentsEvent)).toBe(true);
    expect(sheet.instance.store.getState().traces.items).toEqual([]);

    tb.dispose();
  });

  it('sets and clears sheet background through the Page Layout Background dropdown', async () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('pageLayout');

    const backgroundButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="sheetBackground"]',
    );
    expect(backgroundButton).toBeTruthy();
    backgroundButton?.click();
    const setBackgroundButton = host.querySelector<HTMLButtonElement>(
      '[data-sheet-background-action="set"]',
    );
    expect(setBackgroundButton).toBeTruthy();

    const setEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(setEvent, 'target', { value: setBackgroundButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(setEvent)).toBe(true);

    const input = document.body.querySelector<HTMLInputElement>('.app__dlg input');
    expect(input).toBeTruthy();
    if (!input) throw new Error('Sheet background prompt input was not rendered');
    input.value = ' https://example.test/background.png ';
    const dialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(dialog).toBeTruthy();
    const okButton = dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    expect(okButton).toBeTruthy();
    dialog?.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter' }));
    await Promise.resolve();
    await Promise.resolve();
    await new Promise((resolve) => setTimeout(resolve, 10));
    expect(document.body.querySelector('.app__dlg')?.textContent).toBeUndefined();

    expect(sheet.instance.store.getState().ui.sheetBackgroundImages.get(0)).toBe(
      'https://example.test/background.png',
    );

    backgroundButton?.click();
    const clearBackgroundButton = host.querySelector<HTMLButtonElement>(
      '[data-sheet-background-action="clear"]',
    );
    expect(clearBackgroundButton).toBeTruthy();
    const clearEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(clearEvent, 'target', { value: clearBackgroundButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(clearEvent)).toBe(true);

    expect(sheet.instance.store.getState().ui.sheetBackgroundImages.has(0)).toBe(false);

    tb.dispose();
  });

  it('sets and clears print titles through the Page Layout Print Titles dropdown', () => {
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
    printTitlesButton?.click();
    const rowsButton = host.querySelector<HTMLButtonElement>('[data-print-titles-action="rows"]');
    expect(rowsButton).toBeTruthy();
    const rowsEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(rowsEvent, 'target', { value: rowsButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(rowsEvent)).toBe(true);

    printTitlesButton?.click();
    const colsButton = host.querySelector<HTMLButtonElement>('[data-print-titles-action="cols"]');
    expect(colsButton).toBeTruthy();
    const colsEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(colsEvent, 'target', { value: colsButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(colsEvent)).toBe(true);

    expect(getPageSetup(sheet.instance.store.getState(), 0).printTitleRows).toBe('2:4');
    expect(getPageSetup(sheet.instance.store.getState(), 0).printTitleCols).toBe('C:E');

    printTitlesButton?.click();
    const clearButton = host.querySelector<HTMLButtonElement>('[data-print-titles-action="clear"]');
    expect(clearButton).toBeTruthy();
    const clearEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(clearEvent, 'target', { value: clearButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(clearEvent)).toBe(true);

    expect(getPageSetup(sheet.instance.store.getState(), 0).printTitleRows).toBeUndefined();
    expect(getPageSetup(sheet.instance.store.getState(), 0).printTitleCols).toBeUndefined();

    tb.dispose();
  });

  it('wires every registered dynamic ribbon dropdown to a rendered button and menu', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    for (const [command, menuId] of Object.entries(RIBBON_DROPDOWN_MENU_FOR_COMMAND)) {
      const button = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`);
      const menu = host.querySelector<HTMLDivElement>(`#${menuId}`);
      expect(button, `${command} button`).toBeTruthy();
      expect(menu, `${command} menu`).toBeTruthy();
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
  });

  it('classifies every rendered ribbon menu under a dispatcher or explicit external owner', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const unowned = Array.from(host.querySelectorAll<HTMLDivElement>('.app__menu'))
      .map((menu) => menu.id)
      .filter(
        (menuId) =>
          !(tb.dropdownsApi?.DYNAMIC_RIBBON_DROPDOWN_IDS.has(menuId) ?? false) &&
          !EXTERNALLY_WIRED_MENU_IDS.has(menuId),
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
    borderButton?.click();
    const menu = host.querySelector<HTMLDivElement>('#menu-borders');
    expect(menu?.hidden).toBe(false);
    host.querySelector<HTMLButtonElement>('#menu-borders [data-border-preset="bottom"]')?.click();

    const fmt = sheet.instance.store
      .getState()
      .format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.borders?.bottom).toEqual({ style: 'thin' });

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
    const borderMenu = host.querySelector<HTMLDivElement>('#menu-borders');
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

  it('opens the PivotTable dialog from the Insert dropdown through default wiring', () => {
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
    const fromRangeButton = host.querySelector<HTMLButtonElement>(
      '[data-pivot-table-action="dialog"]',
    );
    expect(fromRangeButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: fromRangeButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);

    expect(document.body.textContent).toContain('Create PivotTable');

    tb.dispose();
  });

  it('opens the Insert Symbol dropdown and inserts the selected symbol', () => {
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
    symbolButton?.click();
    const menu = host.querySelector<HTMLDivElement>('#menu-symbol');
    expect(menu?.hidden).toBe(false);
    const piButton = Array.from(
      menu?.querySelectorAll<HTMLButtonElement>('[data-symbol]') ?? [],
    ).find((button) => button.dataset.symbol === 'π');
    expect(piButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: piButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'π',
    });

    tb.dispose();
  });

  it('opens More Symbols from the Insert Symbol dropdown and inserts custom text', async () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('insert');

    const symbolButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="symbolInsert"]',
    );
    expect(symbolButton).toBeTruthy();
    symbolButton?.click();
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

  it('applies Freeze Panes dropdown actions through default dynamic wiring', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 3, col: 2 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('view');

    const freezeButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="freeze"]');
    expect(freezeButton).toBeTruthy();
    freezeButton?.click();
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
    const unfreezeEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(unfreezeEvent, 'target', { value: unfreeze });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(unfreezeEvent)).toBe(true);
    expect(sheet.instance.store.getState().layout.freezeRows).toBe(0);
    expect(sheet.instance.store.getState().layout.freezeCols).toBe(0);

    tb.dispose();
  });

  it('clears hyperlinks through the Data Links dropdown', () => {
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
    const clearButton = host.querySelector<HTMLButtonElement>('[data-link-action="clear"]');
    expect(clearButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: clearButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);

    expect(hyperlinkAt(sheet.instance.store.getState(), { sheet: 0, row: 0, col: 0 })).toBeNull();

    tb.dispose();
  });

  it('selects matching cells through the Find & Select dropdown', () => {
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
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const findButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="findHome"]');
    expect(findButton).toBeTruthy();
    findButton?.click();
    const formulasButton = host.querySelector<HTMLButtonElement>('[data-find-select="formulas"]');
    expect(formulasButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: formulasButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);

    expect(sheet.instance.store.getState().selection.active).toEqual({
      sheet: 0,
      row: 2,
      col: 2,
    });

    tb.dispose();
  });

  it('opens Define Name from the Formulas Names dropdown', async () => {
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
    const defineButton = host.querySelector<HTMLButtonElement>(
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
