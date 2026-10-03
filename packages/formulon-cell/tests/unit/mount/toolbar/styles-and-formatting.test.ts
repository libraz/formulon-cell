import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { customCellStyleId } from '../../../../src/commands/cell-styles.js';
import {
  customPivotTableStyleId,
  customTableStyleId,
} from '../../../../src/commands/format-as-table.js';
import { addrKey } from '../../../../src/engine/address.js';
import type { MountToolbarOptions } from '../../../../src/mount/toolbar.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators } from '../../../../src/store/store.js';
import { RIBBON_BORDERS_MENU_ID } from '../../../../src/toolbar/ribbon/activation.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { seedNumber, stubHelpers, waitFor } from './fixtures.js';

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

  it('projects active formatting onto ribbon toggle buttons', () => {
    mutators.setCellFormat(sheet.instance.store, { sheet: 0, row: 0, col: 0 }, { bold: true });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });

    const bold = host.querySelector<HTMLButtonElement>('[data-ribbon-command="bold"]');
    expect(bold).toBeTruthy();
    expect(bold?.classList.contains('fc-tb__rb--active')).toBe(true);
    expect(bold?.getAttribute('aria-pressed')).toBe('true');

    mutators.setActive(sheet.instance.store, { sheet: 0, row: 1, col: 0 });
    expect(bold?.classList.contains('fc-tb__rb--active')).toBe(false);
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

    expect(bold?.classList.contains('fc-tb__rb--active')).toBe(true);
    expect(bold?.getAttribute('aria-pressed')).toBe('true');
    expect(sheet.instance.store.getState().ui.pendingFormat).toEqual({
      addr: { sheet: 0, row: 0, col: 0 },
      format: { bold: true },
    });
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')).toBeUndefined();

    tb.dispose();
  });

  it('retains the logical command id through a deferred custom format callback', () => {
    const seen: string[] = [];
    const deferred: Parameters<NonNullable<MountToolbarOptions['applyRibbonFormat']>>[0][] = [];
    sheet.instance.commands.setPolicy({
      operations: { format: true },
      editable: () => true,
      restrict: ({ intent }) => {
        seen.push(`${intent.origin}:${intent.commandId ?? ''}`);
        return intent.origin === 'ribbon' && intent.commandId === 'bold';
      },
    });
    sheet.instance.store.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        extraRanges: [{ sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 }],
      },
    }));
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
      applyRibbonFormat: (fn) => deferred.push(fn),
    });

    expect(tb.applyCommand('bold')).toBe(true);
    expect(deferred).toHaveLength(1);
    deferred[0]?.(sheet.instance.store.getState(), sheet.instance.store);

    expect(seen.filter((entry) => entry === 'ribbon:bold').length).toBeGreaterThanOrEqual(2);
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')?.bold).toBe(true);
    expect(sheet.instance.store.getState().format.formats.get('0:0:3')?.bold).toBe(true);
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
      host.querySelectorAll<HTMLButtonElement>('#menu-underline .fc-tb__menu-item--iconic'),
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
    expect(sheet.instance.store.getState().ui.pendingFormat).toEqual({
      addr: { sheet: 0, row: 0, col: 0 },
      format: { underline: 'double' },
    });

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
        host.querySelectorAll<HTMLButtonElement>('#menu-underline .fc-tb__menu-item--iconic'),
      ).map((item) => item.textContent),
    ).toEqual(['下線', '二重下線']);

    const double = host.querySelector<HTMLButtonElement>(
      '#menu-underline [data-underline-action="double"]',
    );
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: double });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(sheet.instance.store.getState().ui.pendingFormat).toEqual({
      addr: { sheet: 0, row: 0, col: 0 },
      format: { underline: 'double' },
    });

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
    expect(menuTexts('#menu-paste .fc-tb__menu-item--iconic')).toEqual([
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
    expect(menuTexts('#menu-conditional > .fc-tb__menu-item')).toEqual([
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
        host.querySelectorAll<HTMLElement>('#menu-table-style-home .fc-tb__tablestyle-heading'),
      ).map((heading) => heading.textContent),
    ).toEqual(['淡色', '中間', '濃色']);
    expect(menuTexts('#menu-table-style-home > .fc-tb__tablestyle-footer')).toEqual([
      '新しい表スタイル…',
      '新しいピボットテーブル スタイル…',
    ]);

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="findHome"]')?.click();
    expect(menuTexts('#menu-find-select .fc-tb__menu-item--iconic')).toEqual([
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
    expect(menu?.classList.contains('fc-tb__tablestyle-menu')).toBe(true);
    const scrollBody = menu?.querySelector<HTMLElement>(':scope > .fc-tb__tablestyle-scroll');
    expect(scrollBody?.getAttribute('role')).toBe('group');
    expect(scrollBody?.getAttribute('aria-label')).toBe('Format as Table');
    const headings = Array.from(
      scrollBody?.querySelectorAll<HTMLElement>(':scope > .fc-tb__tablestyle-heading') ?? [],
    ).map((heading) => heading.textContent);
    expect(headings).toEqual(['Light', 'Medium', 'Dark']);
    const grids = Array.from(
      scrollBody?.querySelectorAll<HTMLElement>(':scope > .fc-tb__tablestyle-grid') ?? [],
    );
    expect(grids).toHaveLength(3);
    expect(grids.map((grid) => grid.getAttribute('aria-label'))).toEqual([
      'Light',
      'Medium',
      'Dark',
    ]);
    expect(grids.map((grid) => grid.querySelectorAll('.fc-tb__tablestyle-swatch').length)).toEqual([
      28, 28, 7,
    ]);
    const footerActions = Array.from(
      menu?.querySelectorAll<HTMLButtonElement>(':scope > .fc-tb__tablestyle-footer') ?? [],
    );
    expect(footerActions.map((button) => button.dataset.tableStyleFooter)).toEqual([
      'new-table-style',
      'new-pivot-style',
    ]);
    expect(
      footerActions.every((button) => button.classList.contains('fc-tb__menu-item--iconic')),
    ).toBe(true);
    expect(menu?.querySelector('.fc-tb__menu-icon--table-style-new')).toBeTruthy();
    expect(menu?.querySelector('.fc-tb__menu-icon--pivot-style-new')).toBeTruthy();
    const styleButton = host.querySelector<HTMLButtonElement>(
      '[data-table-style="dark"][data-table-color="#4472c4"][data-table-variant="banded"]',
    );
    expect(styleButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: styleButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    await Promise.resolve();

    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    const rangeInput = dialog?.querySelector<HTMLInputElement>('input[type="text"]');
    expect(document.body.textContent).toContain('Create Table');
    expect(rangeInput?.value).toBe('Sheet1!$A$1:$C$4');
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
    expect(activeStyleButton?.classList.contains('fc-tb__menu-item--active')).toBe(true);

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

    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    const rangeInput = dialog?.querySelector<HTMLInputElement>('input[type="text"]');
    await waitFor(() => document.activeElement === rangeInput);
    expect(rangeInput?.value).toBe('Sheet1!$E$5');
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
    const scrollBody = menu?.querySelector<HTMLElement>(':scope > .fc-tb__cellstyle-scroll');
    expect(scrollBody?.getAttribute('role')).toBe('group');
    expect(scrollBody?.getAttribute('aria-label')).toBe('Cell styles');
    expect(
      Array.from(
        scrollBody?.querySelectorAll<HTMLElement>(':scope > .fc-tb__cellstyle-heading') ?? [],
      ).map((heading) => heading.textContent),
    ).toEqual([
      'Good, Bad and Neutral',
      'Data and Model',
      'Titles and Headings',
      'Themed Cell Styles',
      'Number Format',
    ]);
    expect(
      menu?.querySelector<HTMLElement>(':scope > .fc-tb__cellstyle-footer')?.parentElement,
    ).toBe(menu);
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
    expect(goodButton?.classList.contains('fc-tb__menu-item--active')).toBe(true);

    tb.dispose();
  });

  it('renders the Mac 40% accent palette in the Home Cell Styles menu', () => {
    sheet.instance.host.dataset.fcPlatform = 'mac';
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const stylesButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="cellStyles"]',
    );
    stylesButton?.click();
    const accent = host.querySelector<HTMLButtonElement>('[data-cell-style="accent1_40"]');
    expect(accent).toBeTruthy();
    expect(accent?.style.background).toBe('#83cceb');
    expect(host.querySelectorAll('[data-cell-style]').length).toBe(47);
    tb.dispose();
  });

  it('samples the owner platform again for the next ribbon style action', () => {
    sheet.instance.host.dataset.fcPlatform = 'mac';
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const stylesButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="cellStyles"]',
    );
    stylesButton?.click();
    let accent = host.querySelector<HTMLButtonElement>('[data-cell-style="accent1_40"]');
    expect(accent?.style.background).toBe('#83cceb');
    let event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: accent });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')?.fill).toBe('#83cceb');

    sheet.instance.host.dataset.fcPlatform = 'default';
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 });
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 1 });
    stylesButton?.click();
    accent = host.querySelector<HTMLButtonElement>('[data-cell-style="accent1_40"]');
    expect(accent?.style.background).toBe('#b4c7e7');
    event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: accent });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(sheet.instance.store.getState().format.formats.get('0:0:1')?.fill).toBe('#b4c7e7');
    tb.dispose();
  });

  it('renders Mac non-accent fallback colors in the ribbon style menu', () => {
    sheet.instance.host.dataset.fcPlatform = 'mac';
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const stylesButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="cellStyles"]',
    );
    stylesButton?.click();
    expect(host.querySelector<HTMLButtonElement>('[data-cell-style="title"]')?.style.color).toBe(
      '#0e2841',
    );
    expect(host.querySelector<HTMLButtonElement>('[data-cell-style="heading1"]')?.style.color).toBe(
      '#0e2841',
    );
    const check = host.querySelector<HTMLButtonElement>('[data-cell-style="checkCell"]');
    expect(check?.style.color).toBe('#ffffff');
    expect(check?.style.background).toBe('#a5a5a5');
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
    const tableStyleDialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
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
    const pivotStyleDialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
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
    expect(customPivotStyleButton?.classList.contains('fc-tb__menu-item--active')).toBe(true);

    const stylesButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="cellStyles"]',
    );
    expect(stylesButton).toBeTruthy();
    stylesButton?.click();
    expect(
      host.querySelector('#menu-cell-styles-home .fc-tb__menu-icon--cell-style-new'),
    ).toBeTruthy();
    expect(
      host.querySelector('#menu-cell-styles-home .fc-tb__menu-icon--cell-style-merge'),
    ).toBeTruthy();
    const newCellStyleButton = host.querySelector<HTMLButtonElement>(
      '[data-cell-style-footer="new-cell-style"]',
    );
    expect(newCellStyleButton).toBeTruthy();
    const cellEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(cellEvent, 'target', { value: newCellStyleButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(cellEvent)).toBe(true);
    await Promise.resolve();
    const cellStyleDialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
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
    expect(host.querySelector('#menu-currency-home .fc-tb__menu-icon-svg')).toBeFalsy();
    expect(host.querySelectorAll('#menu-currency-home .fc-tb__menu-item__icon-spacer').length).toBe(
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
    expect(activeEurButton?.classList.contains('fc-tb__menu-item--active')).toBe(true);
    expect(inactiveUsdButton?.getAttribute('aria-checked')).toBe('false');

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
    expect(ccwButton?.classList.contains('fc-tb__menu-item--active')).toBe(true);

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
});
