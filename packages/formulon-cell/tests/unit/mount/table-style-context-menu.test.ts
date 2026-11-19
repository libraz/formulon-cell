import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  createTableStyleFromActiveTable,
  customTableStyleId,
  formatAsTable,
  galleryStyleElements,
} from '../../../src/commands/format-as-table.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { applyTableStyle } from '../../../src/commands/table-style-actions.js';
import { builtinTableStyleElements } from '../../../src/engine/builtin-table-styles.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { attachTableStyleContextMenu } from '../../../src/mount/table-style-context-menu.js';
import type { SpreadsheetInstance } from '../../../src/mount/types.js';
import { Spreadsheet } from '../../../src/mount.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';
import { scaledTimeout } from '../../timeouts.js';
import { stubHelpers } from './toolbar/fixtures.js';

vi.setConfig({ testTimeout: scaledTimeout(20_000) });

describe('table style gallery context menu', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;
  let toolbar: ReturnType<typeof Spreadsheet.mountToolbar>;

  beforeEach(async () => {
    const workbook = await WorkbookHandle.createDefault();
    for (const [row, cols] of [
      ['Item', 'Qty'],
      ['A', '1'],
      ['B', '2'],
    ].entries()) {
      for (const [col, text] of cols.entries()) workbook.setText({ sheet: 0, row, col }, text);
    }
    sheet = await mountStubSheet({ locale: 'en', workbook });
    mutators.replaceCells(sheet.instance.store, workbook.cells(0));
    formatAsTable(
      sheet.instance.store,
      { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
      { workbook: sheet.workbook },
    );
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 1, col: 0 });
    host = document.createElement('div');
    document.body.appendChild(host);
    toolbar = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    toolbar.setActiveTab('tableDesign');
  });

  afterEach(() => {
    toolbar.dispose();
    sheet.dispose();
    host.remove();
    document.body.replaceChildren();
  });

  const table = () => sheet.instance.store.getState().tables.tables[0];
  const rightClick = (selector: string): void => {
    const target = host.querySelector<HTMLElement>(selector);
    if (!target) throw new Error(`missing ${selector}`);
    target.dispatchEvent(
      new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 40, clientY: 60 }),
    );
  };
  const items = (): HTMLButtonElement[] =>
    Array.from(host.querySelectorAll<HTMLButtonElement>('.fc-tb__tablestyle-ctx button'));
  const item = (label: string): HTMLButtonElement => {
    const found = items().find((button) => button.textContent === label);
    if (!found) throw new Error(`no ${label}`);
    return found;
  };

  it('offers the desktop items, with Modify and Delete only for a custom style', () => {
    rightClick('[data-table-quick-style="TableStyleMedium3"]');
    expect(items().map((button) => [button.textContent, button.disabled])).toEqual([
      ['Apply and Clear Formatting', false],
      ['Apply (and Maintain Formatting)', false],
      ['Modify...', true],
      ['Duplicate...', false],
      ['Delete', true],
      ['Set As Default', false],
    ]);
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
    expect(items()).toHaveLength(0);
  });

  it('applies a style and clears the cell formatting as one undo step', () => {
    const cell = { sheet: 0, row: 1, col: 0 };
    mutators.setCellFormat(sheet.instance.store, cell, { fill: '#ff0000', bold: true });
    rightClick('[data-table-quick-style="TableStyleMedium3"]');
    item('Apply and Clear Formatting').click();
    expect(table()?.styleName).toBe('TableStyleMedium3');
    expect(sheet.instance.store.getState().format.formats.get('0:1:0')?.fill).toBeUndefined();
    sheet.instance.history.undo();
    expect(table()?.styleName).toBeUndefined();
    expect(sheet.instance.store.getState().format.formats.get('0:1:0')).toMatchObject({
      fill: '#ff0000',
      bold: true,
    });
  });

  it('keeps the cell formatting with Apply (and Maintain Formatting)', () => {
    mutators.setCellFormat(sheet.instance.store, { sheet: 0, row: 1, col: 0 }, { fill: '#ff0000' });
    rightClick('[data-table-quick-style="TableStyleMedium4"]');
    item('Apply (and Maintain Formatting)').click();
    expect(table()?.styleName).toBe('TableStyleMedium4');
    expect(sheet.instance.store.getState().format.formats.get('0:1:0')?.fill).toBe('#ff0000');
  });

  it('sets the workbook default table style', () => {
    rightClick('[data-table-quick-style="TableStyleMedium5"]');
    item('Set As Default').click();
    expect(sheet.workbook.getTableStyleDefaults()?.tableStyle).toBe('TableStyleMedium5');
  });

  it('undoes Set As Default in one step', () => {
    const before = sheet.workbook.getTableStyleDefaults()?.tableStyle;
    rightClick('[data-table-quick-style="TableStyleMedium5"]');
    item('Set As Default').click();
    expect(sheet.instance.history.undo()).toBe(true);
    expect(sheet.workbook.getTableStyleDefaults()?.tableStyle).toBe(before);
  });

  it('deletes a custom style, moving its tables to the default style, as one undo step', () => {
    createTableStyleFromActiveTable(sheet.instance.store, sheet.instance.history, 'Review', {
      elements: galleryStyleElements('dark', '#4472c4', 'banded'),
    });
    const id = customTableStyleId('Review');
    const current = table();
    if (!current) throw new Error('no table');
    applyTableStyle(
      {
        store: sheet.instance.store,
        history: sheet.instance.history,
        origin: 'ribbon',
        commandId: 'tableQuickStyles',
      },
      current,
      id,
    );
    expect(table()?.styleName).toBe('Review');
    // The strip reopens on the row holding the table's style, the custom row.
    toolbar.rerender();
    toolbar.setActiveTab('tableDesign');

    rightClick(`[data-table-quick-style="${id}"]`);
    expect(item('Modify...').disabled).toBe(false);
    item('Delete').click();
    expect(sheet.instance.store.getState().tables.customTableStyles ?? []).toEqual([]);
    expect(table()?.styleName).toBe('TableStyleMedium2');
    sheet.instance.history.undo();
    expect(sheet.instance.store.getState().tables.customTableStyles).toHaveLength(1);
    expect(table()?.styleName).toBe('Review');
  });

  it('greys the strip of a table on a protected sheet and opens no menu on its chips', () => {
    createTableStyleFromActiveTable(sheet.instance.store, sheet.instance.history, 'Review', {
      elements: galleryStyleElements('dark', '#4472c4', 'banded'),
    });
    const id = customTableStyleId('Review');
    const current = table();
    if (!current) throw new Error('no table');
    applyTableStyle(
      {
        store: sheet.instance.store,
        history: sheet.instance.history,
        origin: 'ribbon',
        commandId: 'tableQuickStyles',
      },
      current,
      id,
    );
    setProtectedSheet(sheet.instance.store, 0, true);
    toolbar.rerender();
    toolbar.setActiveTab('tableDesign');
    const chips = host.querySelectorAll<HTMLButtonElement>('[data-table-quick-style]');
    expect(chips.length).toBeGreaterThan(0);
    for (const chip of chips) {
      expect(chip.disabled).toBe(true);
      expect(chip.title).toContain(en.protection.protectedCellEdit);
    }
    const chip = host.querySelector<HTMLElement>(`[data-table-quick-style="${id}"]`);
    const event = new MouseEvent('contextmenu', { bubbles: true, cancelable: true });
    chip?.dispatchEvent(event);
    expect(items()).toHaveLength(0);
    expect(event.defaultPrevented).toBe(true);
  });

  it('deletes a custom style, moving its tables to the workbook default table style', () => {
    sheet.workbook.setTableStyleDefaults({ tableStyle: 'TableStyleLight1' });
    createTableStyleFromActiveTable(sheet.instance.store, sheet.instance.history, 'Review', {
      elements: galleryStyleElements('dark', '#4472c4', 'banded'),
    });
    const id = customTableStyleId('Review');
    const current = table();
    if (!current) throw new Error('no table');
    applyTableStyle(
      {
        store: sheet.instance.store,
        history: sheet.instance.history,
        origin: 'ribbon',
        commandId: 'tableQuickStyles',
      },
      current,
      id,
    );
    expect(table()?.styleName).toBe('Review');
    toolbar.rerender();
    toolbar.setActiveTab('tableDesign');

    toolbar.rerender();
    toolbar.setActiveTab('tableDesign');

    rightClick(`[data-table-quick-style="${id}"]`);
    item('Delete').click();
    expect(table()?.styleName).toBe('TableStyleLight1');
  });

  it('keeps Modify and Delete enabled but inert while another sheet is protected', async () => {
    // Excel for Mac: with any sheet protected, 変更… and 削除 stay enabled and do nothing.
    createTableStyleFromActiveTable(sheet.instance.store, sheet.instance.history, 'Review', {
      elements: galleryStyleElements('dark', '#4472c4', 'banded'),
    });
    const id = customTableStyleId('Review');
    const current = table();
    if (!current) throw new Error('no table');
    // The strip opens on the page of the table's style, so the custom chip shows once applied.
    applyTableStyle(
      {
        store: sheet.instance.store,
        history: sheet.instance.history,
        origin: 'ribbon',
        commandId: 'tableQuickStyles',
      },
      current,
      id,
    );
    setProtectedSheet(sheet.instance.store, 1, true);
    toolbar.rerender();
    toolbar.setActiveTab('tableDesign');
    rightClick(`[data-table-quick-style="${id}"]`);
    expect(item('Modify...').disabled).toBe(false);
    expect(item('Delete').disabled).toBe(false);
    item('Delete').click();
    expect(sheet.instance.store.getState().tables.customTableStyles).toHaveLength(1);
    rightClick(`[data-table-quick-style="${id}"]`);
    item('Modify...').click();
    await Promise.resolve();
    expect(document.body.querySelector('.fc-tb__dlg')).toBeNull();
  });

  it('duplicates a built-in style through the table style dialog', async () => {
    rightClick('[data-table-quick-style="TableStyleMedium3"]');
    item('Duplicate...').click();
    await Promise.resolve();
    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    // Duplicate opens under the Modify title and names the copy after the engine style name.
    expect(dialog?.textContent).toContain('Modify Table Style');
    const name = dialog?.querySelector<HTMLInputElement>('input');
    expect(name?.value).toBe('TableStyleMedium3 2');
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await new Promise((resolve) => setTimeout(resolve, 0));
    const created = sheet.instance.store.getState().tables.customTableStyles ?? [];
    expect(created).toHaveLength(1);
    expect(created[0]?.elements).toEqual(
      builtinTableStyleElements('TableStyleMedium3', sheet.workbook.getTheme()?.colors),
    );
  });

  it('offers no context menu in the expanded galleries', () => {
    toolbar.setActiveTab('home');
    host.querySelector<HTMLButtonElement>('[data-ribbon-command="formatTableHome"]')?.click();
    rightClick('#menu-table-style-home [data-table-style="TableStyleLight2"]');
    expect(items()).toHaveLength(0);
    toolbar.setActiveTab('tableDesign');
    host.querySelector<HTMLButtonElement>('[data-ribbon-command="tableQuickStyles"]')?.click();
    rightClick('#menu-table-quick-styles [data-table-style="TableStyleLight2"]');
    expect(items()).toHaveLength(0);
  });
});

describe('attachTableStyleContextMenu focus restore', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  const setup = () => {
    const sheet = document.createElement('div');
    sheet.tabIndex = -1;
    const strip = document.createElement('div');
    const chip = document.createElement('button');
    chip.dataset.tableQuickStyle = 'TableStyleLight1';
    strip.appendChild(chip);
    document.body.append(sheet, strip);
    const instance = {
      host: sheet,
      store: createSpreadsheetStore(),
      i18n: { strings: en },
      workbook: { capabilities: {} },
    } as unknown as SpreadsheetInstance;
    attachTableStyleContextMenu(
      strip,
      () => instance,
      () => {},
    );
    chip.dispatchEvent(
      new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 5, clientY: 5 }),
    );
    const menu = document.querySelector<HTMLElement>('.fc-tb__tablestyle-ctx');
    if (!menu) throw new Error('context menu did not open');
    return { sheet, chip, menu };
  };

  const pressEscape = (menu: HTMLElement): void => {
    menu.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }),
    );
  };

  it('returns focus to the connected chip on Escape', () => {
    const { chip, menu } = setup();
    pressEscape(menu);
    expect(document.activeElement).toBe(chip);
  });

  it('falls back to the sheet when the chip is detached', () => {
    const { sheet, chip, menu } = setup();
    chip.remove();
    pressEscape(menu);
    expect(document.activeElement).toBe(sheet);
  });
});
