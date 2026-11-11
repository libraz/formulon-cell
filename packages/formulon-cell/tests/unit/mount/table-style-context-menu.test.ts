import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  createTableStyleFromActiveTable,
  customTableStyleId,
  formatAsTable,
} from '../../../src/commands/format-as-table.js';
import { applyTableStyle } from '../../../src/commands/table-style-actions.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { Spreadsheet } from '../../../src/mount.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';
import { stubHelpers } from './toolbar/fixtures.js';

vi.setConfig({ testTimeout: 20_000 });

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
      ['Apply (and Maintain Formatting)', false],
      ['Apply and Clear Formatting', false],
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

  it('deletes a custom style, moving its tables to the default style, as one undo step', () => {
    createTableStyleFromActiveTable(sheet.instance.store, sheet.instance.history, 'Review', {
      style: 'dark',
      color: '#4472c4',
      variant: 'banded',
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

  it('duplicates a built-in style through the table style dialog', async () => {
    rightClick('[data-table-quick-style="TableStyleMedium3"]');
    item('Duplicate...').click();
    await Promise.resolve();
    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(dialog?.textContent).toContain('New Table Style');
    const name = dialog?.querySelector<HTMLInputElement>('input');
    expect(name?.value).toMatch(/ 2$/);
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    expect(sheet.instance.store.getState().tables.customTableStyles).toHaveLength(1);
  });

  it('creates a table with cleared formatting when no table is active', async () => {
    const cell = { sheet: 0, row: 5, col: 4 };
    sheet.workbook.setText(cell, 'x');
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setCellFormat(sheet.instance.store, cell, { fill: '#00ff00' });
    mutators.setActive(sheet.instance.store, cell);
    toolbar.setActiveTab('home');
    host.querySelector<HTMLButtonElement>('[data-ribbon-command="formatTableHome"]')?.click();
    rightClick('#menu-table-style-home [data-table-style="TableStyleLight2"]');
    item('Apply and Clear Formatting').click();
    await Promise.resolve();
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    const tables = sheet.instance.store.getState().tables.tables;
    expect(tables.at(-1)).toMatchObject({ styleName: 'TableStyleLight2' });
    expect(sheet.instance.store.getState().format.formats.get('0:5:4')?.fill).toBeUndefined();
    sheet.instance.history.undo();
    expect(sheet.instance.store.getState().tables.tables).toHaveLength(1);
    expect(sheet.instance.store.getState().format.formats.get('0:5:4')?.fill).toBe('#00ff00');
  });
});
