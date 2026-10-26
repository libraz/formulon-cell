import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { formatAsTable } from '../../../../src/commands/format-as-table.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { stubHelpers } from './fixtures.js';

vi.setConfig({ testTimeout: 20_000 });

describe('Table Design tab of the ribbon', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;

  beforeEach(async () => {
    const workbook = await WorkbookHandle.createDefault();
    for (const [row, cols] of [
      ['Item', 'Region', 'Qty'],
      ['A', 'East', '1'],
      ['B', 'West', '2'],
      ['C', 'East', '3'],
    ].entries()) {
      for (const [col, text] of cols.entries()) {
        if (col === 2 && row > 0) workbook.setNumber({ sheet: 0, row, col }, Number(text));
        else workbook.setText({ sheet: 0, row, col }, text);
      }
    }
    sheet = await mountStubSheet({ locale: 'en', workbook });
    mutators.replaceCells(sheet.instance.store, workbook.cells(0));
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    sheet.dispose();
    host.remove();
  });

  const store = () => sheet.instance.store;
  const tabIds = (): string[] =>
    Array.from(host.querySelectorAll<HTMLElement>('[data-ribbon-tab]')).map(
      (tab) => tab.dataset.ribbonTab ?? '',
    );
  const nameInput = (): HTMLInputElement => {
    const found = host.querySelector<HTMLInputElement>('[data-ribbon-input="tableName"] input');
    if (!found) throw new Error('missing table name box');
    return found;
  };
  const command = (id: string): HTMLButtonElement => {
    const found = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${id}"]`);
    if (!found) throw new Error(`missing ${id}`);
    return found;
  };
  const commitName = (value: string): void => {
    nameInput().value = value;
    nameInput().dispatchEvent(new Event('change', { bubbles: true }));
  };
  const table = () => store().getState().tables.tables[0];
  const makeTable = (): void => {
    formatAsTable(store(), { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 2 }, { workbook: sheet.workbook });
  };

  it('shows only while the active cell is in a table, and returns to the previous tab', () => {
    mutators.setActive(store(), { sheet: 0, row: 1, col: 1 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    expect(tabIds()).not.toContain('tableDesign');
    makeTable();
    expect(tabIds().at(-1)).toBe('tableDesign');
    tb.setActiveTab('tableDesign');
    mutators.setActive(store(), { sheet: 0, row: 8, col: 8 });
    expect(tabIds()).not.toContain('tableDesign');
    expect(tb.getActiveTab()).toBe('home');
    tb.dispose();
  });

  it('shows the table name and renames the table from the box', () => {
    makeTable();
    mutators.setActive(store(), { sheet: 0, row: 1, col: 1 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    expect(nameInput().value).toBe('Table1');
    commitName('Sales');
    expect(table()?.name).toBe('Sales');
    expect(sheet.workbook.getTables().map((t) => t.name)).toEqual(['Sales']);
    expect(nameInput().value).toBe('Sales');
    sheet.instance.undo();
    expect(table()?.name).toBe('Table1');
    tb.dispose();
  });

  it('reports an invalid name and keeps the old one', async () => {
    makeTable();
    mutators.setActive(store(), { sheet: 0, row: 1, col: 1 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    commitName('not valid');
    await Promise.resolve();
    expect(document.body.textContent).toContain('The table name is not valid.');
    expect(table()?.name).toBe('Table1');
    expect(nameInput().value).toBe('Table1');
    tb.dispose();
  });

  it('reports a duplicate name', async () => {
    makeTable();
    formatAsTable(store(), { sheet: 0, r0: 6, c0: 0, r1: 8, c1: 1 }, { workbook: sheet.workbook });
    mutators.setActive(store(), { sheet: 0, row: 1, col: 1 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    commitName('table2');
    await Promise.resolve();
    expect(document.body.textContent).toContain('The table name is already in use.');
    expect(table()?.name).toBe('Table1');
    tb.dispose();
  });

  it('toggles the total row, banded rows and banded columns from the option buttons', () => {
    makeTable();
    mutators.setActive(store(), { sheet: 0, row: 1, col: 1 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    expect(command('tableBandedRows').getAttribute('aria-pressed')).toBe('true');
    expect(command('tableTotalRow').getAttribute('aria-pressed')).toBe('false');

    command('tableHeaderRow').click();
    expect(table()?.showHeader).toBe(false);
    expect(command('tableHeaderRow').getAttribute('aria-pressed')).toBe('false');
    command('tableHeaderRow').click();
    expect(table()?.showHeader).toBe(true);

    command('tableTotalRow').click();
    expect(table()?.showTotal).toBe(true);
    expect(sheet.workbook.getTableStyleDetails(0)?.totalsRowCount).toBe(1);
    expect(command('tableTotalRow').getAttribute('aria-pressed')).toBe('true');

    command('tableBandedColumns').click();
    command('tableBandedRows').click();
    expect(table()).toMatchObject({ colStripes: true, banded: false });
    expect(sheet.workbook.getTableStyleDetails(0)).toMatchObject({
      showColumnStripes: true,
      showRowStripes: false,
    });
    sheet.instance.undo();
    expect(table()?.banded).toBe(true);
    tb.dispose();
  });

  it('disables the Table Design controls for a table that is not stored in the workbook', () => {
    formatAsTable(store(), { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 2 });
    mutators.setActive(store(), { sheet: 0, row: 1, col: 1 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    expect(table()?.source).toBe('session');
    expect(nameInput().disabled).toBe(true);
    for (const id of ['tableTotalRow', 'tableBandedRows', 'tableHeaderRow']) {
      expect(command(id).disabled, id).toBe(true);
      expect(command(id).getAttribute('aria-disabled')).toBe('true');
      expect(command(id).title).toContain('not stored in the workbook');
    }
    tb.dispose();
  });

  it('keeps the table options when a gallery style is picked', () => {
    makeTable();
    mutators.setActive(store(), { sheet: 0, row: 1, col: 1 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    command('tableBandedColumns').click();
    command('tableBandedRows').click();
    command('tableFirstColumn').click();
    expect(table()).toMatchObject({ banded: false, colStripes: true, firstCol: true });

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'formatTableHome', menuId: 'menu-table-style-home' },
      host.querySelector<HTMLButtonElement>('[data-ribbon-command="formatTableHome"]'),
    );
    const swatch = host.querySelector<HTMLButtonElement>(
      '[data-table-style="medium"][data-table-color="#ed7d31"][data-table-variant="banded"]',
    );
    expect(swatch).toBeTruthy();
    const click = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(click, 'target', { value: swatch });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(click)).toBe(true);

    expect(table()).toMatchObject({
      color: '#ed7d31',
      banded: false,
      colStripes: true,
      firstCol: true,
    });
    tb.dispose();
  });

  it('labels the total row in the sheet language', async () => {
    sheet.dispose();
    const workbook = await WorkbookHandle.createDefault();
    for (const [row, cols] of [
      ['Item', 'Qty'],
      ['A', '1'],
      ['B', '2'],
    ].entries()) {
      for (const [col, text] of cols.entries()) {
        if (col === 1 && row > 0) workbook.setNumber({ sheet: 0, row, col }, Number(text));
        else workbook.setText({ sheet: 0, row, col }, text);
      }
    }
    sheet = await mountStubSheet({ locale: 'ja', workbook });
    mutators.replaceCells(store(), workbook.cells(0));
    formatAsTable(store(), { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 }, { workbook });
    mutators.setActive(store(), { sheet: 0, row: 1, col: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    command('tableTotalRow').click();
    expect(workbook.getValue({ sheet: 0, row: 3, col: 0 })).toMatchObject({ value: '集計' });
    tb.dispose();
  });
});
