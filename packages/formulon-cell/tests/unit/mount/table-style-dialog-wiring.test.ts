import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  createTableStyleFromActiveTable,
  customTableStyleId,
  formatAsTable,
} from '../../../src/commands/format-as-table.js';
import { openNewTableStyle } from '../../../src/mount/dynamic-dropdown-defaults/styles.js';
import { Spreadsheet } from '../../../src/mount.js';
import { mutators } from '../../../src/store/store.js';
import type { MountedStubSheet } from '../../test-utils/mount.js';
import { mountPivotSheet, stubHelpers } from './toolbar/fixtures.js';

vi.setConfig({ testTimeout: 20_000 });

const settle = () => new Promise((resolve) => setTimeout(resolve, 0));

// The table style dialog is wired to the name checks and the workbook default through the
// gallery footer and the gallery context menu, not only through their helpers.
describe('table style dialog wiring', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;
  let toolbar: ReturnType<typeof Spreadsheet.mountToolbar>;

  beforeEach(async () => {
    sheet = await mountPivotSheet();
    for (const [row, text] of ['Head', 'a', 'b'].entries()) {
      sheet.workbook.setText({ sheet: 0, row: 20 + row, col: 0 }, text);
    }
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    formatAsTable(
      sheet.instance.store,
      { sheet: 0, r0: 20, c0: 0, r1: 22, c1: 0 },
      { workbook: sheet.workbook },
    );
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 21, col: 0 });
    for (const name of ['Brand', 'Brand 2', 'Table Style 1']) {
      createTableStyleFromActiveTable(sheet.instance.store, sheet.instance.history, name, {
        elements: { headerRow: { bold: true } },
      });
    }
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

  const dialog = () => document.body.querySelector<HTMLElement>('.fc-tablestyle-dlg');
  const nameInput = () =>
    dialog()?.querySelector<HTMLInputElement>('input[data-dialog-field="name"]') ?? null;
  const defaultBox = () =>
    dialog()?.querySelector<HTMLInputElement>('input[data-dialog-field="setAsDefault"]') ?? null;
  const ok = () => dialog()?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
  const labels = () =>
    (sheet.instance.store.getState().tables.customTableStyles ?? []).map((style) => style.label);
  /** Right-clicks a strip chip, paging the strip back to the custom styles first. */
  const menuItem = (chip: string, label: string): HTMLButtonElement => {
    const prev = host.querySelector<HTMLButtonElement>('.fc-tb__tablestyle-quick-prev');
    while (prev && prev.getAttribute('aria-disabled') !== 'true') prev.click();
    host
      .querySelector<HTMLElement>(`[data-table-quick-style="${chip}"]`)
      ?.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }));
    const found = [
      ...host.querySelectorAll<HTMLButtonElement>('.fc-tb__tablestyle-ctx button'),
    ].find((button) => button.textContent === label);
    if (!found) throw new Error(`no ${label}`);
    return found;
  };
  const tableDefault = () => sheet.workbook.getTableStyleDefaults()?.tableStyle;

  it('pre-fills a free name and refuses a taken one in any case', async () => {
    const done = openNewTableStyle(sheet.instance, 'table');
    expect(nameInput()?.value).toBe('Table Style 2');
    const input = nameInput();
    if (!input) throw new Error('no name field');
    input.value = 'bRAND';
    ok();
    expect(dialog()?.querySelector('.fc-tb__dlg__error')?.textContent).toBe(
      'This style name already exists.',
    );
    input.value = 'Fresh';
    ok();
    await done;
    expect(labels()).toEqual(['Brand', 'Brand 2', 'Table Style 1', 'Fresh']);
  });

  it('names a duplicate after its source with the first free number', async () => {
    menuItem(customTableStyleId('Brand'), 'Duplicate...').click();
    await settle();
    expect(nameInput()?.value).toBe('Brand 3');
    ok();
    await settle();
    expect(labels()).toContain('Brand 3');
  });

  it('shows the workbook default in Modify and writes the change', async () => {
    sheet.workbook.setTableStyleDefaults({ tableStyle: 'brand' });
    menuItem(customTableStyleId('Brand'), 'Modify...').click();
    await settle();
    expect(defaultBox()?.checked).toBe(true);
    const box = defaultBox();
    if (!box) throw new Error('no default checkbox');
    box.checked = false;
    ok();
    await settle();
    expect(tableDefault()).not.toBe('Brand');
    expect(tableDefault()?.toLowerCase()).not.toBe('brand');

    menuItem(customTableStyleId('Brand 2'), 'Modify...').click();
    await settle();
    expect(defaultBox()?.checked).toBe(false);
    const second = defaultBox();
    if (!second) throw new Error('no default checkbox');
    second.checked = true;
    ok();
    await settle();
    expect(tableDefault()).toBe('Brand 2');
  });
});
