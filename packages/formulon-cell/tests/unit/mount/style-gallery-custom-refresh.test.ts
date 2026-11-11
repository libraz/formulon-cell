import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  createCellStyleFromActiveFormat,
  customCellStyleId,
} from '../../../src/commands/cell-styles.js';
import {
  customPivotTableStyleId,
  customTableStyleId,
  formatAsTable,
} from '../../../src/commands/format-as-table.js';
import { createCustomStyle } from '../../../src/commands/table-style-actions.js';
import { Spreadsheet } from '../../../src/mount.js';
import { mutators } from '../../../src/store/store.js';
import type { MountedStubSheet } from '../../test-utils/mount.js';
import { mountPivotSheet, stubHelpers } from './toolbar/fixtures.js';

vi.setConfig({ testTimeout: 20_000 });

// The style galleries are built with the custom styles, so one made after the ribbon is up
// must still show without a remount.
describe('a custom style made after the ribbon is up', () => {
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
      {
        workbook: sheet.workbook,
      },
    );
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 21, col: 0 });
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

  const ctx = () => ({ store: sheet.instance.store, history: sheet.instance.history });
  const value = { name: 'Brand', elements: { headerRow: { bold: true } }, stripeSizes: {} };
  const open = (command: string): void =>
    host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`)?.click();

  it('shows in the Table Styles menu', () => {
    expect(
      createCustomStyle(ctx(), sheet.workbook, 'table', { ...value, setAsDefault: false }),
    ).toBe(true);
    open('tableQuickStyles');
    const id = customTableStyleId('Brand');
    expect(
      host.querySelector(`#menu-table-quick-styles [data-table-style="${id}"]`),
    ).not.toBeNull();
  });

  it('shows in the PivotTable Styles menu', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 9, col: 4 });
    toolbar.setActiveTab('pivotDesign');
    expect(
      createCustomStyle(ctx(), sheet.workbook, 'pivot', { ...value, setAsDefault: false }),
    ).toBe(true);
    open('pivotQuickStyles');
    const id = customPivotTableStyleId('Brand');
    expect(
      host.querySelector(`#menu-pivot-quick-styles [data-pivot-table-style="${id}"]`),
    ).not.toBeNull();
  });

  it('shows in the Cell Styles menu', () => {
    toolbar.setActiveTab('home');
    const range = { sheet: 0, r0: 21, c0: 0, r1: 21, c1: 0 };
    expect(
      createCellStyleFromActiveFormat(sheet.instance.store, sheet.instance.history, range, 'Brand'),
    ).toBe(true);
    open('cellStyles');
    const id = customCellStyleId('Brand');
    expect(host.querySelector(`#menu-cell-styles-home [data-cell-style="${id}"]`)).not.toBeNull();
  });
});
