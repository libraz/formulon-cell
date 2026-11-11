import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  createTableStyleFromActiveTable,
  customTableStyleId,
  formatAsTable,
  updateTableOverlay,
} from '../../../src/commands/format-as-table.js';
import { Spreadsheet } from '../../../src/mount.js';
import { mutators } from '../../../src/store/store.js';
import type { MountedStubSheet } from '../../test-utils/mount.js';
import { mountPivotSheet, stubHelpers } from './toolbar/fixtures.js';

vi.setConfig({ testTimeout: 20_000 });

const TABLE_CELL = { sheet: 0, row: 21, col: 0 };
const PIVOT_CELL = { sheet: 0, row: 9, col: 4 };

// The table and PivotTable galleries mark the active object's current style, in the Quick
// Styles strip and in the full menu, and show Clear only inside one.
describe('current style in the table and PivotTable galleries', () => {
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
    host = document.createElement('div');
    document.body.appendChild(host);
    toolbar = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
  });

  afterEach(() => {
    toolbar.dispose();
    sheet.dispose();
    host.remove();
    document.body.replaceChildren();
  });

  const store = () => sheet.instance.store;
  const checked = (menuId: string, attr: string): string[] =>
    [...(host.querySelector(`#${menuId}`)?.querySelectorAll<HTMLElement>(`[${attr}]`) ?? [])]
      .filter((item) => item.getAttribute('aria-checked') === 'true')
      .map((item) => item.getAttribute(attr) ?? '');
  const clearShown = (menuId: string, footer: string): boolean =>
    host.querySelector<HTMLElement>(`#${menuId} [data-table-style-footer="${footer}"]`)?.hidden ===
    false;
  /** The pressed chips across every page of a Quick Styles strip. */
  const pressed = (attr: string): string[] => {
    const gallery = host.querySelector<HTMLElement>(
      `[data-ribbon-gallery="${attr === 'data-table-quick-style' ? 'tableQuickStyles' : 'pivotQuickStyles'}"]`,
    );
    if (!gallery) throw new Error('no gallery');
    gallery.dispatchEvent(new Event('fc-ribbon-control-refresh'));
    const step = (suffix: string) =>
      gallery.querySelector<HTMLButtonElement>(`[class$="-${suffix}"]`) as HTMLButtonElement;
    while (step('prev').getAttribute('aria-disabled') !== 'true') step('prev').click();
    const found: string[] = [];
    for (;;) {
      for (const chip of gallery.querySelectorAll<HTMLElement>(`[${attr}][aria-pressed="true"]`))
        found.push(chip.getAttribute(attr) ?? '');
      if (step('next').getAttribute('aria-disabled') === 'true') return found;
      step('next').click();
    }
  };
  const open = (command: string): void =>
    host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`)?.click();

  it('marks a table whose style names a custom style in another case', () => {
    mutators.setActive(store(), TABLE_CELL);
    expect(
      createTableStyleFromActiveTable(store(), sheet.instance.history, 'Brand', {
        elements: { headerRow: { bold: true } },
      }),
    ).toBe(true);
    const table = store().getState().tables.tables[0];
    if (!table) throw new Error('no table');
    updateTableOverlay(store(), table.id, { styleName: 'brand' });
    toolbar.setActiveTab('tableDesign');
    const brand = customTableStyleId('Brand');
    expect(pressed('data-table-quick-style')).toEqual([brand]);
    open('tableQuickStyles');
    expect(checked('menu-table-quick-styles', 'data-table-style')).toEqual([brand]);
    expect(clearShown('menu-table-quick-styles', 'clear-table-style')).toBe(true);
  });

  it('marks the PivotTable style picked from the menu', () => {
    mutators.setActive(store(), PIVOT_CELL);
    toolbar.setActiveTab('pivotDesign');
    open('pivotQuickStyles');
    const item = host.querySelector<HTMLElement>(
      '#menu-pivot-quick-styles [data-pivot-table-style="PivotStyleMedium9"]',
    );
    if (!item) throw new Error('no PivotStyleMedium9');
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: item });
    expect(toolbar.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(pressed('data-pivot-style')).toEqual(['PivotStyleMedium9']);
    open('pivotQuickStyles');
    expect(checked('menu-pivot-quick-styles', 'data-pivot-table-style')).toEqual([
      'PivotStyleMedium9',
    ]);
    expect(clearShown('menu-pivot-quick-styles', 'clear-pivot-style')).toBe(true);
  });
});
