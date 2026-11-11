import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { customPivotTableStyleId } from '../../../src/commands/format-as-table.js';
import { Spreadsheet } from '../../../src/mount.js';
import { mutators } from '../../../src/store/store.js';
import type { MountedStubSheet } from '../../test-utils/mount.js';
import { mountPivotSheet, stubHelpers } from './toolbar/fixtures.js';

vi.setConfig({ testTimeout: 20_000 });

describe('PivotTable Styles gallery', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;
  let toolbar: ReturnType<typeof Spreadsheet.mountToolbar>;

  beforeEach(async () => {
    sheet = await mountPivotSheet();
    mutators.upsertCustomPivotTableStyle(sheet.instance.store, {
      id: customPivotTableStyleId('Brand'),
      label: 'Brand',
      elements: { headerRow: { fill: '#336699' } },
    });
    host = document.createElement('div');
    document.body.appendChild(host);
    toolbar = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    toolbar.setActiveTab('pivotDesign');
  });

  afterEach(() => {
    toolbar.dispose();
    sheet.dispose();
    host.remove();
    document.body.replaceChildren();
  });

  const pivotStyle = () => sheet.workbook.getPivotStyleInfo(0, 0)?.name;
  const menu = () => host.querySelector<HTMLElement>('#menu-pivot-quick-styles');
  const openMenu = (): void => {
    host.querySelector<HTMLButtonElement>('[data-ribbon-command="pivotQuickStyles"]')?.click();
  };
  const menuClick = (selector: string): void => {
    const target = menu()?.querySelector<HTMLElement>(selector);
    if (!target) throw new Error(`missing ${selector}`);
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: target });
    expect(toolbar.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
  };
  const rightClick = (selector: string): void => {
    const target = host.querySelector<HTMLElement>(selector);
    if (!target) throw new Error(`missing ${selector}`);
    target.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }));
  };
  const item = (label: string): HTMLButtonElement => {
    const found = Array.from(
      host.querySelectorAll<HTMLButtonElement>('.fc-tb__tablestyle-ctx button'),
    ).find((button) => button.textContent === label);
    if (!found) throw new Error(`no ${label}`);
    return found;
  };

  it('opens the full gallery from Quick Styles, Custom first, with its footer', () => {
    openMenu();
    expect(menu()?.hidden).toBe(false);
    const headings = Array.from(
      menu()?.querySelectorAll<HTMLElement>('.fc-tb__tablestyle-heading') ?? [],
    ).map((h) => h.textContent);
    expect(headings).toEqual(['Custom', 'Light', 'Medium', 'Dark']);
    // Custom, None, then 28 of each family.
    expect(menu()?.querySelectorAll('[data-pivot-table-style]')).toHaveLength(1 + 1 + 84);
    const footer = Array.from(
      menu()?.querySelectorAll<HTMLButtonElement>('.fc-tb__tablestyle-footer:not([hidden])') ?? [],
    ).map((b) => b.dataset.tableStyleFooter);
    expect(footer).toEqual(['new-pivot-style', 'clear-pivot-style']);
  });

  it('applies a style from the menu and clears it with Clear', () => {
    openMenu();
    menuClick('[data-pivot-table-style="PivotStyleMedium9"]');
    expect(pivotStyle()).toBe('PivotStyleMedium9');
    openMenu();
    menuClick('[data-table-style-footer="clear-pivot-style"]');
    expect(pivotStyle()).toBe('');
  });

  it('offers the context menu on the strip chips only, and deletes a custom one', () => {
    openMenu();
    rightClick(`[data-pivot-table-style="${customPivotTableStyleId('Brand')}"]`);
    expect(document.querySelectorAll('.fc-tb__tablestyle-ctx')).toHaveLength(0);
    menuClick(`[data-pivot-table-style="${customPivotTableStyleId('Brand')}"]`);
    expect(sheet.instance.store.getState().tables.pivotTableStyles).toEqual([
      { sheetIndex: 0, pivotIndex: 0, styleId: customPivotTableStyleId('Brand') },
    ]);
    // The strip turns to the current style, the custom one.
    host
      .querySelector('[data-ribbon-gallery="pivotQuickStyles"]')
      ?.dispatchEvent(new Event('fc-ribbon-control-refresh'));
    rightClick(`[data-pivot-style="${customPivotTableStyleId('Brand')}"]`);
    expect(item('Modify...').disabled).toBe(false);
    item('Delete').click();
    expect(sheet.instance.store.getState().tables.customPivotTableStyles).toEqual([]);
    expect(pivotStyle()).toBe('PivotStyleLight16');
  });

  it('sets a built-in PivotTable style as the default and keeps None out of the edits', () => {
    const prev = host.querySelector<HTMLButtonElement>('.fc-tb__pivotstyle-prev');
    while (prev && prev.getAttribute('aria-disabled') !== 'true') prev.click();
    rightClick('[data-pivot-style="PivotStyleLight1"]');
    expect(item('Modify...').disabled).toBe(true);
    item('Set As Default').click();
    expect(sheet.workbook.getTableStyleDefaults()?.pivotStyle).toBe('PivotStyleLight1');
    rightClick('[data-pivot-style=""]');
    expect(item('Duplicate...').disabled).toBe(true);
    expect(item('Set As Default').disabled).toBe(true);
  });
});
