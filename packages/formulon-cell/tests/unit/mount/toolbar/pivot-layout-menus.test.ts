import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { Spreadsheet } from '../../../../src/mount.js';
import type { MountedStubSheet } from '../../../test-utils/mount.js';
import { mountPivotSheet, stubHelpers } from './fixtures.js';

describe('PivotTable Design and Analyze menus in the mounted ribbon', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;

  beforeEach(async () => {
    sheet = await mountPivotSheet('ja');
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    sheet.dispose();
    host.remove();
  });

  const pivot = () => {
    const wb = sheet.instance.workbook;
    const bounds = wb.pivotBounds(0)[0];
    if (!bounds) throw new Error('no pivot');
    return { wb, index: bounds.pivotIndex };
  };

  it('turns the grand totals off from the 集計 menu, as one undo step', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const button = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="pivotGrandTotals"]',
    );
    expect(button?.disabled).toBe(false);
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'pivotGrandTotals', menuId: 'menu-pivot-grand-totals' },
      button,
    );
    const off = host.querySelector<HTMLButtonElement>('[data-pivot-layout="totals-off"]');
    const click = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(click, 'target', { value: off });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(click)).toBe(true);

    const { wb, index } = pivot();
    expect(wb.getPivotDefinition(0, index)?.grandTotals).toEqual({ rows: false, cols: false });
    expect(sheet.instance.history.undo()).toBe(true);
    expect(wb.getPivotDefinition(0, index)?.grandTotals).toEqual({ rows: true, cols: true });
    tb.dispose();
  });

  it('disables フィルターのクリア while nothing is filtered', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'pivotClear', menuId: 'menu-pivot-clear' },
      host.querySelector<HTMLButtonElement>('[data-ribbon-command="pivotClear"]'),
    );
    expect(
      host.querySelector('[data-pivot-layout="clear-filters"]')?.getAttribute('aria-disabled'),
    ).toBe('true');
    expect(
      host.querySelector('[data-pivot-layout="clear-all"]')?.getAttribute('aria-disabled'),
    ).not.toBe('true');
    tb.dispose();
  });
});
