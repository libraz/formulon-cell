import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { PivotAxis, PivotReportLayout } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { dictionaries } from '../../../src/i18n/strings.js';
import { createPivotDropdownDefaults } from '../../../src/mount/dynamic-dropdown-defaults/pivot.js';
import type { SpreadsheetInstance } from '../../../src/mount/types.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import { createPivotMenuFactories } from '../../../src/toolbar/ribbon/menus/pivot.js';
import { loadFixture } from '../engine/pivot-excel-parity/fixture.js';

// ja_c0_rows.xlsx: PivotTable1 at A3:C7 on the first sheet.
describe('PivotTable Design and Analyze menu defaults', () => {
  let wb: WorkbookHandle;
  let instance: SpreadsheetInstance;

  beforeEach(async () => {
    wb = await loadFixture('ja_c0_rows.xlsx');
    const store = createSpreadsheetStore();
    mutators.replaceCells(store, wb.cells(0));
    mutators.setActive(store, { sheet: 0, row: 4, col: 1 });
    const host = document.createElement('div');
    document.body.appendChild(host);
    instance = {
      host,
      store,
      workbook: wb,
      history: new History(),
      i18n: { strings: dictionaries.ja },
    } as unknown as SpreadsheetInstance;
  });

  afterEach(() => {
    document.body.replaceChildren();
    wb.dispose();
  });

  it('applies grand totals, layout and subtotal position, each one undo step', () => {
    const { applyPivotLayoutAction } = createPivotDropdownDefaults(instance);
    applyPivotLayoutAction('totals-rows');
    expect(wb.getPivotDefinition(0, 0)?.grandTotals).toEqual({ rows: true, cols: false });
    applyPivotLayoutAction('layout-tabular');
    expect(wb.getPivotDefinition(0, 0)?.layout).toBe(PivotReportLayout.Tabular);
    applyPivotLayoutAction('subtotals-bottom');
    expect(
      wb
        .getPivotDefinition(0, 0)
        ?.fields.some(
          (f) => (f.axis === PivotAxis.Row || f.axis === PivotAxis.Col) && f.subtotalTop,
        ),
    ).toBe(false);

    expect(instance.history.undo()).toBe(true);
    expect(wb.getPivotDefinition(0, 0)?.layout).toBe(PivotReportLayout.Tabular);
    expect(instance.history.undo()).toBe(true);
    expect(wb.getPivotDefinition(0, 0)?.layout).not.toBe(PivotReportLayout.Tabular);
    expect(instance.history.undo()).toBe(true);
    expect(wb.getPivotDefinition(0, 0)?.grandTotals).toEqual({ rows: true, cols: true });
  });

  it('selects the whole report', () => {
    createPivotDropdownDefaults(instance).applyPivotLayoutAction('select-table');
    expect(instance.store.getState().selection.range).toMatchObject({ r0: 2, c0: 0, r1: 6, c1: 2 });
  });

  it('disables Clear Filters while nothing is filtered', () => {
    const menu = createPivotMenuFactories({
      ribbonMenuText: dictionaries.ja.ribbonMenu,
    }).createPivotLayoutMenu('pivotClear');
    createPivotDropdownDefaults(instance).updatePivotLayoutMenu(menu);
    const clearFilters = menu.querySelector<HTMLButtonElement>(
      '[data-pivot-layout="clear-filters"]',
    );
    expect(clearFilters?.getAttribute('aria-disabled')).toBe('true');
    expect(
      menu.querySelector('[data-pivot-layout="clear-all"]')?.getAttribute('aria-disabled'),
    ).not.toBe('true');
  });
});
