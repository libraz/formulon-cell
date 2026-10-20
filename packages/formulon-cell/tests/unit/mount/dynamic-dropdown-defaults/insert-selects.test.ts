import { expect, it } from 'vitest';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createIllustrationDropdownDefaults } from '../../../../src/mount/dynamic-dropdown-defaults/illustrations.js';
import { createInsertDropdownDefaults } from '../../../../src/mount/dynamic-dropdown-defaults/insert.js';
import { mutators } from '../../../../src/store/store.js';
import { mountStubSheet } from '../../../test-utils/mount.js';

it('selects and focuses the chart and the shape the ribbon inserts', async () => {
  const workbook = await WorkbookHandle.createDefault();
  const sheet = await mountStubSheet({ workbook });
  try {
    const { store } = sheet.instance;
    workbook.setText({ sheet: 0, row: 0, col: 1 }, 'North');
    workbook.setText({ sheet: 0, row: 1, col: 0 }, 'Jan');
    workbook.setNumber({ sheet: 0, row: 1, col: 1 }, 10);
    mutators.replaceCells(store, workbook.cells(0));
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    mutators.extendRangeTo(store, { sheet: 0, row: 1, col: 1 });
    const ribbon = document.createElement('button');
    document.body.appendChild(ribbon);
    ribbon.focus();

    createInsertDropdownDefaults(sheet.instance).createChartFromSelection('column');
    const [chart] = store.getState().charts.charts;
    expect(store.getState().ui.selectedObject?.id).toBe(chart?.id);
    expect((document.activeElement as HTMLElement | null)?.dataset.chartId).toBe(chart?.id);

    ribbon.focus();
    createIllustrationDropdownDefaults(sheet.instance).insertShapeFromRibbon('rectangle');
    const [shape] = store.getState().illustrations.illustrations;
    expect(store.getState().ui.selectedObject?.id).toBe(shape?.id);
    expect((document.activeElement as HTMLElement | null)?.dataset.illustrationId).toBe(shape?.id);
    ribbon.remove();
  } finally {
    sheet.dispose();
    workbook.dispose();
  }
});
