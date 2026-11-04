import { expect, it } from 'vitest';
import { setHyperlink } from '../../../../src/commands/hyperlinks.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createInsertDropdownDefaults } from '../../../../src/mount/dynamic-dropdown-defaults/insert.js';
import { mutators } from '../../../../src/store/store.js';
import { mountStubSheet } from '../../../test-utils/mount.js';

it('reports the localized unsupported target through the mounted link action', async () => {
  const sheet = await mountStubSheet({ locale: 'ja' });
  try {
    setHyperlink(
      sheet.instance.store,
      { sheet: 0, row: 0, col: 0 },
      'javascript:void(0)',
      sheet.workbook,
    );
    const action = createInsertDropdownDefaults(sheet.instance).applyLinksAction('open');
    const report = document.body
      .querySelector<HTMLElement>('.fc-tb__dlg__list')
      ?.closest<HTMLElement>('[role="dialog"]');
    expect(report).toBeTruthy();
    expect(report?.textContent).toContain(
      sheet.instance.i18n.strings.ribbonMenu.linkUnsupportedTarget,
    );
    expect(report?.textContent).not.toContain(
      sheet.instance.i18n.strings.ribbonMenu.linkNoHyperlink,
    );
    report?.querySelector<HTMLButtonElement>('.fc-fmtdlg__footer button')?.click();
    await action;
  } finally {
    document.body
      .querySelector<HTMLElement>('[role="dialog"]')
      ?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    sheet.dispose();
  }
});

it('inserts the recommended PivotTable directly on the Mac profile', async () => {
  const workbook = await WorkbookHandle.createDefault();
  const rows: (string | number)[][] = [
    ['Region', 'Amt'],
    ['East', 1],
    ['West', 2],
  ];
  for (const [row, cells] of rows.entries()) {
    for (const [col, value] of cells.entries()) {
      if (typeof value === 'number') workbook.setNumber({ sheet: 0, row, col }, value);
      else workbook.setText({ sheet: 0, row, col }, value);
    }
  }
  const sheet = await mountStubSheet({ locale: 'ja', workbook });
  try {
    mutators.replaceCells(sheet.instance.store, workbook.cells(0));
    sheet.instance.host.dataset.fcPlatform = 'mac';
    const sheets = workbook.sheetCount;
    await createInsertDropdownDefaults(sheet.instance).applyPivotTableAction('recommended');
    expect(workbook.sheetCount).toBe(sheets + 1);
    expect(workbook.getPivotTables()).toHaveLength(1);
    // No report dialog: the field list pane opens on the new report instead.
    expect(document.body.querySelector('.fc-tb__dlg')).toBeNull();
  } finally {
    sheet.dispose();
  }
});
