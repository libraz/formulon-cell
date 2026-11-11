import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { formatAsTable } from '../../../src/commands/format-as-table.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { mutators } from '../../../src/store/store.js';
import { openMacSlicerDialog } from '../../../src/toolbar/ribbon/mac/actions.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

describe('Insert Slicer through the shared dialog', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    const workbook = await WorkbookHandle.createDefault();
    for (const [row, cols] of [
      ['Item', 'Region'],
      ['A', 'East'],
      ['B', 'West'],
    ].entries()) {
      for (const [col, text] of cols.entries()) workbook.setText({ sheet: 0, row, col }, text);
    }
    sheet = await mountStubSheet({ locale: 'en', workbook, features: { slicer: true } });
    mutators.replaceCells(sheet.instance.store, workbook.cells(0));
    formatAsTable(
      sheet.instance.store,
      { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
      { workbook: sheet.workbook },
    );
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 1, col: 0 });
  });

  afterEach(() => sheet.dispose());

  const slicers = () => sheet.instance.store.getState().slicers.slicers;
  const checkAll = (): void => {
    for (const box of document.querySelectorAll<HTMLInputElement>(
      '.fc-macslicerdlg__field input[type="checkbox"]',
    )) {
      box.click();
    }
  };
  const confirm = (): void =>
    document.querySelector<HTMLButtonElement>('.fc-macslicerdlg .fc-fmtdlg__btn--primary')?.click();

  it('adds every checked field as one undo step', () => {
    openMacSlicerDialog(sheet.instance, { commandId: 'tableInsertSlicer' });
    checkAll();
    confirm();
    expect(slicers()).toHaveLength(2);
    sheet.instance.undo();
    expect(slicers()).toHaveLength(0);
  });

  it('adds none when one of the slicers is refused', () => {
    const original = sheet.instance.addSlicer.bind(sheet.instance);
    let calls = 0;
    sheet.instance.addSlicer = (input) => {
      calls += 1;
      if (calls === 2) throw new Error('refused');
      return original(input);
    };
    openMacSlicerDialog(sheet.instance, { commandId: 'tableInsertSlicer' });
    checkAll();
    confirm();
    expect(calls).toBe(2);
    expect(slicers()).toHaveLength(0);
    expect(document.querySelector('[data-mac-slicer-error]')?.textContent).toContain(
      'could not be inserted',
    );
  });
});
