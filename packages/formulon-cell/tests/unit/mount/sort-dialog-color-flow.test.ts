import { afterEach, describe, expect, it } from 'vitest';
import { createEditingDropdownDefaults } from '../../../src/mount/dynamic-dropdown-defaults/editing.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

const settle = () => new Promise((resolve) => setTimeout(resolve, 0));

describe('Sort dialog colour level', () => {
  let sheet: MountedStubSheet | null = null;
  afterEach(() => {
    sheet?.dispose();
    sheet = null;
  });

  it('sorts the rows with the picked fill on top', async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    const { instance } = sheet;
    ['Item', 'Pen', 'Ink', 'Pad'].forEach((value, row) => {
      instance.workbook.setText({ sheet: 0, row, col: 0 }, value);
    });
    mutators.replaceCells(instance.store, instance.workbook.cells(0));
    mutators.setCellFormat(instance.store, { sheet: 0, row: 3, col: 0 }, { fill: '#ffff00' });
    mutators.setActive(instance.store, { sheet: 0, row: 1, col: 0 });
    const done = createEditingDropdownDefaults(instance).applySortMenuAction('custom');
    await settle();
    const dialog = document.querySelector('.fc-sortdlg')?.closest('[role="dialog"]');
    const sortOn = dialog?.querySelector<HTMLSelectElement>('.fc-sortdlg__sort-on');
    if (!sortOn) throw new Error('missing Sort On');
    sortOn.value = 'cellColor';
    sortOn.dispatchEvent(new Event('change'));
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await done;
    await settle();
    expect([1, 2, 3].map((row) => instance.workbook.getValue({ sheet: 0, row, col: 0 }))).toEqual([
      { kind: 'text', value: 'Pad' },
      { kind: 'text', value: 'Pen' },
      { kind: 'text', value: 'Ink' },
    ]);
  });
});
