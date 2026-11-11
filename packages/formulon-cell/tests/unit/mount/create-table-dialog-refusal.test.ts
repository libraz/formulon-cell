import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { formatAsTable } from '../../../src/commands/format-as-table.js';
import { addrKey } from '../../../src/engine/address.js';
import { showCreateTableDialog } from '../../../src/mount/dynamic-dropdown-defaults/styles.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

describe('Create Table dialog that creates nothing', () => {
  let sheet: MountedStubSheet;
  beforeEach(async () => {
    sheet = await mountStubSheet();
  });
  afterEach(() => sheet.dispose());

  it('keeps the formatting of a range it refuses to turn into a table', async () => {
    const { store, history } = sheet.instance;
    formatAsTable(store, { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 });
    mutators.setCellFormat(store, { sheet: 0, row: 1, col: 0 }, { fill: '#ff0000' });
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 4, r1: 3, c1: 5 });
    mutators.setActive(store, { sheet: 0, row: 0, col: 4 });

    const pending = showCreateTableDialog(sheet.instance, { clearFormatting: true });
    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    const input = dialog?.querySelector<HTMLInputElement>('input');
    if (!input) throw new Error('dialog did not open');
    input.value = 'A2:B3';
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await pending;

    expect(store.getState().tables.tables).toHaveLength(1);
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 0 }))?.fill).toBe(
      '#ff0000',
    );
    expect(history.canUndo()).toBe(false);
  });
});
