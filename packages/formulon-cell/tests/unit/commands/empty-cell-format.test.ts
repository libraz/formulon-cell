import { afterEach, describe, expect, it } from 'vitest';
import { toggleBold } from '../../../src/commands/format.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

describe('a format set on an empty cell', () => {
  let sheet: MountedStubSheet | undefined;

  afterEach(() => {
    sheet?.dispose();
    sheet = undefined;
  });

  it('stays on the cell after the selection moves away and back', async () => {
    // Excel: Bold on empty D5, select A1, back to D5 — D5 is still bold.
    sheet = await mountStubSheet({ workbook: await WorkbookHandle.createDefault() });
    const { store, workbook } = sheet.instance;
    const d5 = { sheet: 0, row: 4, col: 3 };
    mutators.setActive(store, d5);
    toggleBold(store.getState(), store);

    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    mutators.replaceCells(store, workbook.cells(0));
    mutators.setActive(store, d5);

    expect(store.getState().format.formats.get('0:4:3')?.bold).toBe(true);
  });
});
