import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { formatAsTable } from '../../../../src/commands/format-as-table.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { scaledTimeout } from '../../../timeouts.js';
import { stubHelpers } from './fixtures.js';

vi.setConfig({ testTimeout: scaledTimeout(20_000) });

describe('Table Design Summarize with PivotTable and Insert Slicer', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;

  beforeEach(async () => {
    const workbook = await WorkbookHandle.createDefault();
    workbook.setText({ sheet: 0, row: 0, col: 0 }, 'Item');
    workbook.setText({ sheet: 0, row: 1, col: 0 }, 'A');
    sheet = await mountStubSheet({ locale: 'en', workbook });
    mutators.replaceCells(sheet.instance.store, workbook.cells(0));
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    sheet.dispose();
    host.remove();
  });

  const command = (id: string): HTMLButtonElement => {
    const found = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${id}"]`);
    if (!found) throw new Error(`missing ${id}`);
    return found;
  };
  const ids = ['tableSummarizePivot', 'tableInsertSlicer'];

  it('greys both for a session-only table and enables them for an engine table', () => {
    const store = sheet.instance.store;
    formatAsTable(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 });
    mutators.setActive(store, { sheet: 0, row: 1, col: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    const titles = ids.map((id) => command(id).title.split('\n')[0]);
    for (const id of ids) {
      expect(command(id).disabled, id).toBe(true);
      expect(command(id).title, id).toContain('not stored in the workbook');
    }
    // The command's own title survives under the reason.
    expect(titles.every((t) => t && t.length > 0)).toBe(true);

    const [first] = store.getState().tables.tables;
    mutators.upsertTableOverlay(store, {
      ...(first as NonNullable<typeof first>),
      source: 'engine',
    });
    for (const [i, id] of ids.entries()) {
      expect(command(id).disabled, id).toBe(false);
      expect(command(id).title, id).toBe(titles[i]);
    }
    tb.dispose();
  });
});
