import { describe, expect, it } from 'vitest';
import { showsHeaderFilterButton } from '../../../src/commands/format-as-table.js';
import { History } from '../../../src/commands/history.js';
import { TABLE_OPTION_COMMANDS, toggleTableOption } from '../../../src/commands/table-design.js';
import { hydrateAutoFilterFromEngine } from '../../../src/engine/auto-filter-sync.js';
import {
  hydrateTableOverlaysFromEngine,
  tableOverlaysFromEngine,
} from '../../../src/engine/table-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createI18nController } from '../../../src/i18n/controller.js';
import type { SpreadsheetInstance } from '../../../src/mount/types.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import { MAC_OBJECT_TAB_DISABLED_COMMAND_IDS } from '../../../src/toolbar/ribbon/mac/object-tab-ids.js';
import { projectTableDesign } from '../../../src/toolbar/ribbon/table-design.js';

const tableRange = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 };

async function tableBook() {
  const wb = await WorkbookHandle.createDefault();
  [
    ['Item', 'Region'],
    ['A', 'East'],
    ['B', 'West'],
    ['C', 'East'],
  ].forEach(([item, region], row) => {
    wb.setText({ sheet: 0, row, col: 0 }, item as string);
    wb.setText({ sheet: 0, row, col: 1 }, region as string);
  });
  expect(wb.createTable(0, { range: 'A1:B4', name: 'Sales', columns: ['Item', 'Region'] })).toBe(0);
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  hydrateTableOverlaysFromEngine(wb, store);
  hydrateAutoFilterFromEngine(wb, store, 0);
  wb.attachStore(store);
  mutators.setActive(store, { sheet: 0, row: 1, col: 0 });
  return { wb, store };
}

describe('Table Design filter button', () => {
  it('is a table option of the ribbon toggle family', () => {
    expect(TABLE_OPTION_COMMANDS.tableFilterButton).toBe('filterButton');
  });

  it('turns the table AutoFilter off and back on, with undo', async () => {
    const { wb, store } = await tableBook();
    const history = new History();
    try {
      const table = () => {
        const overlay = store.getState().tables.tables[0];
        if (!overlay) throw new Error('missing table');
        return overlay;
      };
      expect(table().filterButton).not.toBe(false);
      expect(store.getState().ui.filterRanges).toEqual([tableRange]);
      expect(showsHeaderFilterButton(table(), 0, 0)).toBe(true);

      expect(toggleTableOption(store, history, wb, 'filterButton')).toBe('toggled');
      expect(table().filterButton).toBe(false);
      expect(wb.getTableAutoFilter(0)).toBeNull();
      expect(store.getState().ui.filterRanges).toEqual([]);
      expect(showsHeaderFilterButton(table(), 0, 0)).toBe(false);
      expect(tableOverlaysFromEngine(wb)[0]?.filterButton).toBe(false);

      expect(history.undo()).toBe(true);
      expect(table().filterButton).not.toBe(false);
      expect(wb.getTableAutoFilter(0)?.range).toMatchObject({
        firstRow: 0,
        firstCol: 0,
        lastRow: 3,
        lastCol: 1,
      });
      expect(store.getState().ui.filterRanges).toEqual([tableRange]);
    } finally {
      wb.detachStore(store);
      wb.dispose();
    }
  });

  it('shows the Mac table tab button enabled and pressed while the buttons show', async () => {
    const { wb, store } = await tableBook();
    try {
      expect(MAC_OBJECT_TAB_DISABLED_COMMAND_IDS).not.toContain('tableFilterButton');
      const host = document.createElement('div');
      const button = document.createElement('button');
      button.dataset.ribbonCommand = 'tableFilterButton';
      host.appendChild(button);
      const instance = {
        store,
        workbook: wb,
        i18n: createI18nController({ locale: 'ja' }),
      } as unknown as SpreadsheetInstance;
      projectTableDesign(host, instance);
      expect(button.getAttribute('aria-pressed')).toBe('true');
      expect(button.disabled).toBe(false);
      toggleTableOption(store, null, wb, 'filterButton');
      projectTableDesign(host, instance);
      expect(button.getAttribute('aria-pressed')).toBe('false');
    } finally {
      wb.detachStore(store);
      wb.dispose();
    }
  });
});
