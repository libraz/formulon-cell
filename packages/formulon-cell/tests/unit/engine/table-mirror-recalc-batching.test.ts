import { afterEach, describe, expect, it, vi } from 'vitest';
import { formatAsTable, updateTableOverlay } from '../../../src/commands/format-as-table.js';
import { hydrateAutoFilterFromEngine } from '../../../src/engine/auto-filter-sync.js';
import { hydrateTableOverlaysFromEngine } from '../../../src/engine/table-sync.js';
import type { Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const range: Range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 3 };
const open: WorkbookHandle[] = [];

afterEach(() => {
  for (const wb of open.splice(0)) wb.dispose();
});

const spyEngineRecalc = (wb: WorkbookHandle): ReturnType<typeof vi.fn> => {
  const internals = wb as unknown as { wb: { recalc: () => { ok: boolean } } };
  const spy = vi.fn(internals.wb.recalc.bind(internals.wb));
  internals.wb.recalc = spy;
  return spy;
};

describe('table mirror recalc batching', () => {
  it('toggling the header row recalcs once, not once per header cell', async () => {
    const wb = await WorkbookHandle.createDefault();
    open.push(wb);
    ['A', 'B', 'C', 'D'].forEach((h, col) => {
      wb.setText({ sheet: 0, row: 0, col }, h);
      for (let row = 1; row <= 3; row += 1) wb.setNumber({ sheet: 0, row, col }, row + col);
    });
    const store = createSpreadsheetStore();
    mutators.replaceCells(store, wb.cells(0));
    hydrateTableOverlaysFromEngine(wb, store);
    hydrateAutoFilterFromEngine(wb, store, 0);
    wb.attachStore(store);
    const overlay = formatAsTable(store, range, { workbook: wb });
    const id = overlay?.id ?? '';
    updateTableOverlay(store, id, { showHeader: false });

    const recalc = spyEngineRecalc(wb);
    expect(updateTableOverlay(store, id, { showHeader: true })).not.toBeNull();

    expect(recalc.mock.calls.length).toBeLessThanOrEqual(1);
  });
});
