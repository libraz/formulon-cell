import { describe, expect, it, vi } from 'vitest';
import { clearSelectedContents } from '../../../src/commands/clear-contents.js';
import { History } from '../../../src/commands/history.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 3 };
const deleteDataTables = vi.hoisted(() => vi.fn());

vi.mock('../../../src/commands/pivot-guard.js', () => ({
  planPivotSelection: () => ({ partial: false, whole: [{ pivotIndex: 0, range }] }),
  alertPivotEdit: () => {},
  pivotGuardAlert: () => ({ message: 'pivot' }),
}));
vi.mock('../../../src/commands/data-table-guard.js', () => ({
  planDataTableSelection: () => ({ partial: false, whole: [{ range }] }),
  alertDataTableEdit: () => {},
  dataTableGuardAlert: () => ({ message: 'table' }),
}));
vi.mock('../../../src/commands/pivot-table.js', () => ({ deletePivotTables: () => false }));
vi.mock('../../../src/commands/data-table.js', () => ({ deleteDataTables }));

describe('clearSelectedContents object removal', () => {
  it('leaves data tables and history untouched when a report cannot be removed', () => {
    deleteDataTables.mockImplementation((ctx: { history: History | null }) => {
      ctx.history?.push({ undo: () => {}, redo: () => {} });
      return true;
    });
    const store = createSpreadsheetStore();
    const history = new History();
    const result = clearSelectedContents({
      store,
      workbook: {} as unknown as WorkbookHandle,
      history,
    });
    expect(result.status).toBe('rejected');
    expect(deleteDataTables).not.toHaveBeenCalled();
    expect(history.canUndo()).toBe(false);
  });
});
