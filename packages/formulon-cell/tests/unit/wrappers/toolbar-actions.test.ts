import { afterEach, describe, expect, it, vi } from 'vitest';
import { getRecentFunctions } from '../../../src/commands/function-history.js';
import { History } from '../../../src/commands/history.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import type { SpreadsheetInstance } from '../../../src/mount/types.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import { handleAutoSum, handleAutoSumAction } from '../../../src/wrappers/toolbar-actions.js';

describe('toolbar AutoSum actions', () => {
  let workbook: WorkbookHandle | undefined;

  afterEach(() => {
    workbook?.dispose();
    workbook = undefined;
  });

  const createInstance = async (): Promise<SpreadsheetInstance> => {
    const store = createSpreadsheetStore();
    workbook = await WorkbookHandle.createDefault({ preferStub: true });
    return {
      host: document.createElement('div'),
      store,
      workbook,
      history: new History(),
      openFunctionArguments: vi.fn(),
      cellEditor: { begin: vi.fn(), isActive: () => false },
    } as unknown as SpreadsheetInstance;
  };

  it('opens the suggested formula in the cell editor for an empty single cell', async () => {
    const instance = await createInstance();
    const addr = { sheet: 0, row: 0, col: 0 };
    instance.workbook.setNumber(addr, 12);
    mutators.setCell(instance.store, addr, { kind: 'number', value: 12 });
    mutators.setActive(instance.store, { sheet: 0, row: 1, col: 0 });

    expect(handleAutoSum(instance, 'SUM')).toBe(true);
    expect(getRecentFunctions(instance.store)).toEqual(['SUM']);
    expect(instance.cellEditor.begin).toHaveBeenCalledWith('=SUM(A1:A1)', { mode: 'edit' });
    expect(instance.workbook.cellFormula({ sheet: 0, row: 1, col: 0 })).toBeNull();
  });

  it('still writes the total below when the active cell holds a number', async () => {
    const instance = await createInstance();
    const addr = { sheet: 0, row: 0, col: 0 };
    instance.workbook.setNumber(addr, 12);
    mutators.setCell(instance.store, addr, { kind: 'number', value: 12 });
    mutators.setActive(instance.store, addr);

    expect(handleAutoSum(instance, 'SUM')).toBe(true);
    expect(instance.cellEditor.begin).not.toHaveBeenCalled();
    expect(instance.workbook.cellFormula({ sheet: 0, row: 1, col: 0 })).toBe('=SUM(A1:A1)');
    expect(instance.store.getState().selection.active).toEqual({ sheet: 0, row: 1, col: 0 });
  });

  it('keeps a range selection in place after writing its totals', async () => {
    const instance = await createInstance();
    for (const row of [0, 1]) {
      const addr = { sheet: 0, row, col: 0 };
      instance.workbook.setNumber(addr, row + 1);
      mutators.setCell(instance.store, addr, { kind: 'number', value: row + 1 });
    }
    mutators.setActive(instance.store, { sheet: 0, row: 0, col: 0 });
    mutators.extendRangeTo(instance.store, { sheet: 0, row: 2, col: 0 });
    const before = instance.store.getState().selection;

    expect(handleAutoSum(instance, 'SUM')).toBe(true);
    expect(instance.workbook.cellFormula({ sheet: 0, row: 2, col: 0 })).toBe('=SUM(A1:A2)');
    expect(instance.store.getState().selection.active).toEqual(before.active);
    expect(instance.store.getState().selection.range).toEqual(before.range);
  });

  it('does not record an unsuccessful AutoSum or the function-picker action', async () => {
    const instance = await createInstance();
    mutators.setActive(instance.store, { sheet: 0, row: 5, col: 5 });

    expect(handleAutoSum(instance, 'AVERAGE')).toBe(false);
    expect(handleAutoSumAction(instance, 'MORE')).toBe(true);
    expect(getRecentFunctions(instance.store)).toEqual([]);
    expect(instance.openFunctionArguments).toHaveBeenCalledOnce();
  });
});
