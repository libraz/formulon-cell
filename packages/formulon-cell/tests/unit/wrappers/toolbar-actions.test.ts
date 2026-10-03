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
    } as unknown as SpreadsheetInstance;
  };

  it('records the selected aggregate when AutoSum inserts a formula', async () => {
    const instance = await createInstance();
    const addr = { sheet: 0, row: 0, col: 0 };
    instance.workbook.setNumber(addr, 12);
    mutators.setCell(instance.store, addr, { kind: 'number', value: 12 });
    mutators.setActive(instance.store, { sheet: 0, row: 1, col: 0 });

    expect(handleAutoSum(instance, 'SUM')).toBe(true);
    expect(getRecentFunctions(instance.store)).toEqual(['SUM']);
    expect(instance.workbook.cellFormula({ sheet: 0, row: 1, col: 0 })).toBe('=SUM(A1:A1)');
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
