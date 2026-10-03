import { afterEach, describe, expect, it, vi } from 'vitest';
import { getRecentFunctions } from '../../../src/commands/function-history.js';
import { History } from '../../../src/commands/history.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createDefaultDynamicDropdownsCtx } from '../../../src/mount/dynamic-dropdowns-defaults.js';
import type { SpreadsheetInstance } from '../../../src/mount/types.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('default dynamic dropdown AutoSum', () => {
  let workbook: WorkbookHandle | undefined;

  afterEach(() => {
    workbook?.dispose();
    workbook = undefined;
  });

  const createInstance = async (): Promise<SpreadsheetInstance> => {
    const store = createSpreadsheetStore();
    workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const host = document.createElement('div');
    const instance = {
      host,
      store,
      workbook,
      history: new History(),
      openFunctionArguments: vi.fn(),
    } as unknown as SpreadsheetInstance;
    return instance;
  };

  it('records the chosen function only when AutoSum places a formula', async () => {
    const instance = await createInstance();
    const addr = { sheet: 0, row: 0, col: 0 };
    instance.workbook.setNumber(addr, 12);
    mutators.setCell(instance.store, addr, { kind: 'number', value: 12 });
    mutators.setActive(instance.store, { sheet: 0, row: 1, col: 0 });
    const ctx = createDefaultDynamicDropdownsCtx(instance);

    ctx.applyAutoSumFormula('AVERAGE');

    expect(getRecentFunctions(instance.store)).toEqual(['AVERAGE']);
    expect(instance.workbook.cellFormula({ sheet: 0, row: 1, col: 0 })).toBe('=AVERAGE(A1:A1)');
  });

  it('does not record a no-op or the MORE dialog action', async () => {
    const instance = await createInstance();
    const ctx = createDefaultDynamicDropdownsCtx(instance);

    ctx.applyAutoSumFormula('SUM');
    ctx.applyAutoSumFormula('MORE');

    expect(getRecentFunctions(instance.store)).toEqual([]);
    expect(instance.openFunctionArguments).toHaveBeenCalledOnce();
  });
});
