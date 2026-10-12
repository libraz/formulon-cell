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

  it('circles native invalid cells and validation-formatted blanks through the UI action', async () => {
    const store = createSpreadsheetStore();
    workbook = await WorkbookHandle.createDefault();
    const host = document.createElement('div');
    const history = new History();
    const instance = {
      host,
      store,
      workbook,
      history,
    } as unknown as SpreadsheetInstance;
    expect(workbook.capabilities.validationEvaluation).toBe(true);
    const listInvalidSpy = vi.spyOn(workbook, 'listInvalidCells');
    expect(
      workbook.addValidationEntry(0, {
        ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 }],
        type: 1,
        op: 0,
        allowBlank: false,
        formula1: '1',
        formula2: '10',
      }),
    ).toBe(true);
    workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 99);
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 });
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        validation: { kind: 'whole', op: 'between', a: 1, b: 10, allowBlank: false },
      },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 1, col: 0 },
      {
        validation: { kind: 'whole', op: 'between', a: 1, b: 10, allowBlank: false },
      },
    );
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'number', value: 99 });

    createDefaultDynamicDropdownsCtx(instance).applyDataValidationAction('circle-invalid');

    expect(listInvalidSpy).toHaveBeenCalled();
    expect([...store.getState().errorIndicators.validationCircles]).toEqual(['0:0:0', '0:1:0']);
    expect(history.undo()).toBe(true);
    expect(store.getState().errorIndicators.validationCircles.size).toBe(0);
  });

  it('keeps illustration handler keys and override semantics stable', async () => {
    const instance = await createInstance();
    const base = createDefaultDynamicDropdownsCtx(instance);
    const illustrationKeys = [
      'updateArrangeMenu',
      'applyArrangeAction',
      'insertPictureFromRibbon',
      'insertShapeFromRibbon',
      'insertScreenshotFromRibbon',
    ] as const;
    const illustrationKeySet = new Set<string>(illustrationKeys);
    const baseKeys = Object.keys(base);

    expect(baseKeys.filter((key) => illustrationKeySet.has(key))).toEqual(illustrationKeys);

    const objectOverride = vi.fn();
    const objectCtx = createDefaultDynamicDropdownsCtx(instance, {
      overrides: { insertShapeFromRibbon: objectOverride },
    });
    expect(Object.keys(objectCtx)).toEqual(baseKeys);
    expect(objectCtx.insertShapeFromRibbon).toBe(objectOverride);

    const getterOverride = vi.fn(() => ({ insertShapeFromRibbon: objectOverride }));
    const getterCtx = createDefaultDynamicDropdownsCtx(instance, { overrides: getterOverride });
    const descriptor = Object.getOwnPropertyDescriptor(getterCtx, 'insertShapeFromRibbon');
    expect(Object.keys(getterCtx)).toEqual(baseKeys);
    expect(descriptor?.enumerable).toBe(true);
    expect(descriptor?.get).toBeTypeOf('function');
    expect(getterCtx.insertShapeFromRibbon).toBe(objectOverride);
    expect(getterOverride).toHaveBeenCalled();
  });
});
