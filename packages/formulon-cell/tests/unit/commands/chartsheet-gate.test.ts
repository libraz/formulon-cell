import { describe, expect, it } from 'vitest';
import { canExecuteBuiltIn } from '../../../src/commands/built-in-command-policy.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { isCellWritable } from '../../../src/commands/protection.js';
import { SheetKind } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const chartsheetStore = () => {
  const store = createSpreadsheetStore();
  store.setState((s) => ({
    ...s,
    layout: { ...s.layout, sheetKinds: new Map([[1, SheetKind.Chartsheet]]) },
  }));
  return store;
};

describe('a sheet without cells', () => {
  it('refuses cell writes on a chartsheet and keeps them on a worksheet', () => {
    const store = chartsheetStore();
    expect(isCellWritable(store.getState(), { sheet: 1, row: 0, col: 0 })).toBe(false);
    expect(isCellWritable(store.getState(), { sheet: 0, row: 0, col: 0 })).toBe(true);
  });

  it('refuses cell commands from the built-in gate only while it is the active sheet', () => {
    const store = chartsheetStore();
    expect(canExecuteBuiltIn(store, 'bold', 'ribbon').allowed).toBe(true);
    mutators.setSheetIndex(store, 1);
    const refused = canExecuteBuiltIn(store, 'bold', 'ribbon');
    expect(refused).toMatchObject({ allowed: false, reasonKey: 'sheetHasNoCells' });
    expect(canExecuteBuiltIn(store, 'insertRows', 'ribbon').allowed).toBe(false);
    expect(canExecuteBuiltIn(store, 'paste', 'keyboard').allowed).toBe(false);
  });

  it('lets the controller refuse a cell edit without a policy, and still allow object edits', async () => {
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, sheetKinds: new Map([[0, SheetKind.Chartsheet]]) },
    }));
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const controller = new InteractionController({
      store,
      getWb: () => workbook,
      history: new History(),
    });
    const unregister = registerInteractionController(store, controller);
    try {
      const cell = { sheet: 0, row: 0, col: 0 };
      expect(
        controller.canExecute({
          operation: 'valueEdit',
          origin: 'editor',
          effects: [{ kind: 'cells', cells: [cell] }],
        }),
      ).toMatchObject({ allowed: false, reasonKey: 'sheetHasNoCells' });
      expect(
        controller.canExecute({
          operation: 'object',
          origin: 'ribbon',
          effects: [{ kind: 'sheet', sheet: 0 }],
        }).allowed,
      ).toBe(true);
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });
});
