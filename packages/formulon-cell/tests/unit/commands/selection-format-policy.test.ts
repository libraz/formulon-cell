import { describe, expect, it } from 'vitest';
import { canExecuteBuiltIn } from '../../../src/commands/built-in-command-policy.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { fixedFormPolicy } from '../../../src/commands/interaction-policy.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('selection-format built-in command policy', () => {
  it('authorizes selection-format commands against every selected area', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const controller = new InteractionController({
      store,
      getWb: () => workbook,
      history: new History(),
    });
    const unregister = registerInteractionController(store, controller);
    store.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        extraRanges: [{ sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 }],
      },
    }));
    try {
      const policy = fixedFormPolicy({
        ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }],
      });
      controller.setPolicy({
        ...policy,
        operations: { ...policy.operations, format: true },
      });
      expect(canExecuteBuiltIn(store, 'bold', 'ribbon')).toMatchObject({
        allowed: false,
        code: 'cellIneligible',
      });

      store.setState((state) => ({
        ...state,
        selection: { ...state.selection, extraRanges: [] },
      }));
      for (const id of [
        'bold',
        'strike',
        'currency',
        'percent',
        'comma',
        'alignLeft',
        'alignCenter',
        'alignRight',
        'alignL',
        'alignC',
        'alignR',
        'top',
        'middle',
        'bottomAlign',
        'decUp',
        'decDown',
        'wrap',
        'general',
        'textOrientation',
        'indentDecrease',
        'indentIncrease',
        'fontGrow',
        'fontShrink',
        'fontFamily',
        'fontSize',
        'fontColor',
        'fillColor',
        'numberFormat',
      ]) {
        expect(canExecuteBuiltIn(store, id, 'ribbon'), id).toEqual({ allowed: true });
      }
      for (const id of ['borders', 'formatCells', 'editPhonetic', 'pasteFormatsOnly']) {
        expect(canExecuteBuiltIn(store, id, 'ribbon'), id).toMatchObject({
          allowed: false,
          code: 'unsupported',
        });
      }
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });
});
