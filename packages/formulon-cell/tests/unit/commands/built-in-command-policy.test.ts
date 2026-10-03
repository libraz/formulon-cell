import { describe, expect, it } from 'vitest';
import {
  BUILT_IN_COMMAND_OPERATION,
  canExecuteBuiltIn,
} from '../../../src/commands/built-in-command-policy.js';
import { recordRecentFunction } from '../../../src/commands/function-history.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { viewerPolicy } from '../../../src/commands/interaction-policy.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createI18nController } from '../../../src/i18n/controller.js';
import type { SpreadsheetInstance } from '../../../src/mount/types.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import {
  type ApplyRibbonCommandDeps,
  applyRibbonCommand,
} from '../../../src/toolbar/ribbon/apply-ribbon-command.js';

describe('built-in Mac command policy catalog', () => {
  it('classifies function leaves from the engine signatures and rejects unknown ids', () => {
    expect(BUILT_IN_COMMAND_OPERATION['mac.function.SUM']).toBe('formulaEdit');
    expect(BUILT_IN_COMMAND_OPERATION['mac.function.NOT_A_FUNCTION']).toBeUndefined();
  });

  it('classifies direct Mac controls and command menu leaves by operation', () => {
    expect(BUILT_IN_COMMAND_OPERATION['mac.data.sortAsc']).toBe('sort');
    expect(BUILT_IN_COMMAND_OPERATION['mac.data.filter']).toBe('filter');
    expect(BUILT_IN_COMMAND_OPERATION['mac.data.showDetail']).toBe('format');
    expect(BUILT_IN_COMMAND_OPERATION['mac.formulas.removeArrows.all']).toBe('formulaEdit');
    expect(BUILT_IN_COMMAND_OPERATION['mac.autosum.SUM']).toBe('formulaEdit');
    expect(BUILT_IN_COMMAND_OPERATION['mac.insert.shapeOval']).toBe('object');
    expect(BUILT_IN_COMMAND_OPERATION['mac.data.validation.clearRules']).toBe('validation');
    expect(BUILT_IN_COMMAND_OPERATION['mac.page.size.a4']).toBe('pageSetup');
    expect(BUILT_IN_COMMAND_OPERATION.scaleWidth).toBe('pageSetup');
    expect(BUILT_IN_COMMAND_OPERATION.sheetViewSelect).toBe('format');
  });

  it('leaves selection-only and removed ids out of the operation map', () => {
    expect(BUILT_IN_COMMAND_OPERATION['mac.formulas.errorCheck.run']).toBeUndefined();
    expect(BUILT_IN_COMMAND_OPERATION['mac.formulas.category.statistical']).toBeUndefined();
    expect(BUILT_IN_COMMAND_OPERATION['mac.no.such.command']).toBeUndefined();
  });

  it('keeps direct menu-root API calls from executing denied primary actions', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const history = new History();
    const controller = new InteractionController({ store, getWb: () => workbook, history });
    const unregister = registerInteractionController(store, controller);
    const host = document.createElement('div');
    let openedValidation = false;
    const instance = {
      store,
      workbook,
      history,
      host,
      i18n: createI18nController({ locale: 'en' }),
      openDataValidationDialog: () => {
        openedValidation = true;
      },
    } as unknown as SpreadsheetInstance;
    const deps = { inst: instance } as ApplyRibbonCommandDeps;
    try {
      controller.setPolicy({ ...viewerPolicy(), selection: false });
      const before = store.getState();
      for (const root of [
        'mac.page.pageBreaks',
        'mac.data.validation',
        'mac.formulas.errorCheck',
        'mac.formulas.removeArrows',
        'mac.formulas.autoSum',
      ]) {
        expect(canExecuteBuiltIn(store, root, 'ribbon').allowed, root).toBe(true);
        expect(applyRibbonCommand(root, deps), root).toBe(true);
        expect(store.getState(), root).toBe(before);
      }
      expect(openedValidation).toBe(false);
      expect(history.canUndo()).toBe(false);
      expect(document.querySelector('[role="dialog"]')).toBeNull();
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });

  it('allows submenu roots and zoom navigation in viewer mode while denying their operation leaves', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const controller = new InteractionController({
      store,
      getWb: () => workbook,
      history: new History(),
    });
    const unregister = registerInteractionController(store, controller);
    try {
      controller.setPolicy(viewerPolicy());
      for (const id of [
        'mac.formulas.removeArrows',
        'mac.formulas.errorCheck',
        'mac.data.validation',
        'mac.view.zoom100',
        'mac.formulas.errorCheck.run',
        'mac.data.workbookLinks',
      ]) {
        expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(true);
      }
      for (const id of [
        'mac.formulas.removeArrows.all',
        'mac.formulas.errorCheck.trace',
        'mac.data.validation.clearRules',
        'mac.data.sortAsc',
        'mac.formulas.precedents',
        'mac.formulas.dependents',
        'mac.formulas.recalc',
        'mac.formulas.sheetRecalc',
        'mac.no.such.command',
      ]) {
        expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(false);
      }
      controller.setPolicy({ ...viewerPolicy(), selection: false });
      expect(canExecuteBuiltIn(store, 'mac.formulas.errorCheck.run', 'ribbon').allowed).toBe(false);
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });

  it('authorizes a dynamic function leaf only after a successful live-catalog MRU record', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const history = new History();
    const controller = new InteractionController({ store, getWb: () => workbook, history });
    const unregister = registerInteractionController(store, controller);
    const range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };
    try {
      controller.setPolicy({
        editable: [range],
        operations: { formulaEdit: true },
        defaultOperation: 'deny',
        selection: true,
      });
      expect(canExecuteBuiltIn(store, 'mac.function.LIVE_ONLY_FN', 'ribbon')).toMatchObject({
        allowed: false,
        code: 'unsupported',
      });
      expect(recordRecentFunction(store, 'LIVE_ONLY_FN', new Set(['LIVE_ONLY_FN']))).toBe(true);
      expect(canExecuteBuiltIn(store, 'mac.function.LIVE_ONLY_FN', 'ribbon')).toEqual({
        allowed: true,
      });
      expect(canExecuteBuiltIn(store, 'mac.function.NOT_A_FUNCTION', 'ribbon')).toMatchObject({
        allowed: false,
        code: 'unsupported',
      });
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });
});
