import { describe, expect, it } from 'vitest';
import { BUILT_IN_COMMAND_IDS } from '../../../src/commands/built-in-command-ids.js';
import {
  BUILT_IN_COMMAND_OPERATION,
  canExecuteBuiltIn,
  canExecuteBuiltInRequest,
} from '../../../src/commands/built-in-command-policy.js';
import { recordRecentFunction } from '../../../src/commands/function-history.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { viewerPolicy } from '../../../src/commands/interaction-policy.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createI18nController } from '../../../src/i18n/controller.js';
import type { SpreadsheetInstance } from '../../../src/mount/types.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import {
  type ApplyRibbonCommandDeps,
  applyRibbonCommand,
} from '../../../src/toolbar/ribbon/apply-ribbon-command.js';

describe('built-in Mac command policy catalog', () => {
  it('classifies every built-in command id', () => {
    expect(Object.keys(BUILT_IN_COMMAND_OPERATION).sort()).toEqual([...BUILT_IN_COMMAND_IDS]);
  });

  it('classifies direct Mac controls and command menu leaves by operation', () => {
    expect(BUILT_IN_COMMAND_OPERATION['mac.data.sortAsc']).toBe('sort');
    expect(BUILT_IN_COMMAND_OPERATION['mac.data.filter']).toBe('filter');
    expect(BUILT_IN_COMMAND_OPERATION['mac.data.showDetail']).toBe('resizeRows');
    expect(BUILT_IN_COMMAND_OPERATION['mac.formulas.removeArrows.all']).toBeNull();
    expect(BUILT_IN_COMMAND_OPERATION.pivotRefresh).toBe('pivot');
    expect(BUILT_IN_COMMAND_OPERATION['mac.autosum.SUM']).toBe('formulaEdit');
    expect(BUILT_IN_COMMAND_OPERATION['mac.insert.shapeOval']).toBe('object');
    expect(BUILT_IN_COMMAND_OPERATION['mac.insert.recommendedChart']).toBe('object');
    expect(BUILT_IN_COMMAND_OPERATION['mac.data.validation.clearRules']).toBe('validation');
    expect(BUILT_IN_COMMAND_OPERATION['mac.page.size.a4']).toBe('pageSetup');
    expect(BUILT_IN_COMMAND_OPERATION.scaleWidth).toBe('pageSetup');
    expect(BUILT_IN_COMMAND_OPERATION.sheetViewSelect).toBe('filter');
  });

  it('allows Recommended Chart for normal instances but denies it in viewer mode', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const controller = new InteractionController({
      store,
      getWb: () => workbook,
      history: new History(),
    });
    const unregister = registerInteractionController(store, controller);
    try {
      controller.setPolicy(undefined);
      expect(canExecuteBuiltIn(store, 'mac.insert.recommendedChart', 'ribbon')).toEqual({
        allowed: true,
      });

      controller.setPolicy(viewerPolicy());
      expect(canExecuteBuiltIn(store, 'mac.insert.recommendedChart', 'ribbon').allowed).toBe(false);
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });

  it('classifies selection-only commands as non-mutating and keeps removed ids out', () => {
    expect(BUILT_IN_COMMAND_OPERATION['mac.formulas.errorCheck.run']).toBeNull();
    expect(Object.hasOwn(BUILT_IN_COMMAND_OPERATION, 'mac.formulas.category.statistical')).toBe(
      false,
    );
    expect(Object.hasOwn(BUILT_IN_COMMAND_OPERATION, 'mac.no.such.command')).toBe(false);
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
        'mac.formulas.autoSum',
      ] as const) {
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
        'mac.review.previousNote',
        'mac.review.nextNote',
        'mac.review.previousComment',
        'mac.review.nextComment',
        'mac.formulas.recalc',
        'mac.formulas.sheetRecalc',
        // Trace arrows are an overlay, not workbook state.
        'mac.formulas.removeArrows.all',
        'mac.formulas.errorCheck.trace',
        'mac.formulas.precedents',
        'mac.formulas.dependents',
      ] as const) {
        expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(true);
      }
      for (const id of [
        'mac.data.validation.clearRules',
        'mac.data.sortAsc',
        'pivotRefresh',
      ] as const) {
        expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(false);
      }
      expect(canExecuteBuiltInRequest(store, 'mac.no.such.command', 'ribbon')).toMatchObject({
        allowed: false,
        code: 'invalid',
      });
      controller.setPolicy({ ...viewerPolicy(), selection: false });
      for (const id of [
        'mac.formulas.errorCheck.run',
        'mac.review.previousNote',
        'mac.review.nextNote',
        'mac.review.previousComment',
        'mac.review.nextComment',
      ] as const) {
        expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(false);
      }
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
        code: 'invalid',
      });
      expect(recordRecentFunction(store, 'LIVE_ONLY_FN', new Set(['LIVE_ONLY_FN']))).toBe(true);
      expect(canExecuteBuiltIn(store, 'mac.function.LIVE_ONLY_FN', 'ribbon')).toEqual({
        allowed: true,
      });
      expect(canExecuteBuiltIn(store, 'mac.function.SUM', 'ribbon')).toEqual({ allowed: true });
      expect(canExecuteBuiltIn(store, 'mac.function.NOT_A_FUNCTION', 'ribbon')).toMatchObject({
        allowed: false,
        code: 'invalid',
      });
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });

  describe('on a protected sheet', () => {
    const auditing = [
      'precedents',
      'dependents',
      'clearArrows',
      'mac.formulas.precedents',
      'mac.formulas.dependents',
      'mac.formulas.removeArrows.all',
      'mac.formulas.removeArrows.precedents',
      'mac.formulas.removeArrows.dependents',
    ] as const;

    it('refuses the auditing arrows and pivot refresh whatever the options, policy or not', () => {
      const store = createSpreadsheetStore();
      for (const id of [...auditing, 'pivotRefresh'] as const) {
        expect(canExecuteBuiltIn(store, id, 'ribbon'), id).toEqual({ allowed: true });
      }
      setProtectedSheet(store, 0, true, {
        permissions: { pivotTables: true, objects: true, formatCells: true },
      });
      for (const id of [...auditing, 'pivotRefresh'] as const) {
        expect(canExecuteBuiltIn(store, id, 'ribbon'), id).toMatchObject({
          allowed: false,
          code: 'protected',
        });
      }
    });

    it('never gates the calculation mode', () => {
      const store = createSpreadsheetStore();
      setProtectedSheet(store, 0, true);
      for (const id of [
        'mac.formulas.calc.auto',
        'mac.formulas.calc.autoNoTable',
        'mac.formulas.calc.manual',
        'mac.formulas.calc.iterative',
      ] as const) {
        expect(BUILT_IN_COMMAND_OPERATION[id], id).toBeNull();
        expect(canExecuteBuiltIn(store, id, 'ribbon'), id).toEqual({ allowed: true });
      }
    });

    it('lets pivot layout changes through only with the PivotTable option', async () => {
      const store = createSpreadsheetStore();
      const workbook = await WorkbookHandle.createDefault({ preferStub: true });
      const controller = new InteractionController({
        store,
        getWb: () => workbook,
        history: new History(),
      });
      const unregister = registerInteractionController(store, controller);
      try {
        controller.setPolicy({ defaultOperation: 'allow' });
        setProtectedSheet(store, 0, true);
        expect(canExecuteBuiltIn(store, 'pivotSortAsc', 'ribbon')).toMatchObject({
          allowed: false,
          code: 'protected',
        });
        setProtectedSheet(store, 0, true, { permissions: { pivotTables: true } });
        expect(canExecuteBuiltIn(store, 'pivotSortAsc', 'ribbon')).toEqual({ allowed: true });
        expect(canExecuteBuiltIn(store, 'pivotRefresh', 'ribbon')).toMatchObject({
          allowed: false,
          code: 'protected',
        });
      } finally {
        unregister();
        controller.dispose();
        workbook.dispose();
      }
    });
  });
});
