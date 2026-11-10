import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import type { InteractionPolicy } from '../../../src/commands/interaction-policy.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { planSelectionFormat } from '../../../src/commands/selection-format-plan.js';
import {
  circleInvalidData,
  clearValidationRules,
  type ValidationActionContext,
} from '../../../src/commands/validation-actions.js';
import { addrKey } from '../../../src/engine/address.js';
import type { Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { runFormatDialogTransaction } from '../../../src/interact/format-dialog-transaction.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import type { CellValidation } from '../../../src/store/types.js';

const RULE: CellValidation = { kind: 'whole', op: 'between', a: 1, b: 10 };
const RANGE: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 };
const A1 = { sheet: 0, row: 0, col: 0 };
const A2 = { sheet: 0, row: 1, col: 0 };

const allowAll: InteractionPolicy = { defaultOperation: 'allow' };
const denyValidation: InteractionPolicy = {
  defaultOperation: 'allow',
  operations: { validation: false },
};
const denyOrigin = (denied: string): InteractionPolicy => ({
  defaultOperation: 'allow',
  restrict: ({ origin }) => origin !== denied,
});

describe('validation writes under the interaction policy', () => {
  const handles: WorkbookHandle[] = [];
  const controllers: InteractionController[] = [];

  afterEach(() => {
    for (const controller of controllers.splice(0)) controller.dispose();
    for (const handle of handles.splice(0)) handle.dispose();
  });

  const setup = async () => {
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    handles.push(workbook);
    const store = createSpreadsheetStore();
    const history = new History();
    const controller = new InteractionController({ store, getWb: () => workbook, history });
    controllers.push(controller);
    registerInteractionController(store, controller);
    mutators.setRange(store, RANGE);
    const ctx: ValidationActionContext = {
      store,
      history,
      workbook,
      origin: 'ribbon',
      commandId: 'dataValidation',
    };
    const ruleAt = (addr: typeof A1) => store.getState().format.formats.get(addrKey(addr));
    const seedRules = (): void => {
      mutators.setCellFormat(store, A1, { validation: RULE });
      mutators.setCellFormat(store, A2, { validation: RULE });
    };
    const commitDialog = (patch: { validation?: CellValidation }) => {
      const state = store.getState();
      const plan = planSelectionFormat(state);
      if (!plan) throw new Error('no plan');
      return runFormatDialogTransaction({
        store,
        history,
        getWb: () => workbook,
        state,
        liveWb: workbook,
        plan,
        action: { patch },
        merge: { action: null, range: state.selection.range },
      });
    };
    return { store, history, controller, ctx, ruleAt, seedRules, commitDialog };
  };

  describe('Format Cells dialog commit', () => {
    it('writes a rule under a policy and replays it through undo and redo', async () => {
      const { history, controller, ruleAt, commitDialog } = await setup();
      controller.setPolicy(allowAll);
      expect(commitDialog({ validation: RULE })).toBe(true);
      expect(ruleAt(A1)?.validation).toEqual(RULE);
      expect(history.undo()).toBe(true);
      expect(ruleAt(A1)?.validation).toBeUndefined();
      expect(history.redo()).toBe(true);
      expect(ruleAt(A2)?.validation).toEqual(RULE);
    });

    it('refuses the commit when the policy denies validation', async () => {
      const { history, controller, ruleAt, commitDialog } = await setup();
      controller.setPolicy(denyValidation);
      expect(commitDialog({ validation: RULE })).toBe(false);
      expect(ruleAt(A1)).toBeUndefined();
      expect(history.canUndo()).toBe(false);
    });

    it('refuses the commit on locked cells of a protected sheet', async () => {
      const { store, history, controller, ruleAt, commitDialog } = await setup();
      controller.setPolicy(allowAll);
      setProtectedSheet(store, 0, true);
      expect(commitDialog({ validation: RULE })).toBe(false);
      expect(ruleAt(A1)).toBeUndefined();
      expect(history.canUndo()).toBe(false);
    });

    it('refuses replay once the policy denies the undo or redo', async () => {
      const { history, controller, ruleAt, commitDialog } = await setup();
      controller.setPolicy(allowAll);
      expect(commitDialog({ validation: RULE })).toBe(true);
      controller.setPolicy(denyOrigin('undo'));
      expect(history.undo()).toBe(false);
      expect(ruleAt(A1)?.validation).toEqual(RULE);
      controller.setPolicy(allowAll);
      expect(history.undo()).toBe(true);
      controller.setPolicy(denyValidation);
      expect(history.redo()).toBe(false);
      expect(ruleAt(A1)?.validation).toBeUndefined();
    });
  });

  describe('Clear Validation', () => {
    it('clears the rules under a policy and replays through undo and redo', async () => {
      const { history, controller, ctx, ruleAt, seedRules } = await setup();
      seedRules();
      controller.setPolicy(allowAll);
      expect(clearValidationRules(ctx, RANGE)).toMatchObject({ ok: true });
      expect(ruleAt(A1)?.validation).toBeUndefined();
      expect(history.undo()).toBe(true);
      expect(ruleAt(A1)?.validation).toEqual(RULE);
      expect(history.redo()).toBe(true);
      expect(ruleAt(A2)?.validation).toBeUndefined();
    });

    it('refuses before clearing when the policy denies validation', async () => {
      const { history, controller, ctx, ruleAt, seedRules } = await setup();
      seedRules();
      controller.setPolicy(denyValidation);
      expect(clearValidationRules(ctx, RANGE)).toMatchObject({
        ok: false,
        decision: { code: 'operationDenied' },
      });
      expect(ruleAt(A1)?.validation).toEqual(RULE);
      expect(history.canUndo()).toBe(false);
    });

    it('refuses on locked cells of a protected sheet, with or without a policy', async () => {
      const { store, controller, ctx, ruleAt, seedRules } = await setup();
      seedRules();
      setProtectedSheet(store, 0, true);
      expect(clearValidationRules(ctx, RANGE)).toMatchObject({
        ok: false,
        decision: { code: 'protected' },
      });
      controller.setPolicy(allowAll);
      expect(clearValidationRules(ctx, RANGE)).toMatchObject({
        ok: false,
        decision: { code: 'protected' },
      });
      expect(ruleAt(A1)?.validation).toEqual(RULE);
    });

    it('refuses replay once the policy denies validation', async () => {
      const { history, controller, ctx, ruleAt, seedRules } = await setup();
      seedRules();
      controller.setPolicy(allowAll);
      clearValidationRules(ctx, RANGE);
      controller.setPolicy(denyValidation);
      expect(history.undo()).toBe(false);
      expect(ruleAt(A1)?.validation).toBeUndefined();
    });
  });

  describe('Circle Invalid Data', () => {
    const seedInvalid = (store: ReturnType<typeof createSpreadsheetStore>): void => {
      mutators.setCellFormat(store, A1, { validation: RULE });
      mutators.setCell(store, A1, { kind: 'number', value: 99 });
    };

    it('circles invalid cells under a policy and replays through undo and redo', async () => {
      const { store, history, controller, ctx } = await setup();
      seedInvalid(store);
      controller.setPolicy(allowAll);
      expect(circleInvalidData(ctx, RANGE)).toMatchObject({ ok: true });
      expect([...store.getState().errorIndicators.validationCircles]).toEqual([addrKey(A1)]);
      expect(history.undo()).toBe(true);
      expect(store.getState().errorIndicators.validationCircles.size).toBe(0);
      expect(history.redo()).toBe(true);
      expect(store.getState().errorIndicators.validationCircles.size).toBe(1);
    });

    it('refuses when the policy denies validation or the sheet is protected', async () => {
      const { store, history, controller, ctx } = await setup();
      seedInvalid(store);
      controller.setPolicy(denyValidation);
      expect(circleInvalidData(ctx, RANGE).ok).toBe(false);
      controller.setPolicy(allowAll);
      setProtectedSheet(store, 0, true);
      expect(circleInvalidData(ctx, RANGE)).toMatchObject({
        ok: false,
        decision: { code: 'protected' },
      });
      expect(store.getState().errorIndicators.validationCircles.size).toBe(0);
      expect(history.canUndo()).toBe(false);
    });

    it('refuses replay once the policy denies the redo', async () => {
      const { store, history, controller, ctx } = await setup();
      seedInvalid(store);
      controller.setPolicy(allowAll);
      circleInvalidData(ctx, RANGE);
      expect(history.undo()).toBe(true);
      controller.setPolicy(denyOrigin('redo'));
      expect(history.redo()).toBe(false);
      expect(store.getState().errorIndicators.validationCircles.size).toBe(0);
    });
  });
});
