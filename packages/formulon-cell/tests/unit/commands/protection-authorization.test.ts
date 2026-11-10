import { afterEach, describe, expect, it } from 'vitest';
import { canExecuteBuiltIn } from '../../../src/commands/built-in-command-policy.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import type { InteractionPolicy } from '../../../src/commands/interaction-policy.js';
import {
  isSheetProtected,
  isWorkbookStructureProtected,
  protectionCheckpoint,
  setProtectedSheet,
} from '../../../src/commands/protection.js';
import {
  addAllowedEditRangeAction,
  clearAllowedEditRangesAction,
  type ProtectionActionContext,
  removeAllowedEditRangeAction,
  setCellsLockedAction,
  setSheetProtectionAction,
  setWorkbookProtectionAction,
  updateAllowedEditRangeAction,
} from '../../../src/commands/protection-actions.js';
import { planSelectionFormat } from '../../../src/commands/selection-format-plan.js';
import { addrKey } from '../../../src/engine/address.js';
import type { Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { runFormatDialogTransaction } from '../../../src/interact/format-dialog-transaction.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const RANGE: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 };
const A1 = { sheet: 0, row: 0, col: 0 };

const allowAll: InteractionPolicy = { defaultOperation: 'allow' };
const denyProtection: InteractionPolicy = {
  defaultOperation: 'allow',
  operations: { protection: false },
};
const denyOrigin = (denied: string): InteractionPolicy => ({
  defaultOperation: 'allow',
  restrict: ({ origin }) => origin !== denied,
});

describe('protection writes under the interaction policy', () => {
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
    const ctx: ProtectionActionContext = {
      store,
      history,
      origin: 'ribbon',
      commandId: 'protectReview',
    };
    const locked = () => store.getState().format.formats.get(addrKey(A1))?.locked;
    const ranges = () => store.getState().protection.allowedEditRanges;
    return { store, workbook, history, controller, ctx, locked, ranges };
  };

  it('protects and unprotects a sheet under a policy, unprotect on a protected sheet', async () => {
    const { store, controller, ctx } = await setup();
    controller.setPolicy(allowAll);
    expect(setSheetProtectionAction(ctx, 0, true)).toBe(true);
    expect(isSheetProtected(store.getState(), 0)).toBe(true);
    expect(setSheetProtectionAction(ctx, 0, false)).toBe(true);
    expect(isSheetProtected(store.getState(), 0)).toBe(false);
  });

  it('protects and unprotects the workbook structure under a policy', async () => {
    const { store, controller, ctx } = await setup();
    controller.setPolicy(allowAll);
    expect(setWorkbookProtectionAction(ctx, true)).toBe(true);
    expect(isWorkbookStructureProtected(store.getState())).toBe(true);
    expect(setWorkbookProtectionAction(ctx, false)).toBe(true);
    expect(isWorkbookStructureProtected(store.getState())).toBe(false);
  });

  it('refuses every protection write and command when the policy denies protection', async () => {
    const { store, controller, ctx, locked, ranges } = await setup();
    controller.setPolicy(denyProtection);
    expect(setSheetProtectionAction(ctx, 0, true)).toBe(false);
    expect(setWorkbookProtectionAction(ctx, true)).toBe(false);
    expect(addAllowedEditRangeAction(ctx, RANGE, { title: 'R' })).toBeNull();
    expect(setCellsLockedAction(ctx, RANGE, false)).toBe(false);
    expect(isSheetProtected(store.getState(), 0)).toBe(false);
    expect(isWorkbookStructureProtected(store.getState())).toBe(false);
    expect(ranges()).toEqual([]);
    expect(locked()).toBeUndefined();
    for (const id of [
      'protect',
      'protectReview',
      'protectWorkbookReview',
      'protectionReview',
      'mac.review.protectSheet',
      'mac.review.protectWorkbook',
    ] as const) {
      expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(false);
    }
  });

  it('keeps protection commands available on a protected sheet', async () => {
    const { store, controller, ctx } = await setup();
    controller.setPolicy(allowAll);
    setProtectedSheet(store, 0, true);
    expect(canExecuteBuiltIn(store, 'protectReview', 'ribbon').allowed).toBe(true);
    expect(setSheetProtectionAction(ctx, 0, false)).toBe(true);
  });

  it('authorizes allowed edit range changes and keeps the manager checkpoint restorable', async () => {
    const { store, controller, ctx, ranges } = await setup();
    controller.setPolicy(allowAll);
    const kept = addAllowedEditRangeAction(ctx, RANGE, { title: 'Kept' });
    const restore = protectionCheckpoint(store, undefined);
    const added = addAllowedEditRangeAction(ctx, { ...RANGE, c0: 2, c1: 2 }, { title: 'New' });
    if (!kept || !added) throw new Error('range not added');
    expect(updateAllowedEditRangeAction(ctx, kept, { title: 'Renamed' })).toBe(true);
    expect(removeAllowedEditRangeAction(ctx, added)).toBe(true);
    expect(ranges().map((entry) => entry.title)).toEqual(['Renamed']);
    controller.setPolicy(denyProtection);
    expect(clearAllowedEditRangesAction(ctx, 0)).toBe(false);
    restore();
    expect(ranges().map((entry) => entry.title)).toEqual(['Kept']);
    controller.setPolicy(allowAll);
    expect(clearAllowedEditRangesAction(ctx, 0)).toBe(true);
    expect(ranges()).toEqual([]);
  });

  it('locks cells under a policy, refuses locked cells of a protected sheet, and replays', async () => {
    const { store, history, controller, ctx, locked } = await setup();
    controller.setPolicy(allowAll);
    expect(setCellsLockedAction(ctx, RANGE, false)).toBe(true);
    expect(locked()).toBe(false);
    controller.setPolicy(denyOrigin('undo'));
    expect(history.undo()).toBe(false);
    controller.setPolicy(allowAll);
    expect(history.undo()).toBe(true);
    expect(locked()).toBeUndefined();
    controller.setPolicy(denyProtection);
    expect(history.redo()).toBe(false);
    controller.setPolicy(allowAll);
    setProtectedSheet(store, 0, true);
    expect(setCellsLockedAction(ctx, RANGE, false)).toBe(false);
    expect(locked()).toBeUndefined();
  });

  it('decides Format Cells locked and hidden as protection, not format', async () => {
    const { store, workbook, history, controller, locked } = await setup();
    controller.setPolicy({ defaultOperation: 'allow', operations: { protection: false } });
    const commit = () => {
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
        action: { patch: { locked: false, formulaHidden: true } },
        merge: { action: null, range: state.selection.range },
        entry: { origin: 'ribbon', commandId: 'formatCells' },
      });
    };
    expect(commit()).toBe(false);
    expect(locked()).toBeUndefined();
    controller.setPolicy({ defaultOperation: 'allow', operations: { format: false } });
    expect(commit()).toBe(true);
    expect(locked()).toBe(false);
  });
  it('lets formatting through on locked cells of a protected sheet only with Format cells on', async () => {
    const { store, workbook, history, controller } = await setup();
    const bold = () => store.getState().format.formats.get(addrKey(A1))?.bold;
    const commit = () => {
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
        action: { patch: { bold: true } },
        merge: { action: null, range: state.selection.range },
        entry: { origin: 'ribbon', commandId: 'formatCells' },
      });
    };
    for (const policy of [undefined, allowAll]) {
      controller.setPolicy(policy);
      setProtectedSheet(store, 0, true);
      expect(commit(), 'option off').toBe(false);
      expect(bold()).toBeUndefined();
      setProtectedSheet(store, 0, false);
      setProtectedSheet(store, 0, true, { permissions: { formatCells: true } });
      expect(commit(), 'option on').toBe(true);
      expect(bold()).toBe(true);
      setProtectedSheet(store, 0, false);
      for (const row of [0, 1]) mutators.setCellFormat(store, { sheet: 0, row, col: 0 }, null);
    }
  });
});
