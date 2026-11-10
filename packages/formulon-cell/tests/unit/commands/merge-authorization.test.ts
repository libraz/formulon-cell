import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import type { InteractionPolicy } from '../../../src/commands/interaction-policy.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { planSelectionFormat } from '../../../src/commands/selection-format-plan.js';
import type { Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { defaultStrings } from '../../../src/i18n/strings.js';
import type { DialogEntry } from '../../../src/interact/dialog-entry.js';
import { runFormatDialogTransaction } from '../../../src/interact/format-dialog-transaction.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import { applyMergeAction, type MergeAction } from '../../../src/toolbar/merge-action.js';

const RANGE: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };

const allowAll: InteractionPolicy = { defaultOperation: 'allow' };
const denyMerge: InteractionPolicy = { defaultOperation: 'allow', operations: { merge: false } };
const denyOrigin = (denied: string): InteractionPolicy => ({
  defaultOperation: 'allow',
  restrict: ({ origin }) => origin !== denied,
});

describe('merge writes under the interaction policy', () => {
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
    workbook.attachHistory(history);
    mutators.setRange(store, RANGE);
    const merged = (): boolean => store.getState().merges.byAnchor.size > 0;
    const run = (action: MergeAction) =>
      applyMergeAction(
        { store, workbook, history, strings: defaultStrings, origin: 'ribbon', commandId: 'merge' },
        action,
      );
    const commitDialog = (entry: DialogEntry, action: 'merge' | 'unmerge') => {
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
        action: { patch: {} },
        merge: { action, range: RANGE },
        entry,
      });
    };
    return { store, history, controller, merged, run, commitDialog };
  };

  describe('merge controls', () => {
    it('merges under a policy and replays through undo and redo', async () => {
      const { history, controller, merged, run } = await setup();
      controller.setPolicy(allowAll);
      await expect(run('mergeCells')).resolves.toBe(true);
      expect(merged()).toBe(true);
      expect(history.undo()).toBe(true);
      expect(merged()).toBe(false);
      expect(history.redo()).toBe(true);
      expect(merged()).toBe(true);
    });

    it('refuses every merge action when the policy denies merge', async () => {
      const { store, history, controller, merged, run } = await setup();
      controller.setPolicy(denyMerge);
      for (const action of ['mergeCells', 'mergeCenter', 'mergeAcross'] as const) {
        await expect(run(action)).resolves.toBe(false);
      }
      expect(merged()).toBe(false);
      controller.setPolicy(undefined);
      mutators.mergeRange(store, RANGE);
      controller.setPolicy(denyMerge);
      await expect(run('unmergeCells')).resolves.toBe(false);
      expect(merged()).toBe(true);
      expect(history.canUndo()).toBe(false);
    });

    it('refuses Merge & Center when the policy denies its alignment format', async () => {
      const { controller, merged, run } = await setup();
      controller.setPolicy({ defaultOperation: 'allow', operations: { format: false } });
      await expect(run('mergeCenter')).resolves.toBe(false);
      expect(merged()).toBe(false);
      await expect(run('mergeCells')).resolves.toBe(true);
    });

    it('refuses on locked cells of a protected sheet, with or without a policy', async () => {
      const { store, controller, merged, run } = await setup();
      setProtectedSheet(store, 0, true);
      await expect(run('mergeCells')).resolves.toBe(false);
      controller.setPolicy(allowAll);
      await expect(run('mergeCells')).resolves.toBe(false);
      expect(merged()).toBe(false);
    });

    it('refuses replay once the policy denies the undo or redo', async () => {
      const { history, controller, merged, run } = await setup();
      controller.setPolicy(allowAll);
      await run('mergeCells');
      controller.setPolicy(denyOrigin('undo'));
      expect(history.undo()).toBe(false);
      expect(merged()).toBe(true);
      controller.setPolicy(allowAll);
      expect(history.undo()).toBe(true);
      controller.setPolicy(denyMerge);
      expect(history.redo()).toBe(false);
      expect(merged()).toBe(false);
    });
  });

  describe('Format Cells alignment merge', () => {
    it('merges under a policy, decided by the entry that opened the dialog', async () => {
      const { history, controller, merged, commitDialog } = await setup();
      controller.setPolicy(denyOrigin('contextMenu'));
      expect(commitDialog({ origin: 'contextMenu', commandId: 'formatCells' }, 'merge')).toBe(
        false,
      );
      expect(merged()).toBe(false);
      expect(commitDialog({ origin: 'keyboard', commandId: 'formatCells' }, 'merge')).toBe(true);
      expect(merged()).toBe(true);
      expect(history.undo()).toBe(true);
      expect(merged()).toBe(false);
      expect(history.redo()).toBe(true);
      expect(merged()).toBe(true);
    });

    it('refuses the merge when the policy denies it, and replay after it changes', async () => {
      const { history, controller, merged, commitDialog } = await setup();
      const entry: DialogEntry = { origin: 'ribbon', commandId: 'formatCells' };
      controller.setPolicy(denyMerge);
      expect(commitDialog(entry, 'merge')).toBe(false);
      expect(history.canUndo()).toBe(false);
      controller.setPolicy(allowAll);
      expect(commitDialog(entry, 'merge')).toBe(true);
      controller.setPolicy(denyMerge);
      expect(history.undo()).toBe(false);
      expect(merged()).toBe(true);
    });

    it('refuses an unmerge on locked cells of a protected sheet', async () => {
      const { store, controller, merged, commitDialog } = await setup();
      mutators.mergeRange(store, RANGE);
      controller.setPolicy(allowAll);
      setProtectedSheet(store, 0, true);
      expect(commitDialog({ origin: 'ribbon', commandId: 'formatCells' }, 'unmerge')).toBe(false);
      expect(merged()).toBe(true);
    });
  });
});
