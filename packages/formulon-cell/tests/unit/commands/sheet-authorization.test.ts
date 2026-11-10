import { afterEach, describe, expect, it } from 'vitest';
import { canExecuteBuiltIn } from '../../../src/commands/built-in-command-policy.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import type { InteractionPolicy } from '../../../src/commands/interaction-policy.js';
import {
  setProtectedSheet,
  setWorkbookStructureProtected,
} from '../../../src/commands/protection.js';
import {
  deleteSheetAction,
  insertSheetAction,
  moveSheetAction,
  renameSheetAction,
  type SheetActionContext,
  setSheetHiddenAction,
  setSheetTabColorAction,
} from '../../../src/commands/sheet-actions.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const SHEET_IDS = [
  'insertSheet',
  'deleteSheet',
  'renameSheet',
  'moveSheet',
  'hideSheet',
  'unhideSheet',
  'sheetTabColor',
] as const;

const allowAll: InteractionPolicy = { defaultOperation: 'allow' };
const denySheet: InteractionPolicy = { defaultOperation: 'allow', operations: { sheet: false } };
const denyOrigin = (denied: string): InteractionPolicy => ({
  defaultOperation: 'allow',
  restrict: ({ origin }) => origin !== denied,
});

describe('sheet commands under the interaction policy', () => {
  const handles: WorkbookHandle[] = [];
  const controllers: InteractionController[] = [];

  afterEach(() => {
    for (const controller of controllers.splice(0)) controller.dispose();
    for (const handle of handles.splice(0)) handle.dispose();
  });

  const setup = async () => {
    const workbook = await WorkbookHandle.createDefault();
    handles.push(workbook);
    workbook.addSheet('Second');
    const store = createSpreadsheetStore();
    const history = new History();
    const controller = new InteractionController({ store, getWb: () => workbook, history });
    controllers.push(controller);
    registerInteractionController(store, controller);
    workbook.attachHistory(history);
    const ctx: SheetActionContext = { store, workbook, history, origin: 'contextMenu' };
    const names = (): string[] =>
      Array.from({ length: workbook.sheetCount }, (_, i) => workbook.sheetName(i));
    return { store, workbook, history, controller, ctx, names };
  };

  it('runs every sheet command under a policy, undoable where the command is', async () => {
    const { store, history, controller, ctx, names } = await setup();
    controller.setPolicy(allowAll);
    expect(renameSheetAction(ctx, 0, 'Data')).toBe(true);
    expect(names()).toEqual(['Data', 'Second']);
    expect(history.undo()).toBe(true);
    expect(names()[0]).not.toBe('Data');
    expect(history.redo()).toBe(true);
    expect(names()).toEqual(['Data', 'Second']);

    expect(moveSheetAction(ctx, 0, 1)).toBe(true);
    expect(names()).toEqual(['Second', 'Data']);
    expect(history.undo()).toBe(true);
    expect(names()).toEqual(['Data', 'Second']);

    expect(setSheetTabColorAction(ctx, 1, '#ff0000')).toBe(true);
    expect(store.getState().layout.sheetTabColors.get(1)).toBe('#ff0000');
    expect(history.undo()).toBe(true);
    expect(store.getState().layout.sheetTabColors.has(1)).toBe(false);

    expect(setSheetHiddenAction(ctx, 1, true)).toBe(true);
    expect(store.getState().layout.hiddenSheets.has(1)).toBe(true);
    expect(setSheetHiddenAction(ctx, 1, false)).toBe(true);
    expect(store.getState().layout.hiddenSheets.has(1)).toBe(false);

    const added = insertSheetAction(ctx);
    expect(added).toBe(2);
    expect(deleteSheetAction(ctx, added)).toBe(true);
    expect(names()).toEqual(['Data', 'Second']);
  });

  it('refuses every sheet command when the policy denies the sheet operation', async () => {
    const { store, history, controller, ctx, names } = await setup();
    controller.setPolicy(denySheet);
    expect(renameSheetAction(ctx, 0, 'Data')).toBe(false);
    expect(moveSheetAction(ctx, 0, 1)).toBe(false);
    expect(setSheetTabColorAction(ctx, 1, '#ff0000')).toBe(false);
    expect(setSheetHiddenAction(ctx, 1, true)).toBe(false);
    expect(insertSheetAction(ctx)).toBe(-1);
    expect(deleteSheetAction(ctx, 1)).toBe(false);
    expect(names()).toHaveLength(2);
    expect(names()[0]).not.toBe('Data');
    expect(store.getState().layout.sheetTabColors.size).toBe(0);
    expect(history.canUndo()).toBe(false);
    for (const id of SHEET_IDS) {
      expect(canExecuteBuiltIn(store, id, 'contextMenu').allowed, id).toBe(false);
    }
  });

  for (const [label, policy] of [
    ['without a policy', undefined],
    ['under an allow-all policy', allowAll],
  ] as const satisfies readonly (readonly [string, InteractionPolicy | undefined])[]) {
    it(`follows workbook-structure protection, not sheet protection, ${label}`, async () => {
      const { store, controller, ctx, names } = await setup();
      controller.setPolicy(policy);
      setProtectedSheet(store, 0, true);
      setProtectedSheet(store, 1, true);
      for (const id of SHEET_IDS) {
        expect(canExecuteBuiltIn(store, id, 'contextMenu').allowed, id).toBe(true);
      }
      expect(renameSheetAction(ctx, 0, 'Data')).toBe(true);

      setWorkbookStructureProtected(store, true);
      for (const id of SHEET_IDS) {
        expect(canExecuteBuiltIn(store, id, 'contextMenu'), id).toMatchObject({
          allowed: false,
          code: 'protected',
        });
      }
      expect(renameSheetAction(ctx, 0, 'Other')).toBe(false);
      expect(moveSheetAction(ctx, 0, 1)).toBe(false);
      expect(setSheetTabColorAction(ctx, 0, '#00ff00')).toBe(false);
      expect(insertSheetAction(ctx)).toBe(-1);
      expect(names()).toEqual(['Data', 'Second']);
    });
  }

  it('refuses replay once the policy denies the undo or redo', async () => {
    const { history, controller, ctx, names } = await setup();
    controller.setPolicy(allowAll);
    expect(renameSheetAction(ctx, 0, 'Data')).toBe(true);
    controller.setPolicy(denyOrigin('undo'));
    expect(history.undo()).toBe(false);
    expect(names()[0]).toBe('Data');
    controller.setPolicy(allowAll);
    expect(history.undo()).toBe(true);
    controller.setPolicy(denySheet);
    expect(history.redo()).toBe(false);
    expect(names()[0]).not.toBe('Data');
  });
});
