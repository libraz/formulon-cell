import { afterEach, describe, expect, it } from 'vitest';
import { commentAt, setComment } from '../../../src/commands/comment.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import type { InteractionPolicy } from '../../../src/commands/interaction-policy.js';
import { setCellLocked, setProtectedSheet } from '../../../src/commands/protection.js';
import { executeRibbonClearAction } from '../../../src/commands/ribbon-clear.js';
import { executeRibbonCommentAction } from '../../../src/commands/ribbon-comment.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachCommentDialog } from '../../../src/interact/comment-dialog.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const A1 = { sheet: 0, row: 0, col: 0 };
const B1 = { sheet: 0, row: 0, col: 1 };
const allowAll: InteractionPolicy = { defaultOperation: 'allow' };
const denyComment: InteractionPolicy = {
  defaultOperation: 'allow',
  operations: { comment: false },
};
const denyOrigin = (denied: string): InteractionPolicy => ({
  defaultOperation: 'allow',
  restrict: ({ origin }) => origin !== denied,
});

describe('note and comment writes under the interaction policy', () => {
  const cleanups: (() => void)[] = [];

  afterEach(() => {
    for (const cleanup of cleanups.splice(0)) cleanup();
    document.body.replaceChildren();
  });

  const setup = async () => {
    const workbook = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    const history = new History();
    const controller = new InteractionController({ store, getWb: () => workbook, history });
    registerInteractionController(store, controller);
    cleanups.push(() => {
      controller.dispose();
      workbook.dispose();
    });
    setComment(store, A1, 'first', workbook);
    setComment(store, B1, 'second', workbook);
    mutators.setActive(store, A1);
    const del = (action: 'delete-active' | 'delete-all') =>
      executeRibbonCommentAction({
        store,
        workbook,
        history,
        action,
        origin: 'ribbon',
        commandId: 'deleteCommentReview',
      });
    return { workbook, store, history, controller, del };
  };

  it('refuses deleting notes the policy denies, and replays an allowed delete', async () => {
    const { store, history, controller, del } = await setup();
    controller.setPolicy(denyComment);
    del('delete-active');
    expect(commentAt(store.getState(), A1)).toBe('first');
    controller.setPolicy(allowAll);
    del('delete-all');
    expect(commentAt(store.getState(), A1)).toBeNull();
    expect(commentAt(store.getState(), B1)).toBeNull();
    controller.setPolicy(denyOrigin('undo'));
    expect(history.undo()).toBe(false);
    controller.setPolicy(allowAll);
    expect(history.undo()).toBe(true);
    expect(commentAt(store.getState(), B1)).toBe('second');
  });

  for (const [label, policy] of [
    ['without a policy', undefined],
    ['under an allow-all policy', allowAll],
  ] as const satisfies readonly (readonly [string, InteractionPolicy | undefined])[]) {
    it(`decides notes on a protected sheet by Edit objects, not the cell's lock ${label}`, async () => {
      const { workbook, store, history, controller, del } = await setup();
      controller.setPolicy(policy);
      setCellLocked(store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }, false);
      setProtectedSheet(store, 0, true);
      // Without Edit objects no note changes, the unlocked cell's included; no alert either.
      del('delete-active');
      expect(commentAt(store.getState(), A1)).toBe('first');
      expect(setComment(store, B1, 'edited', workbook)).toBe(false);
      expect(setComment(store, { sheet: 0, row: 5, col: 5 }, 'new', workbook)).toBe(false);
      mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
      executeRibbonClearAction({ store, workbook, history, action: 'comments' });
      expect(commentAt(store.getState(), A1)).toBe('first');
      expect(commentAt(store.getState(), B1)).toBe('second');
      // With Edit objects every note changes, the locked cell's included.
      setProtectedSheet(store, 0, true, { permissions: { objects: true } });
      expect(setComment(store, A1, 'edited', workbook)).toBe(true);
      executeRibbonClearAction({ store, workbook, history, action: 'comments' });
      expect(commentAt(store.getState(), A1)).toBeNull();
      expect(commentAt(store.getState(), B1)).toBeNull();
    });
  }

  it('commits the note dialog under the entry point that opened it', async () => {
    const { workbook, store, history, controller } = await setup();
    const host = document.createElement('div');
    document.body.appendChild(host);
    const dialog = attachCommentDialog({ host, store, history, getWb: () => workbook });
    cleanups.push(() => dialog.detach());
    const save = (text: string): void => {
      const area = document.querySelector<HTMLTextAreaElement>('.fc-cmtnote__textarea');
      if (area) area.value = text;
      document.querySelector<HTMLButtonElement>('.fc-cmtnote__btn--primary')?.click();
    };
    controller.setPolicy(denyOrigin('contextMenu'));
    dialog.open({ mode: 'note', entry: { origin: 'contextMenu', commandId: 'insertComment' } });
    save('from the menu');
    expect(commentAt(store.getState(), A1)).toBe('first');
    dialog.close();
    dialog.open({ mode: 'note', entry: { origin: 'keyboard', commandId: 'insertComment' } });
    save('from Shift+F2');
    expect(commentAt(store.getState(), A1)).toBe('from Shift+F2');
  });
});
