import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import {
  BUILT_IN_COMMAND_OPERATION,
  canExecuteBuiltIn,
} from '../../../src/commands/built-in-command-policy.js';
import { allNotesVisible, noteVisibleAt, setComment } from '../../../src/commands/comment.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { viewerPolicy } from '../../../src/commands/interaction-policy.js';
import { executeNoteVisibilityAction } from '../../../src/commands/ribbon-comment.js';
import { hydrateCommentsAndHyperlinksFromEngine } from '../../../src/engine/format-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';
import { cancelOpenAppDialogs } from '../../../src/toolbar/dialogs/shell.js';

const a = (row: number, col: number) => ({ sheet: 0, row, col });

const openDialog = (): Element | null =>
  document.querySelector('[role="dialog"], [role="alertdialog"]');

describe('Show/Hide Note is a view operation', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let history: History;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await WorkbookHandle.createDefault();
    history = new History();
  });

  afterEach(() => {
    cancelOpenAppDialogs();
    document.body.replaceChildren();
    wb.dispose();
  });

  it('shows a note on a protected sheet without an undo entry', () => {
    setComment(store, a(0, 0), 'note', wb);
    mutators.setSheetProtected(store, 0, true);
    mutators.setActive(store, a(0, 0));
    expect(executeNoteVisibilityAction({ store, workbook: wb, action: 'toggle-active' })).toBe(
      true,
    );
    expect(noteVisibleAt(store.getState(), a(0, 0))).toBe(true);
    expect(wb.getCommentVisible(0, 0, 0)).toBe(true);
    expect(history.canUndo()).toBe(false);
  });

  it('never asks for an allowed-range password', () => {
    setComment(store, a(0, 0), 'note', wb);
    mutators.addAllowedEditRange(store, {
      title: 'Guarded',
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      password: 'secret',
    });
    mutators.setSheetProtected(store, 0, true);
    executeNoteVisibilityAction({ store, workbook: wb, action: 'toggle-all' });
    expect(openDialog()).toBeNull();
    expect(noteVisibleAt(store.getState(), a(0, 0))).toBe(true);
  });

  it('Show All Notes shows, then hides again, including a blank-cell note from the file', async () => {
    expect(wb.setCommentEntry(0, 5, 5, 'Alice', 'loaded on a blank cell')).toBe(true);
    const loaded = await WorkbookHandle.loadBytes(wb.save());
    try {
      hydrateCommentsAndHyperlinksFromEngine(loaded, store, 0);
      setComment(store, a(1, 1), 'typed', loaded);
      executeNoteVisibilityAction({ store, workbook: loaded, action: 'toggle-all' });
      expect(allNotesVisible(store.getState())).toBe(true);
      expect(loaded.getCommentVisible(0, 5, 5)).toBe(true);
      executeNoteVisibilityAction({ store, workbook: loaded, action: 'toggle-all' });
      expect(allNotesVisible(store.getState())).toBe(false);
      expect(noteVisibleAt(store.getState(), a(5, 5))).toBe(false);
      expect(noteVisibleAt(store.getState(), a(1, 1))).toBe(false);
      expect(loaded.getCommentVisible(0, 5, 5)).toBe(false);
    } finally {
      loaded.dispose();
    }
  });

  it('runs under a read-only host policy', () => {
    setComment(store, a(0, 0), 'note', wb);
    const controller = new InteractionController({ store, getWb: () => wb, history });
    const unregister = registerInteractionController(store, controller);
    controller.setPolicy(viewerPolicy());
    for (const id of [
      'showHideNoteReview',
      'showAllNotesReview',
      'mac.review.showHideNote',
      'mac.review.showNotes',
    ]) {
      expect(BUILT_IN_COMMAND_OPERATION[id]).toBeUndefined();
      expect(canExecuteBuiltIn(store, id, 'ribbon').allowed).toBe(true);
    }
    mutators.setActive(store, a(0, 0));
    executeNoteVisibilityAction({ store, workbook: wb, action: 'toggle-active' });
    expect(noteVisibleAt(store.getState(), a(0, 0))).toBe(true);
    unregister();
  });
});
