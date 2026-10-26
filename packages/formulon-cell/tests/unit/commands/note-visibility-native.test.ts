// @vitest-environment node

import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { deleteCells, insertCells } from '../../../src/commands/cell-shift.js';
import { pasteSpecial } from '../../../src/commands/clipboard/paste-special.js';
import { captureSnapshot } from '../../../src/commands/clipboard/snapshot.js';
import {
  allNotesVisible,
  clearComment,
  noteVisibleAt,
  setComment,
} from '../../../src/commands/comment.js';
import { History } from '../../../src/commands/history.js';
import { executeNoteVisibilityAction } from '../../../src/commands/ribbon-comment.js';
import { sortRange } from '../../../src/commands/sort.js';
import { hydrateCommentsAndHyperlinksFromEngine } from '../../../src/engine/format-sync.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const a = (row: number, col: number) => ({ sheet: 0, row, col });

describe('note visibility (native engine)', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let history: History;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await WorkbookHandle.createDefault();
    history = new History();
    expect(wb.isStub).toBe(false);
    expect(wb.capabilities.commentVisibility).toBe(true);
  });

  afterEach(() => {
    wb.dispose();
  });

  it('round-trips a shown note through xlsx and hydrates it into the store', async () => {
    expect(wb.setCommentEntry(0, 1, 1, 'Alice', 'shown', true)).toBe(true);
    expect(wb.setCommentEntry(0, 2, 2, 'Bob', 'hidden')).toBe(true);
    expect(wb.getComment(0, 1, 1)).toEqual({ author: 'Alice', text: 'shown', visible: true });
    expect(wb.getComment(0, 2, 2)).toEqual({ author: 'Bob', text: 'hidden' });

    const reloaded = await WorkbookHandle.loadBytes(wb.save());
    try {
      expect(reloaded.getCommentVisible(0, 1, 1)).toBe(true);
      expect(reloaded.getCommentVisible(0, 2, 2)).toBe(false);
      const loaded = createSpreadsheetStore();
      hydrateCommentsAndHyperlinksFromEngine(reloaded, loaded, 0);
      expect(loaded.getState().format.formats.get(addrKey(a(1, 1)))).toMatchObject({
        comment: 'shown',
        commentVisible: true,
      });
      expect(loaded.getState().format.formats.get(addrKey(a(2, 2)))?.commentVisible).toBe(
        undefined,
      );
    } finally {
      reloaded.dispose();
    }
  });

  it('Show/Hide Note toggles the active note without an undo step, and editing keeps it shown', () => {
    setComment(store, a(0, 0), 'note', wb);
    mutators.setActive(store, a(0, 0));
    const undoDepth = history.canUndo();
    executeNoteVisibilityAction({ store, workbook: wb, action: 'toggle-active' });
    expect(noteVisibleAt(store.getState(), a(0, 0))).toBe(true);
    expect(wb.getCommentVisible(0, 0, 0)).toBe(true);
    expect(history.canUndo()).toBe(undoDepth);

    setComment(store, a(0, 0), 'edited', wb);
    expect(noteVisibleAt(store.getState(), a(0, 0))).toBe(true);
    expect(wb.getCommentVisible(0, 0, 0)).toBe(true);

    executeNoteVisibilityAction({ store, workbook: wb, action: 'toggle-active' });
    expect(noteVisibleAt(store.getState(), a(0, 0))).toBe(false);
    expect(wb.getCommentVisible(0, 0, 0)).toBe(false);

    executeNoteVisibilityAction({ store, workbook: wb, action: 'toggle-active' });
    clearComment(store, a(0, 0), wb);
    expect(store.getState().format.formats.get(addrKey(a(0, 0)))?.commentVisible).toBe(undefined);
  });

  it('Show All Notes shows every note, then hides them all', () => {
    setComment(store, a(0, 0), 'one', wb);
    setComment(store, a(3, 2), 'two', wb);
    executeNoteVisibilityAction({ store, workbook: wb, action: 'toggle-all' });
    expect(allNotesVisible(store.getState())).toBe(true);
    expect(wb.getCommentVisible(0, 3, 2)).toBe(true);
    executeNoteVisibilityAction({ store, workbook: wb, action: 'toggle-all' });
    expect(allNotesVisible(store.getState())).toBe(false);
    expect(wb.getCommentVisible(0, 0, 0)).toBe(false);
    expect(wb.getCommentVisible(0, 3, 2)).toBe(false);
  });

  it('keeps a shown note shown when sort moves it', () => {
    wb.setNumber(a(0, 0), 2);
    wb.setNumber(a(1, 0), 1);
    wb.recalc();
    mutators.replaceCells(store, wb.cells(0));
    expect(wb.setCommentEntry(0, 0, 0, 'Alice', 'two', true)).toBe(true);
    expect(
      sortRange(
        store.getState(),
        store,
        wb,
        { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 },
        { byCol: 0, direction: 'asc' },
        history,
      ),
    ).toBe(true);
    expect(wb.getComment(0, 1, 0)).toEqual({ author: 'Alice', text: 'two', visible: true });
    expect(wb.getComment(0, 0, 0)).toBeNull();
    expect(noteVisibleAt(store.getState(), a(1, 0))).toBe(true);
  });

  it('keeps a shown note shown through insert/delete cells and their undo', () => {
    expect(wb.setCommentEntry(0, 1, 2, 'Alice', 'note', true)).toBe(true);
    wb.attachHistory(history);
    const at = { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 };
    expect(insertCells(store, wb, history, at, 'down')).toBe(true);
    expect(wb.getComment(0, 2, 2)).toEqual({ author: 'Alice', text: 'note', visible: true });
    expect(history.undo()).toBe(true);
    expect(wb.getComment(0, 1, 2)).toEqual({ author: 'Alice', text: 'note', visible: true });
    expect(deleteCells(store, wb, history, { ...at, r0: 0, r1: 0 }, 'up')).toBe(true);
    expect(wb.getComment(0, 0, 2)).toEqual({ author: 'Alice', text: 'note', visible: true });
  });

  it('Paste Special All copies the shown state with the note', () => {
    mutators.setCellFormat(store, a(0, 0), {
      comment: 'src',
      commentAuthor: 'Alice',
      commentVisible: true,
    });
    expect(wb.setCommentEntry(0, 0, 0, 'Alice', 'src', true)).toBe(true);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    if (!snap) throw new Error('expected snapshot');
    mutators.setActive(store, a(4, 4));
    mutators.setRange(store, { sheet: 0, r0: 4, c0: 4, r1: 4, c1: 4 });
    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    expect(wb.getComment(0, 4, 4)).toEqual({ author: 'Alice', text: 'src', visible: true });
    expect(noteVisibleAt(store.getState(), a(4, 4))).toBe(true);
  });
});
