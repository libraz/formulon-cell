import { describe, expect, it, vi } from 'vitest';

import { commentAt, setComment } from '../../../src/commands/comment.js';
import { History } from '../../../src/commands/history.js';
import { executeRibbonCommentAction } from '../../../src/commands/ribbon-comment.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const commentWorkbook = (): WorkbookHandle & {
  setCommentEntry: ReturnType<typeof vi.fn>;
} =>
  ({
    capabilities: { comments: true },
    setCommentEntry: vi.fn(),
  }) as unknown as WorkbookHandle & { setCommentEntry: ReturnType<typeof vi.fn> };

describe('executeRibbonCommentAction', () => {
  it('deletes only the active cell comment and records a single undoable change', () => {
    const store = createSpreadsheetStore();
    const wb = commentWorkbook();
    const history = new History();
    const active = { sheet: 0, row: 1, col: 1 };
    const other = { sheet: 0, row: 2, col: 2 };
    setComment(store, active, 'active', wb);
    setComment(store, other, 'other', wb);
    store.setState((state) => ({
      ...state,
      selection: { ...state.selection, active },
    }));
    wb.setCommentEntry.mockClear();

    executeRibbonCommentAction({
      store,
      workbook: wb,
      history,
      action: 'delete-active',
    });

    expect(commentAt(store.getState(), active)).toBeNull();
    expect(commentAt(store.getState(), other)).toBe('other');
    expect(wb.setCommentEntry).toHaveBeenCalledWith(0, 1, 1, '', '');

    expect(history.undo()).toBe(true);
    expect(commentAt(store.getState(), active)).toBe('active');
    expect(commentAt(store.getState(), other)).toBe('other');
  });

  it('is a no-op when deleting active comment from an uncommented cell', () => {
    const store = createSpreadsheetStore();
    const wb = commentWorkbook();
    const history = new History();

    executeRibbonCommentAction({
      store,
      workbook: wb,
      history,
      action: 'delete-active',
    });

    expect(wb.setCommentEntry).not.toHaveBeenCalled();
    expect(history.canUndo()).toBe(false);
  });

  it('deletes every comment on the active sheet without touching other sheets', () => {
    const store = createSpreadsheetStore();
    const wb = commentWorkbook();
    const history = new History();
    const first = { sheet: 0, row: 0, col: 0 };
    const second = { sheet: 0, row: 2, col: 1 };
    const otherSheet = { sheet: 1, row: 0, col: 0 };
    setComment(store, first, 'first', wb);
    setComment(store, second, 'second', wb);
    setComment(store, otherSheet, 'other sheet', wb);
    wb.setCommentEntry.mockClear();

    executeRibbonCommentAction({
      store,
      workbook: wb,
      history,
      action: 'delete-all',
    });

    expect(commentAt(store.getState(), first)).toBeNull();
    expect(commentAt(store.getState(), second)).toBeNull();
    expect(commentAt(store.getState(), otherSheet)).toBe('other sheet');
    expect(wb.setCommentEntry.mock.calls.map((call) => call.slice(0, 5))).toEqual([
      [0, 0, 0, '', ''],
      [0, 2, 1, '', ''],
    ]);

    expect(history.undo()).toBe(true);
    expect(commentAt(store.getState(), first)).toBe('first');
    expect(commentAt(store.getState(), second)).toBe('second');
    expect(commentAt(store.getState(), otherSheet)).toBe('other sheet');
  });
});
