// @vitest-environment node

import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { deleteCells, insertCells } from '../../../src/commands/cell-shift.js';
import { History } from '../../../src/commands/history.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

describe('native cell shift sheet-reference parity', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await WorkbookHandle.createDefault();
    expect(wb.isStub).toBe(false);
  });

  afterEach(() => {
    wb.dispose();
  });

  it('updates external formulas through insert/delete and one-step undo/redo', () => {
    expect(wb.addSheet('Target')).toBe(1);
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    wb.setNumber({ sheet: 0, row: 1, col: 0 }, 2);
    wb.setNumber({ sheet: 0, row: 2, col: 0 }, 3);
    wb.setFormula({ sheet: 0, row: 0, col: 2 }, '=Sheet1!$A$2');
    wb.setFormula({ sheet: 1, row: 0, col: 3 }, '=Sheet1!$A$2');
    wb.setFormula({ sheet: 1, row: 1, col: 3 }, '=SUM(Sheet1!A1:A3)');
    wb.setFormula({ sheet: 1, row: 2, col: 3 }, '=A2');

    const history = new History();
    wb.attachHistory(history);
    const insertRange = { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 0 };
    expect(insertCells(store, wb, history, insertRange, 'down')).toBe(true);
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 2 })).toBe('=Sheet1!$A$3');
    expect(wb.cellFormula({ sheet: 1, row: 0, col: 3 })).toBe('=Sheet1!$A$3');
    expect(wb.cellFormula({ sheet: 1, row: 1, col: 3 })).toBe('=SUM(Sheet1!A1:A4)');
    expect(wb.cellFormula({ sheet: 1, row: 2, col: 3 })).toBe('=A2');

    expect(history.undo()).toBe(true);
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 2 })).toBe('=Sheet1!$A$2');
    expect(wb.cellFormula({ sheet: 1, row: 0, col: 3 })).toBe('=Sheet1!$A$2');
    expect(wb.cellFormula({ sheet: 1, row: 1, col: 3 })).toBe('=SUM(Sheet1!A1:A3)');
    expect(history.redo()).toBe(true);

    const deleteRange = { sheet: 0, r0: 2, c0: 0, r1: 2, c1: 0 };
    expect(deleteCells(store, wb, history, deleteRange, 'up')).toBe(true);
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 2 })).toBe('=Sheet1!#REF!');
    expect(wb.cellFormula({ sheet: 1, row: 0, col: 3 })).toBe('=Sheet1!#REF!');
    expect(wb.cellFormula({ sheet: 1, row: 1, col: 3 })).toBe('=SUM(Sheet1!A1:A3)');
    expect(history.undo()).toBe(true);
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 2 })).toBe('=Sheet1!$A$3');
    expect(wb.cellFormula({ sheet: 1, row: 1, col: 3 })).toBe('=SUM(Sheet1!A1:A4)');
    expect(history.redo()).toBe(true);
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 2 })).toBe('=Sheet1!#REF!');
    expect(wb.cellFormula({ sheet: 1, row: 1, col: 3 })).toBe('=SUM(Sheet1!A1:A3)');
  });

  it('moves engine comments with cells and persists their undo state', async () => {
    expect(wb.capabilities.comments).toBe(true);
    expect(wb.capabilities.commentsEnumerable).toBe(true);
    const note = { sheet: 0, row: 1, col: 2 };
    expect(wb.setCommentEntry(note.sheet, note.row, note.col, 'Alice', 'note')).toBe(true);
    expect(store.getState().format.formats.get(`0:${note.row}:${note.col}`)).toBeUndefined();
    expect(wb.getComment(note.sheet, note.row, note.col)).toEqual({
      author: 'Alice',
      text: 'note',
    });

    const history = new History();
    wb.attachHistory(history);
    const insertRange = { sheet: 0, r0: note.row, c0: note.col, r1: note.row, c1: note.col };
    expect(insertCells(store, wb, history, insertRange, 'down')).toBe(true);
    const moved = { ...note, row: note.row + 1 };
    expect(wb.getComment(note.sheet, note.row, note.col)).toBeNull();
    expect(wb.getComment(moved.sheet, moved.row, moved.col)).toEqual({
      author: 'Alice',
      text: 'note',
    });
    expect(wb.getComments(0)).toContainEqual({
      row: moved.row,
      col: moved.col,
      author: 'Alice',
      text: 'note',
    });
    expect(store.getState().format.formats.get(`0:${moved.row}:${moved.col}`)).toMatchObject({
      comment: 'note',
      commentAuthor: 'Alice',
    });

    const reloaded = await WorkbookHandle.loadBytes(wb.save());
    try {
      expect(reloaded.getComment(moved.sheet, moved.row, moved.col)).toEqual({
        author: 'Alice',
        text: 'note',
      });
    } finally {
      reloaded.dispose();
    }

    expect(history.undo()).toBe(true);
    expect(wb.getComment(note.sheet, note.row, note.col)).toEqual({
      author: 'Alice',
      text: 'note',
    });
    expect(wb.getComment(moved.sheet, moved.row, moved.col)).toBeNull();
    expect(history.redo()).toBe(true);
    expect(wb.getComment(moved.sheet, moved.row, moved.col)).toEqual({
      author: 'Alice',
      text: 'note',
    });

    const deleteRange = { sheet: 0, r0: moved.row, c0: moved.col, r1: moved.row, c1: moved.col };
    expect(deleteCells(store, wb, history, deleteRange, 'up')).toBe(true);
    expect(wb.getComment(moved.sheet, moved.row, moved.col)).toBeNull();
    expect(wb.getComments(0)).not.toContainEqual({
      row: moved.row,
      col: moved.col,
      author: 'Alice',
      text: 'note',
    });
    expect(history.undo()).toBe(true);
    expect(wb.getComment(moved.sheet, moved.row, moved.col)).toEqual({
      author: 'Alice',
      text: 'note',
    });
    expect(history.redo()).toBe(true);
    expect(wb.getComment(moved.sheet, moved.row, moved.col)).toBeNull();
  });

  it('rejects last-column comment, format, and merge overflow before mutation', () => {
    const history = new History();
    const lastColumn = { sheet: 0, row: 2, col: 16383 };
    mutators.setCellFormat(store, lastColumn, { italic: true, comment: 'edge' });
    expect(wb.setCommentEntry(0, lastColumn.row, lastColumn.col, 'Alice', 'edge')).toBe(true);
    expect(insertCells(store, wb, history, { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 }, 'right')).toBe(
      false,
    );
    expect(wb.getComment(0, lastColumn.row, lastColumn.col)).toEqual({
      author: 'Alice',
      text: 'edge',
    });
    expect(store.getState().format.formats.get('0:2:16383')).toMatchObject({
      italic: true,
      comment: 'edge',
    });
    expect(history.canUndo()).toBe(false);

    const engineOnly = { sheet: 0, row: 4, col: 16383 };
    expect(wb.setCommentEntry(0, engineOnly.row, engineOnly.col, 'Bob', 'engine-only')).toBe(true);
    expect(store.getState().format.formats.get('0:4:16383')).toBeUndefined();
    expect(
      insertCells(store, wb, history, { sheet: 0, r0: 4, c0: 16382, r1: 4, c1: 16382 }, 'right'),
    ).toBe(false);
    expect(wb.getComment(0, engineOnly.row, engineOnly.col)).toEqual({
      author: 'Bob',
      text: 'engine-only',
    });
    expect(history.canUndo()).toBe(false);

    const merge = { sheet: 0, r0: 5, c0: 16382, r1: 5, c1: 16383 };
    mutators.mergeRange(store, merge);
    expect(wb.engineAddMerge(0, merge)).toBe(true);
    expect(
      insertCells(store, wb, history, { sheet: 0, r0: 5, c0: 16381, r1: 5, c1: 16381 }, 'right'),
    ).toBe(false);
    expect(wb.getMerges(0)).toContainEqual(merge);
    expect(history.canUndo()).toBe(false);
  });
});
