// @vitest-environment node

import { describe, expect, it, vi } from 'vitest';
import { recordDialogFormatChange } from '../../../src/commands/dialog-format-history.js';
import { applySelectionFormatAction } from '../../../src/commands/format.js';
import { History } from '../../../src/commands/history.js';
import { flushFormatToEngine } from '../../../src/engine/cell-format-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const target = { sheet: 0, row: 0, col: 0 };

const fontAt = (
  workbook: WorkbookHandle,
  addr = target,
): ReturnType<WorkbookHandle['getFontRecord']> => {
  const xfIndex = workbook.getCellXfIndex(addr.sheet, addr.row, addr.col);
  const xf = workbook.getCellXf(xfIndex ?? 0);
  return xf ? workbook.getFontRecord(xf.fontIndex) : null;
};

describe('Format Cells dialog native engine contract', () => {
  it('publishes XF, hyperlink, validation, and comments through one scoped history step', async () => {
    const workbook = await WorkbookHandle.createDefault();
    expect(workbook.isStub).toBe(false);
    expect(workbook.capabilities.comments).toBe(true);
    expect(workbook.capabilities.hyperlinks).toBe(true);
    expect(workbook.capabilities.dataValidation).toBe(true);
    const store = createSpreadsheetStore();
    const history = new History();

    try {
      workbook.setNumber(target, 7);
      expect(workbook.addHyperlink(0, 0, 0, 'https://old.example', 'Old', 'old tip')).toBe(true);
      expect(workbook.setCommentEntry(0, 0, 0, 'Alice', 'old note')).toBe(true);
      mutators.setCellFormat(store, target, {
        bold: true,
        fontSize: 12,
        hyperlink: 'https://old.example',
        hyperlinkDisplay: 'Old',
        hyperlinkTooltip: 'old tip',
        comment: 'old note',
        commentAuthor: 'Alice',
        validation: { kind: 'list', source: ['A', 'B'] },
      });
      flushFormatToEngine(workbook, store, 0);
      expect(fontAt(workbook)?.bold).toBe(true);
      expect(fontAt(workbook)?.size).toBe(12);

      const after = {
        bold: false,
        fontSize: 7,
        hyperlink: 'https://new.example',
        comment: 'new note',
        commentAuthor: 'Bob',
        validation: { kind: 'list' as const, source: ['X', 'Y'] },
      };
      expect(
        recordDialogFormatChange({
          history,
          store,
          workbook,
          sheet: 0,
          targets: [target],
          pendingBefore: null,
          mutate: () =>
            applySelectionFormatAction(
              store.getState(),
              store,
              { patch: after },
              {
                allowPending: false,
              },
            ),
        }),
      ).toBe(true);

      expect(fontAt(workbook)?.bold).toBe(false);
      expect(fontAt(workbook)?.size).toBe(7);
      expect(workbook.getHyperlinks(0)).toContainEqual({
        row: 0,
        col: 0,
        target: 'https://new.example',
        display: '',
        tooltip: '',
      });
      expect(workbook.getValidationsForSheet(0)).toHaveLength(1);
      expect(workbook.getValidationsForSheet(0)[0]?.formula1).toContain('X');
      expect(workbook.getComment(0, 0, 0)).toEqual({ author: 'Bob', text: 'new note' });

      expect(history.undo()).toBe(true);
      expect(fontAt(workbook)?.bold).toBe(true);
      expect(fontAt(workbook)?.size).toBe(12);
      expect(workbook.getHyperlinks(0)).toContainEqual({
        row: 0,
        col: 0,
        target: 'https://old.example',
        display: 'Old',
        tooltip: 'old tip',
      });
      expect(workbook.getValidationsForSheet(0)[0]?.formula1).toContain('A');
      expect(workbook.getComment(0, 0, 0)).toEqual({ author: 'Alice', text: 'old note' });

      expect(history.redo()).toBe(true);
      expect(fontAt(workbook)?.bold).toBe(false);
      expect(fontAt(workbook)?.size).toBe(7);
      expect(workbook.getComment(0, 0, 0)).toEqual({ author: 'Bob', text: 'new note' });
    } finally {
      workbook.dispose();
    }
  });

  it('does not rewrite divergent physical comments for visual-only edits', async () => {
    const workbook = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    const history = new History();
    try {
      workbook.setNumber(target, 1);
      expect(workbook.setCommentEntry(0, 0, 0, 'Engine', 'engine-only')).toBe(true);
      mutators.setCellFormat(store, target, { bold: true });
      flushFormatToEngine(workbook, store, 0);
      const setComment = vi.spyOn(workbook, 'setCommentEntry');

      expect(
        recordDialogFormatChange({
          history,
          store,
          workbook,
          sheet: 0,
          targets: [target],
          pendingBefore: null,
          mutate: () => {
            mutators.setCellFormat(store, target, { bold: false });
            return true;
          },
        }),
      ).toBe(true);
      expect(setComment).not.toHaveBeenCalled();
      expect(workbook.getComment(0, 0, 0)).toEqual({ author: 'Engine', text: 'engine-only' });
    } finally {
      workbook.dispose();
    }
  });

  it('writes only changed comments when a union contains visual-only targets', async () => {
    const workbook = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    const history = new History();
    const other = { sheet: 0, row: 0, col: 1 };
    try {
      workbook.setNumber(target, 1);
      workbook.setNumber(other, 2);
      expect(workbook.setCommentEntry(0, 0, 0, 'A', 'old A')).toBe(true);
      expect(workbook.setCommentEntry(0, 0, 1, 'B', 'old B')).toBe(true);
      mutators.setCellFormat(store, target, {
        bold: true,
        comment: 'old A',
        commentAuthor: 'A',
      });
      mutators.setCellFormat(store, other, {
        bold: true,
        comment: 'old B',
        commentAuthor: 'B',
      });
      flushFormatToEngine(workbook, store, 0);
      const setComment = vi.spyOn(workbook, 'setCommentEntry');

      expect(
        recordDialogFormatChange({
          history,
          store,
          workbook,
          sheet: 0,
          targets: [target, other],
          pendingBefore: null,
          mutate: () => {
            mutators.setCellFormat(store, target, {
              bold: false,
              comment: 'new A',
              commentAuthor: 'New A',
            });
            mutators.setCellFormat(store, other, { bold: false });
            return true;
          },
        }),
      ).toBe(true);
      expect(setComment).toHaveBeenCalledTimes(1);
      expect(workbook.getComment(0, 0, 0)).toEqual({ author: 'New A', text: 'new A' });
      expect(workbook.getComment(0, 0, 1)).toEqual({ author: 'B', text: 'old B' });
    } finally {
      workbook.dispose();
    }
  });

  it('restores every native XF after a later target setter returns false', async () => {
    const workbook = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    const history = new History();
    const other = { sheet: 0, row: 0, col: 1 };
    const originalSetXf = workbook.setCellXfIndex.bind(workbook);
    const setXf = vi.spyOn(workbook, 'setCellXfIndex');
    try {
      workbook.setNumber(target, 1);
      workbook.setNumber(other, 2);
      mutators.setCellFormat(store, target, { bold: true });
      mutators.setCellFormat(store, other, { bold: true });
      flushFormatToEngine(workbook, store, 0);
      expect(fontAt(workbook, target)?.bold).toBe(true);
      expect(fontAt(workbook, other)?.bold).toBe(true);

      let failed = false;
      setXf.mockImplementation((sheet, row, col, xfIndex) => {
        if (!failed && row === other.row && col === other.col) {
          failed = true;
          return false;
        }
        return originalSetXf(sheet, row, col, xfIndex);
      });
      expect(() =>
        recordDialogFormatChange({
          history,
          store,
          workbook,
          sheet: 0,
          targets: [target, other],
          pendingBefore: null,
          mutate: () => {
            mutators.setCellFormat(store, target, { bold: false });
            mutators.setCellFormat(store, other, { bold: false });
            return true;
          },
        }),
      ).toThrow(/setCellXfIndex/);
      expect(store.getState().format.formats.get('0:0:0')?.bold).toBe(true);
      expect(store.getState().format.formats.get('0:0:1')?.bold).toBe(true);
      expect(fontAt(workbook, target)?.bold).toBe(true);
      expect(fontAt(workbook, other)?.bold).toBe(true);
      expect(history.canUndo()).toBe(false);
    } finally {
      setXf.mockRestore();
      workbook.dispose();
    }
  });

  it('restores physical comments when a forward comment setter returns false', async () => {
    const workbook = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    const history = new History();
    const setComment = vi.spyOn(workbook, 'setCommentEntry');
    try {
      workbook.setNumber(target, 1);
      expect(workbook.setCommentEntry(0, 0, 0, 'Alice', 'old')).toBe(true);
      mutators.setCellFormat(store, target, {
        comment: 'old',
        commentAuthor: 'Alice',
        bold: true,
      });
      flushFormatToEngine(workbook, store, 0);
      setComment.mockImplementationOnce(() => false);

      expect(() =>
        recordDialogFormatChange({
          history,
          store,
          workbook,
          sheet: 0,
          targets: [target],
          pendingBefore: null,
          mutate: () => {
            mutators.setCellFormat(store, target, {
              comment: 'new',
              commentAuthor: 'Bob',
              bold: false,
            });
            return true;
          },
        }),
      ).toThrow('dialog comment engine write failed');
      expect(store.getState().format.formats.get('0:0:0')).toMatchObject({
        comment: 'old',
        commentAuthor: 'Alice',
        bold: true,
      });
      expect(workbook.getComment(0, 0, 0)).toEqual({ author: 'Alice', text: 'old' });
      expect(history.canUndo()).toBe(false);
    } finally {
      setComment.mockRestore();
      workbook.dispose();
    }
  });

  it('restores the immediate native comment and format after a failed replay', async () => {
    const workbook = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    const history = new History();
    const setComment = vi.spyOn(workbook, 'setCommentEntry');
    try {
      workbook.setNumber(target, 1);
      expect(workbook.setCommentEntry(0, 0, 0, 'Alice', 'old')).toBe(true);
      mutators.setCellFormat(store, target, {
        comment: 'old',
        commentAuthor: 'Alice',
        bold: true,
      });
      flushFormatToEngine(workbook, store, 0);
      expect(
        recordDialogFormatChange({
          history,
          store,
          workbook,
          sheet: 0,
          targets: [target],
          pendingBefore: null,
          mutate: () => {
            mutators.setCellFormat(store, target, {
              comment: 'new',
              commentAuthor: 'Bob',
              bold: false,
            });
            return true;
          },
        }),
      ).toBe(true);
      expect(workbook.getComment(0, 0, 0)).toEqual({ author: 'Bob', text: 'new' });
      setComment.mockImplementationOnce(() => false);

      expect(() => history.undo()).toThrow('dialog comment engine write failed');
      expect(history.canUndo()).toBe(true);
      expect(store.getState().format.formats.get('0:0:0')).toMatchObject({
        comment: 'new',
        commentAuthor: 'Bob',
        bold: false,
      });
      expect(workbook.getComment(0, 0, 0)).toEqual({ author: 'Bob', text: 'new' });
    } finally {
      setComment.mockRestore();
      workbook.dispose();
    }
  });
});
