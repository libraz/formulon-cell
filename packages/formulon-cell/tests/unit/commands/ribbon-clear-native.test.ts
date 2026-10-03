// @vitest-environment node

import { describe, expect, it } from 'vitest';
import { commentAt } from '../../../src/commands/comment.js';
import { History } from '../../../src/commands/history.js';
import { executeRibbonClearAction } from '../../../src/commands/ribbon-clear.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('executeRibbonClearAction native metadata replay', () => {
  it('clears and restores native values, formulas, comments, and hyperlinks as one All step', async () => {
    const workbook = await WorkbookHandle.createDefault();
    expect(workbook.isStub).toBe(false);
    expect(workbook.capabilities.comments).toBe(true);
    expect(workbook.capabilities.hyperlinks).toBe(true);
    const store = createSpreadsheetStore();
    const history = new History();
    const primary = { sheet: 0, row: 1, col: 1 };
    const extra = { sheet: 0, row: 20, col: 4 };

    try {
      workbook.setNumber(primary, 7);
      workbook.setFormula(extra, '=B2*2');
      expect(workbook.setCommentEntry(0, primary.row, primary.col, 'Alice', 'primary note')).toBe(
        true,
      );
      expect(workbook.setCommentEntry(0, extra.row, extra.col, 'Bob', 'extra note')).toBe(true);
      expect(
        workbook.addHyperlink(0, extra.row, extra.col, 'https://example.test', 'Example', 'tip'),
      ).toBe(true);
      mutators.setCellFormat(store, primary, { comment: 'primary note', commentAuthor: 'Alice' });
      mutators.setCellFormat(store, extra, {
        comment: 'extra note',
        commentAuthor: 'Bob',
        hyperlink: 'https://example.test',
        hyperlinkDisplay: 'Example',
        hyperlinkTooltip: 'tip',
        bold: true,
      });
      mutators.replaceCells(store, workbook.cells(0));
      mutators.setRange(store, {
        sheet: 0,
        r0: primary.row,
        c0: primary.col,
        r1: primary.row,
        c1: primary.col,
      });
      store.setState((state) => ({
        ...state,
        selection: {
          ...state.selection,
          extraRanges: [{ sheet: 0, r0: extra.row, c0: extra.col, r1: extra.row, c1: extra.col }],
        },
      }));

      executeRibbonClearAction({ store, workbook, history, action: 'all' });
      expect(workbook.getValue(primary)).toEqual({ kind: 'blank' });
      expect(workbook.getValue(extra)).toEqual({ kind: 'blank' });
      expect(workbook.cellFormula(extra)).toBeNull();
      expect(workbook.getComment(0, primary.row, primary.col)).toBeNull();
      expect(workbook.getComment(0, extra.row, extra.col)).toBeNull();
      expect(workbook.getHyperlinks(0)).not.toContainEqual(
        expect.objectContaining({ row: extra.row, col: extra.col }),
      );

      expect(history.undo()).toBe(true);
      expect(workbook.getValue(primary)).toEqual({ kind: 'number', value: 7 });
      expect(workbook.getValue(extra)).toEqual({ kind: 'number', value: 14 });
      expect(workbook.cellFormula(extra)).toBe('=B2*2');
      expect(workbook.getComment(0, primary.row, primary.col)).toEqual({
        author: 'Alice',
        text: 'primary note',
      });
      expect(workbook.getComment(0, extra.row, extra.col)).toEqual({
        author: 'Bob',
        text: 'extra note',
      });
      expect(workbook.getHyperlinks(0)).toContainEqual({
        row: extra.row,
        col: extra.col,
        target: 'https://example.test',
        display: 'Example',
        tooltip: 'tip',
      });
      expect(commentAt(store.getState(), primary)).toBe('primary note');

      expect(history.redo()).toBe(true);
      expect(workbook.getValue(primary)).toEqual({ kind: 'blank' });
      expect(workbook.getComment(0, primary.row, primary.col)).toBeNull();
      expect(workbook.getHyperlinks(0)).not.toContainEqual(
        expect.objectContaining({ row: extra.row, col: extra.col }),
      );
    } finally {
      workbook.dispose();
    }
  });
});
