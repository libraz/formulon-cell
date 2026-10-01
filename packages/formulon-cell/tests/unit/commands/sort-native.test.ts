// @vitest-environment node

import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { sortRange } from '../../../src/commands/sort.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const source = { sheet: 0, row: 1, col: 7 };
const destination = { sheet: 0, row: 0, col: 7 };

describe('sortRange native parity', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await WorkbookHandle.createDefault();
    expect(wb.isStub).toBe(false);
    expect(wb.capabilities.comments).toBe(true);
  });

  afterEach(() => {
    wb.dispose();
  });

  it('moves relative formulas, formats, and comments as one undoable action', async () => {
    const gValues = [3, 1, 2];
    const jValues = [100, 200, 300];
    for (let row = 0; row < gValues.length; row += 1) {
      wb.setNumber({ sheet: 0, row, col: 6 }, gValues[row] ?? 0);
      wb.setNumber({ sheet: 0, row, col: 9 }, jValues[row] ?? 0);
      wb.setFormula({ sheet: 0, row, col: 7 }, `=G${row + 1}+$J$1+J${row + 1}`);
    }
    wb.setFormula({ sheet: 0, row: 0, col: 10 }, '=G1');
    wb.recalc();
    mutators.replaceCells(store, wb.cells(0));
    mutators.setCellFormat(store, source, { bold: true });
    expect(wb.setCommentEntry(0, source.row, source.col, 'Alice', 'moved note')).toBe(true);
    expect(store.getState().format.formats.get(addrKey(source))?.comment).toBeUndefined();

    const history = new History();
    wb.attachHistory(history);
    history.begin();
    const ok = sortRange(
      store.getState(),
      store,
      wb,
      { sheet: 0, r0: 0, c0: 6, r1: 2, c1: 7 },
      { byCol: 6, direction: 'asc' },
      history,
    );
    history.end();

    expect(ok).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 0, col: 6 })).toEqual({ kind: 'number', value: 1 });
    expect(wb.getValue({ sheet: 0, row: 1, col: 6 })).toEqual({ kind: 'number', value: 2 });
    expect(wb.getValue({ sheet: 0, row: 2, col: 6 })).toEqual({ kind: 'number', value: 3 });
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 7 })).toBe('=G1+$J$1+J1');
    expect(wb.cellFormula({ sheet: 0, row: 1, col: 7 })).toBe('=G2+$J$1+J2');
    expect(wb.cellFormula({ sheet: 0, row: 2, col: 7 })).toBe('=G3+$J$1+J3');
    expect(wb.getValue({ sheet: 0, row: 0, col: 7 })).toEqual({ kind: 'number', value: 201 });
    expect(wb.getValue({ sheet: 0, row: 1, col: 7 })).toEqual({ kind: 'number', value: 302 });
    expect(wb.getValue({ sheet: 0, row: 2, col: 7 })).toEqual({ kind: 'number', value: 403 });
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 10 })).toBe('=G1');
    expect(store.getState().format.formats.get(addrKey(destination))).toMatchObject({
      bold: true,
      comment: 'moved note',
      commentAuthor: 'Alice',
    });
    if (wb.capabilities.comments) {
      expect(wb.getComment(0, destination.row, destination.col)).toEqual({
        author: 'Alice',
        text: 'moved note',
      });
    }

    const reloaded = await WorkbookHandle.loadBytes(wb.save());
    try {
      expect(reloaded.cellFormula({ sheet: 0, row: 0, col: 7 })).toBe('=G1+$J$1+J1');
      if (reloaded.capabilities.comments) {
        expect(reloaded.getComment(0, destination.row, destination.col)).toEqual({
          author: 'Alice',
          text: 'moved note',
        });
      }
    } finally {
      reloaded.dispose();
    }

    expect(history.undo()).toBe(true);
    expect(history.canUndo()).toBe(false);
    expect(wb.getValue({ sheet: 0, row: 0, col: 6 })).toEqual({ kind: 'number', value: 3 });
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 7 })).toBe('=G1+$J$1+J1');
    expect(store.getState().format.formats.get(addrKey(source))).toMatchObject({
      bold: true,
      comment: 'moved note',
      commentAuthor: 'Alice',
    });
    if (wb.capabilities.comments) {
      expect(wb.getComment(0, source.row, source.col)).toEqual({
        author: 'Alice',
        text: 'moved note',
      });
    }

    expect(history.redo()).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 0, col: 6 })).toEqual({ kind: 'number', value: 1 });
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 7 })).toBe('=G1+$J$1+J1');
    if (wb.capabilities.comments) {
      expect(wb.getComment(0, destination.row, destination.col)).toEqual({
        author: 'Alice',
        text: 'moved note',
      });
    }
  });
});
