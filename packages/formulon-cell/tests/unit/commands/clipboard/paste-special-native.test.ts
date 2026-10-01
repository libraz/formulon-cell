// @vitest-environment node

import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { pasteSpecial } from '../../../../src/commands/clipboard/paste-special.js';
import { captureSnapshot } from '../../../../src/commands/clipboard/snapshot.js';
import { History } from '../../../../src/commands/history.js';
import { addrKey, WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';

const source = { sheet: 0, row: 0, col: 0 };
const destination = { sheet: 0, row: 3, col: 3 };

const setActive = (store: SpreadsheetStore, addr: typeof destination): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      active: addr,
      anchor: addr,
      range: { sheet: addr.sheet, r0: addr.row, c0: addr.col, r1: addr.row, c1: addr.col },
    },
  }));
};

const mirrorNumber = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  addr: typeof source,
  value: number,
): void => {
  wb.setNumber(addr, value);
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    cells.set(addrKey(addr), { value: { kind: 'number', value }, formula: null });
    return { ...s, data: { ...s.data, cells } };
  });
};

const requireSnapshot = <T>(snapshot: T | null): T => {
  if (snapshot === null) throw new Error('expected clipboard snapshot');
  return snapshot;
};

describe('pasteSpecial native comments', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await WorkbookHandle.createDefault();
    expect(wb.isStub).toBe(false);
    expect(wb.capabilities.comments).toBe(true);
    expect(wb.capabilities.commentsEnumerable).toBe(true);
  });

  afterEach(() => {
    wb.dispose();
  });

  it('copies All comments to the native engine and preserves them through XLSX round-trip', async () => {
    mirrorNumber(store, wb, source, 7);
    mirrorNumber(store, wb, destination, 11);
    mutators.setCellFormat(store, source, {
      bold: true,
      comment: 'source note',
      commentAuthor: 'Alice',
    });
    mutators.setCellFormat(store, destination, {
      italic: true,
      comment: 'destination note',
      commentAuthor: 'Bob',
    });
    expect(wb.setCommentEntry(0, source.row, source.col, 'Alice', 'source note')).toBe(true);
    expect(wb.setCommentEntry(0, destination.row, destination.col, 'Bob', 'destination note')).toBe(
      true,
    );

    const snap = captureSnapshot(store.getState(), {
      sheet: 0,
      r0: source.row,
      c0: source.col,
      r1: source.row,
      c1: source.col,
    });
    const snapshot = requireSnapshot(snap);
    setActive(store, destination);

    pasteSpecial(store.getState(), store, wb, snapshot, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });

    expect(wb.getComment(0, destination.row, destination.col)).toEqual({
      author: 'Alice',
      text: 'source note',
    });
    expect(wb.getComment(0, source.row, source.col)).toEqual({
      author: 'Alice',
      text: 'source note',
    });
    expect(wb.getComments(0)).toContainEqual({
      row: destination.row,
      col: destination.col,
      author: 'Alice',
      text: 'source note',
    });

    const reloaded = await WorkbookHandle.loadBytes(wb.save());
    try {
      expect(reloaded.getComment(0, source.row, source.col)).toEqual({
        author: 'Alice',
        text: 'source note',
      });
      expect(reloaded.getComment(0, destination.row, destination.col)).toEqual({
        author: 'Alice',
        text: 'source note',
      });
    } finally {
      reloaded.dispose();
    }
  });

  it('Formats preserves the destination note in the native engine', () => {
    mirrorNumber(store, wb, source, 7);
    mutators.setCellFormat(store, source, {
      bold: true,
      comment: 'source note',
      commentAuthor: 'Alice',
    });
    mutators.setCellFormat(store, destination, {
      italic: true,
      comment: 'destination note',
      commentAuthor: 'Bob',
    });
    expect(wb.setCommentEntry(0, source.row, source.col, 'Alice', 'source note')).toBe(true);
    expect(wb.setCommentEntry(0, destination.row, destination.col, 'Bob', 'destination note')).toBe(
      true,
    );

    const snap = captureSnapshot(store.getState(), {
      sheet: 0,
      r0: source.row,
      c0: source.col,
      r1: source.row,
      c1: source.col,
    });
    const snapshot = requireSnapshot(snap);
    setActive(store, destination);

    pasteSpecial(store.getState(), store, wb, snapshot, {
      what: 'formats',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });

    expect(wb.getComment(0, destination.row, destination.col)).toEqual({
      author: 'Bob',
      text: 'destination note',
    });
    expect(store.getState().format.formats.get(addrKey(destination))).toEqual({
      bold: true,
      comment: 'destination note',
      commentAuthor: 'Bob',
    });
  });

  it('cuts comments and restores native notes on undo and redo', () => {
    mirrorNumber(store, wb, source, 7);
    mirrorNumber(store, wb, destination, 11);
    mutators.setCellFormat(store, source, {
      comment: 'source note',
      commentAuthor: 'Alice',
    });
    mutators.setCellFormat(store, destination, {
      comment: 'destination note',
      commentAuthor: 'Bob',
    });
    expect(wb.setCommentEntry(0, source.row, source.col, 'Alice', 'source note')).toBe(true);
    expect(wb.setCommentEntry(0, destination.row, destination.col, 'Bob', 'destination note')).toBe(
      true,
    );

    const snap = captureSnapshot(
      store.getState(),
      {
        sheet: 0,
        r0: source.row,
        c0: source.col,
        r1: source.row,
        c1: source.col,
      },
      'cut',
    );
    const snapshot = requireSnapshot(snap);
    setActive(store, destination);
    const history = new History();
    wb.attachHistory(history);
    history.begin();
    pasteSpecial(
      store.getState(),
      store,
      wb,
      snapshot,
      { what: 'all', operation: 'none', skipBlanks: false, transpose: false },
      history,
    );
    history.end();

    expect(wb.getComment(0, source.row, source.col)).toBeNull();
    expect(wb.getComment(0, destination.row, destination.col)).toEqual({
      author: 'Alice',
      text: 'source note',
    });
    expect(wb.getComments(0)).toEqual([
      { row: destination.row, col: destination.col, author: 'Alice', text: 'source note' },
    ]);
    expect(history.undo()).toBe(true);
    expect(wb.getComment(0, source.row, source.col)).toEqual({
      author: 'Alice',
      text: 'source note',
    });
    expect(wb.getComment(0, destination.row, destination.col)).toEqual({
      author: 'Bob',
      text: 'destination note',
    });
    expect(wb.getComments(0)).toHaveLength(2);
    expect(wb.getComments(0)).toEqual(
      expect.arrayContaining([
        { row: source.row, col: source.col, author: 'Alice', text: 'source note' },
        { row: destination.row, col: destination.col, author: 'Bob', text: 'destination note' },
      ]),
    );
    expect(history.redo()).toBe(true);
    expect(wb.getComment(0, source.row, source.col)).toBeNull();
    expect(wb.getComment(0, destination.row, destination.col)).toEqual({
      author: 'Alice',
      text: 'source note',
    });
    expect(wb.getComments(0)).toEqual([
      { row: destination.row, col: destination.col, author: 'Alice', text: 'source note' },
    ]);
  });
});
