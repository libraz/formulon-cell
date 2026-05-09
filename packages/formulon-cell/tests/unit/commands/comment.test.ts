import { describe, expect, it } from 'vitest';
import {
  clearComment,
  commentAt,
  listComments,
  setComment,
} from '../../../src/commands/comment.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('comment commands', () => {
  it('returns null for a cell with no comment', () => {
    const store = createSpreadsheetStore();
    expect(commentAt(store.getState(), { sheet: 0, row: 0, col: 0 })).toBeNull();
  });

  it('setComment writes the text into format.comment', () => {
    const store = createSpreadsheetStore();
    setComment(store, { sheet: 0, row: 1, col: 2 }, 'hello');
    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 2 }));
    expect(fmt?.comment).toBe('hello');
  });

  it('commentAt round-trips a setComment call', () => {
    const store = createSpreadsheetStore();
    setComment(store, { sheet: 0, row: 0, col: 0 }, 'note');
    expect(commentAt(store.getState(), { sheet: 0, row: 0, col: 0 })).toBe('note');
  });

  it('lists comments on a sheet in row-major order', () => {
    const store = createSpreadsheetStore();

    setComment(store, { sheet: 0, row: 2, col: 1 }, 'C');
    setComment(store, { sheet: 0, row: 0, col: 0 }, 'A');
    setComment(store, { sheet: 0, row: 0, col: 2 }, 'B');
    setComment(store, { sheet: 1, row: 0, col: 0 }, 'hidden');

    expect(listComments(store.getState(), 0)).toEqual([
      { addr: { sheet: 0, row: 0, col: 0 }, text: 'A' },
      { addr: { sheet: 0, row: 0, col: 2 }, text: 'B' },
      { addr: { sheet: 0, row: 2, col: 1 }, text: 'C' },
    ]);
  });

  it('setComment with empty string clears the comment', () => {
    const store = createSpreadsheetStore();
    setComment(store, { sheet: 0, row: 0, col: 0 }, 'note');
    setComment(store, { sheet: 0, row: 0, col: 0 }, '');
    expect(commentAt(store.getState(), { sheet: 0, row: 0, col: 0 })).toBeNull();
  });

  it('clearComment removes the comment field', () => {
    const store = createSpreadsheetStore();
    setComment(store, { sheet: 0, row: 0, col: 0 }, 'note');
    clearComment(store, { sheet: 0, row: 0, col: 0 });
    expect(commentAt(store.getState(), { sheet: 0, row: 0, col: 0 })).toBeNull();
  });

  it('mirrors setComment to the engine when wb supports comments', () => {
    const store = createSpreadsheetStore();
    const calls: { row: number; col: number; author: string; text: string }[] = [];
    const wb = {
      capabilities: { comments: true },
      setCommentEntry: (_sheet: number, row: number, col: number, author: string, text: string) => {
        calls.push({ row, col, author, text });
        return true;
      },
    } as unknown as WorkbookHandle;
    setComment(store, { sheet: 0, row: 1, col: 2 }, 'note', wb);
    expect(calls).toEqual([{ row: 1, col: 2, author: '', text: 'note' }]);
  });

  it('mirrors setComment with empty text to the engine (engine treats empty as remove)', () => {
    const store = createSpreadsheetStore();
    const calls: { text: string }[] = [];
    const wb = {
      capabilities: { comments: true },
      setCommentEntry: (
        _sheet: number,
        _row: number,
        _col: number,
        _author: string,
        text: string,
      ) => {
        calls.push({ text });
        return true;
      },
    } as unknown as WorkbookHandle;
    setComment(store, { sheet: 0, row: 0, col: 0 }, '', wb);
    expect(calls).toEqual([{ text: '' }]);
  });

  it('mirrors clearComment to the engine when wb supports comments', () => {
    const store = createSpreadsheetStore();
    const calls: { text: string }[] = [];
    const wb = {
      capabilities: { comments: true },
      setCommentEntry: (
        _sheet: number,
        _row: number,
        _col: number,
        _author: string,
        text: string,
      ) => {
        calls.push({ text });
        return true;
      },
    } as unknown as WorkbookHandle;
    setComment(store, { sheet: 0, row: 0, col: 0 }, 'note', wb);
    clearComment(store, { sheet: 0, row: 0, col: 0 }, wb);
    expect(calls).toEqual([{ text: 'note' }, { text: '' }]);
  });

  it('skips engine call when capability flag is off', () => {
    const store = createSpreadsheetStore();
    const calls: unknown[] = [];
    const wb = {
      capabilities: { comments: false },
      setCommentEntry: () => {
        calls.push(1);
        return true;
      },
    } as unknown as WorkbookHandle;
    setComment(store, { sheet: 0, row: 0, col: 0 }, 'note', wb);
    clearComment(store, { sheet: 0, row: 0, col: 0 }, wb);
    expect(calls).toEqual([]);
  });

  it('does not affect other format fields when setting a comment', () => {
    const store = createSpreadsheetStore();
    // Seed an existing format that we shouldn't clobber.
    store.setState((s) => {
      const formats = new Map(s.format.formats);
      formats.set(addrKey({ sheet: 0, row: 0, col: 0 }), { bold: true, color: '#ff0000' });
      return { ...s, format: { formats } };
    });
    setComment(store, { sheet: 0, row: 0, col: 0 }, 'note');
    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.bold).toBe(true);
    expect(fmt?.color).toBe('#ff0000');
    expect(fmt?.comment).toBe('note');
  });
});
