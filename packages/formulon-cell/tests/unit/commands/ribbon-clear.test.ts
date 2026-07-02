import { describe, expect, it, vi } from 'vitest';

import { commentAt, setComment } from '../../../src/commands/comment.js';
import { History } from '../../../src/commands/history.js';
import { setCellLocked, setProtectedSheet } from '../../../src/commands/protection.js';
import { executeRibbonClearAction } from '../../../src/commands/ribbon-clear.js';
import type { Addr, CellValue } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

type CellEntry = { addr: Addr; value: CellValue; formula: string | null };

const key = (addr: Addr): string => `${addr.sheet}:${addr.row}:${addr.col}`;

const makeWorkbook = (
  entries: CellEntry[] = [],
): WorkbookHandle & {
  setBlank: ReturnType<typeof vi.fn>;
  setCommentEntry: ReturnType<typeof vi.fn>;
} => {
  const cells = new Map(entries.map((entry) => [key(entry.addr), entry]));
  const wb = {
    capabilities: { comments: true },
    setBlank: vi.fn((addr: Addr) => {
      cells.delete(key(addr));
    }),
    setCommentEntry: vi.fn(),
    *cells(sheet: number) {
      for (const entry of cells.values()) {
        if (entry.addr.sheet === sheet) yield entry;
      }
    },
  };
  return wb as unknown as WorkbookHandle & {
    setBlank: ReturnType<typeof vi.fn>;
    setCommentEntry: ReturnType<typeof vi.fn>;
  };
};

describe('executeRibbonClearAction', () => {
  it('clears contents only for writable cells and synchronizes store cells from the workbook', () => {
    const a1 = { sheet: 0, row: 0, col: 0 };
    const b1 = { sheet: 0, row: 0, col: 1 };
    const store = createSpreadsheetStore();
    const workbook = makeWorkbook([
      { addr: a1, value: { kind: 'number', value: 10 }, formula: null },
      { addr: b1, value: { kind: 'number', value: 20 }, formula: null },
    ]);
    mutators.replaceCells(store, workbook.cells(0));
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    setProtectedSheet(store, 0, true);
    setCellLocked(store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }, false);

    executeRibbonClearAction({
      store,
      workbook,
      history: new History(),
      action: 'contents',
    });

    expect(workbook.setBlank).toHaveBeenCalledTimes(1);
    expect(workbook.setBlank).toHaveBeenCalledWith(b1);
    expect(store.getState().data.cells.get(key(a1))?.value).toEqual({ kind: 'number', value: 10 });
    expect(store.getState().data.cells.has(key(b1))).toBe(false);
  });

  it('clears whole-column contents by visiting only materialized workbook cells', () => {
    const inColumn = { sheet: 0, row: 4, col: 2 };
    const outside = { sheet: 0, row: 4, col: 3 };
    const store = createSpreadsheetStore();
    const workbook = makeWorkbook([
      { addr: inColumn, value: { kind: 'text', value: 'clear' }, formula: null },
      { addr: outside, value: { kind: 'text', value: 'keep' }, formula: null },
    ]);
    mutators.replaceCells(store, workbook.cells(0));
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 2, r1: 1048575, c1: 2 });

    executeRibbonClearAction({
      store,
      workbook,
      history: new History(),
      action: 'contents',
    });

    expect(workbook.setBlank).toHaveBeenCalledTimes(1);
    expect(workbook.setBlank).toHaveBeenCalledWith(inColumn);
    expect(store.getState().data.cells.has(key(inColumn))).toBe(false);
    expect(store.getState().data.cells.get(key(outside))?.value).toEqual({
      kind: 'text',
      value: 'keep',
    });
  });

  it('clears comments in the selected range without touching comments outside it and supports undo', () => {
    const store = createSpreadsheetStore();
    const workbook = makeWorkbook();
    const history = new History();
    const inside = { sheet: 0, row: 0, col: 0 };
    const outside = { sheet: 0, row: 1, col: 0 };
    setComment(store, inside, 'inside', workbook);
    setComment(store, outside, 'outside', workbook);
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    workbook.setCommentEntry.mockClear();

    executeRibbonClearAction({
      store,
      workbook,
      history,
      action: 'comments',
    });

    expect(commentAt(store.getState(), inside)).toBeNull();
    expect(commentAt(store.getState(), outside)).toBe('outside');
    expect(workbook.setCommentEntry).toHaveBeenCalledWith(0, 0, 0, '', '');

    expect(history.undo()).toBe(true);
    expect(commentAt(store.getState(), inside)).toBe('inside');
    expect(commentAt(store.getState(), outside)).toBe('outside');
  });

  it('clears whole-column comments by visiting only formatted cells with comments', () => {
    const store = createSpreadsheetStore();
    const workbook = makeWorkbook();
    const inside = { sheet: 0, row: 4, col: 2 };
    const outside = { sheet: 0, row: 4, col: 3 };
    setComment(store, inside, 'inside', workbook);
    setComment(store, outside, 'outside', workbook);
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 2, r1: 1048575, c1: 2 });
    workbook.setCommentEntry.mockClear();

    executeRibbonClearAction({
      store,
      workbook,
      history: new History(),
      action: 'comments',
    });

    expect(commentAt(store.getState(), inside)).toBeNull();
    expect(commentAt(store.getState(), outside)).toBe('outside');
    expect(workbook.setCommentEntry).toHaveBeenCalledTimes(1);
    expect(workbook.setCommentEntry).toHaveBeenCalledWith(0, 4, 2, '', '');
  });

  it('clears only visual format properties for formats action and leaves comments intact', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const addr = { sheet: 0, row: 0, col: 0 };
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    mutators.setCellFormat(store, addr, {
      bold: true,
      fill: '#ffff00',
      comment: 'keep',
      validation: { kind: 'whole', op: 'between', a: 1, b: 10 },
    });

    executeRibbonClearAction({
      store,
      workbook: makeWorkbook(),
      history,
      action: 'formats',
    });

    expect(store.getState().format.formats.get(key(addr))).toEqual({
      comment: 'keep',
      validation: { kind: 'whole', op: 'between', a: 1, b: 10 },
    });

    expect(history.undo()).toBe(true);
    expect(store.getState().format.formats.get(key(addr))).toMatchObject({
      bold: true,
      fill: '#ffff00',
      comment: 'keep',
    });
  });
});
