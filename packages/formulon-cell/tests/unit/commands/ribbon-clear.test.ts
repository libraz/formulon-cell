import { describe, expect, it, vi } from 'vitest';

import { commentAt, setComment } from '../../../src/commands/comment.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { fixedFormPolicy } from '../../../src/commands/interaction-policy.js';
import { setCellLocked, setProtectedSheet } from '../../../src/commands/protection.js';
import { executeRibbonClearAction } from '../../../src/commands/ribbon-clear.js';
import { flushFormatToEngine } from '../../../src/engine/cell-format-sync.js';
import type { Addr, CellValue } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
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
    setCommentEntry: vi.fn(() => true),
    getComment: vi.fn(() => null),
    *cells(sheet: number) {
      for (const entry of cells.values()) {
        if (entry.addr.sheet === sheet) yield entry;
      }
    },
    *physicalCells(sheet: number) {
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
  it('clears contents only for writable cells and synchronizes store cells from the workbook', async () => {
    const a1 = { sheet: 0, row: 0, col: 0 };
    const b1 = { sheet: 0, row: 0, col: 1 };
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    workbook.setNumber(a1, 10);
    workbook.setNumber(b1, 20);
    mutators.replaceCells(store, workbook.cells(0));
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    setProtectedSheet(store, 0, true);
    setCellLocked(store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }, false);

    try {
      executeRibbonClearAction({
        store,
        workbook,
        history: new History(),
        action: 'contents',
      });

      expect(workbook.getValue(a1)).toEqual({ kind: 'number', value: 10 });
      expect(workbook.getValue(b1)).toEqual({ kind: 'blank' });
      expect(store.getState().data.cells.get(key(a1))?.value).toEqual({
        kind: 'number',
        value: 10,
      });
      expect(store.getState().data.cells.has(key(b1))).toBe(false);
    } finally {
      workbook.dispose();
    }
  });

  it('clears whole-column contents by visiting only materialized workbook cells', async () => {
    const inColumn = { sheet: 0, row: 4, col: 2 };
    const outside = { sheet: 0, row: 4, col: 3 };
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    workbook.setText(inColumn, 'clear');
    workbook.setText(outside, 'keep');
    mutators.replaceCells(store, workbook.cells(0));
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 2, r1: 1048575, c1: 2 });

    try {
      executeRibbonClearAction({
        store,
        workbook,
        history: new History(),
        action: 'contents',
      });

      expect(workbook.getValue(inColumn)).toEqual({ kind: 'blank' });
      expect(workbook.getValue(outside)).toEqual({ kind: 'text', value: 'keep' });
      expect(store.getState().data.cells.has(key(inColumn))).toBe(false);
      expect(store.getState().data.cells.get(key(outside))?.value).toEqual({
        kind: 'text',
        value: 'keep',
      });
    } finally {
      workbook.dispose();
    }
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

  it('repeats a Clear Formats onto the selection F4 was pressed on', () => {
    const store = createSpreadsheetStore();
    const workbook = makeWorkbook();
    const history = new History();
    mutators.setRangeFormat(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, { bold: true });
    mutators.setRangeFormat(store, { sheet: 0, r0: 5, c0: 0, r1: 5, c1: 0 }, { italic: true });
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });

    executeRibbonClearAction({ store, workbook, history, action: 'formats' });
    mutators.setRange(store, { sheet: 0, r0: 5, c0: 0, r1: 5, c1: 0 });
    expect(history.repeatLast()).toBe(true);

    expect(store.getState().format.formats.get('0:0:0')?.bold).toBeUndefined();
    expect(store.getState().format.formats.get('0:5:0')?.italic).toBeUndefined();
  });

  it('clears comments, hyperlinks, and conditional rules across the full union', () => {
    const store = createSpreadsheetStore();
    const workbook = makeWorkbook();
    const history = new History();
    const primary = { sheet: 0, row: 0, col: 0 };
    const extra = { sheet: 0, row: 4, col: 4 };
    const hole = { sheet: 0, row: 2, col: 2 };
    const outside = { sheet: 0, row: 8, col: 8 };
    setComment(store, primary, 'primary', workbook);
    setComment(store, extra, 'extra', workbook);
    setComment(store, hole, 'hole', workbook);
    setComment(store, outside, 'outside', workbook);
    mutators.setCellFormat(store, primary, {
      hyperlink: 'https://primary.example',
      hyperlinkDisplay: 'Primary',
      hyperlinkTooltip: 'primary tip',
    });
    mutators.setCellFormat(store, extra, {
      hyperlink: 'https://extra.example',
      hyperlinkDisplay: 'Extra',
      hyperlinkTooltip: 'extra tip',
    });
    mutators.setCellFormat(store, hole, { hyperlink: 'https://hole.example' });
    mutators.setCellFormat(store, outside, { hyperlink: 'https://outside.example' });
    mutators.setRange(store, {
      sheet: 0,
      r0: primary.row,
      c0: primary.col,
      r1: primary.row,
      c1: primary.col,
    });
    store.setState((s) => ({
      ...s,
      selection: {
        ...s.selection,
        extraRanges: [
          {
            sheet: 0,
            r0: extra.row,
            c0: extra.col,
            r1: extra.row,
            c1: extra.col,
          },
        ],
      },
      conditional: {
        rules: [
          {
            kind: 'duplicates',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 3 },
            apply: { fill: '#f00' },
          },
          {
            kind: 'unique',
            range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
            apply: { fill: '#0f0' },
          },
          {
            kind: 'top-bottom',
            range: { sheet: 0, r0: 4, c0: 4, r1: 4, c1: 4 },
            mode: 'top',
            n: 10,
            percent: false,
            apply: { fill: '#ff0' },
          },
          {
            kind: 'average',
            range: { sheet: 0, r0: 8, c0: 8, r1: 8, c1: 8 },
            mode: 'above',
            apply: { fill: '#00f' },
          },
        ],
      },
    }));

    executeRibbonClearAction({ store, workbook, history, action: 'comments' });
    expect(commentAt(store.getState(), primary)).toBeNull();
    expect(commentAt(store.getState(), extra)).toBeNull();
    expect(commentAt(store.getState(), hole)).toBe('hole');
    expect(commentAt(store.getState(), outside)).toBe('outside');

    executeRibbonClearAction({ store, workbook, history, action: 'hyperlinks' });
    expect(store.getState().format.formats.get(key(primary))?.hyperlink).toBeUndefined();
    expect(store.getState().format.formats.get(key(extra))?.hyperlink).toBeUndefined();
    expect(store.getState().format.formats.get(key(hole))?.hyperlink).toBe('https://hole.example');
    expect(store.getState().format.formats.get(key(outside))?.hyperlink).toBe(
      'https://outside.example',
    );

    executeRibbonClearAction({ store, workbook, history, action: 'conditional' });
    expect(store.getState().conditional.rules).toHaveLength(2);
    expect(store.getState().conditional.rules.map((rule) => rule.kind)).toEqual([
      'unique',
      'average',
    ]);
  });

  it('clears All as one unrestricted history step and restores every selected category', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const history = new History();
    const primary = { sheet: 0, row: 0, col: 0 };
    const extra = { sheet: 0, row: 4, col: 4 };
    workbook.setNumber(primary, 7);
    workbook.setFormula(extra, '=1+2');
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    store.setState((s) => ({
      ...s,
      selection: {
        ...s.selection,
        extraRanges: [{ sheet: 0, r0: 4, c0: 4, r1: 4, c1: 4 }],
      },
    }));
    setComment(store, primary, 'primary');
    setComment(store, extra, 'extra');
    mutators.setCellFormat(store, primary, {
      hyperlink: 'https://primary.example',
      hyperlinkDisplay: 'Primary',
      hyperlinkTooltip: 'tip',
      validation: { kind: 'whole', op: 'between', a: 1, b: 9 },
      bold: true,
    });
    mutators.setCellFormat(store, extra, {
      hyperlink: 'https://extra.example',
      validation: { kind: 'whole', op: 'between', a: 1, b: 9 },
      italic: true,
    });
    mutators.setPendingFormat(store, { addr: extra, format: { fill: '#ff0' } });
    store.setState((s) => ({
      ...s,
      conditional: {
        rules: [
          {
            kind: 'duplicates',
            range: { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 4 },
            apply: { fill: '#f00' },
          },
        ],
      },
    }));
    mutators.replaceCells(store, workbook.cells(0));

    try {
      executeRibbonClearAction({ store, workbook, history, action: 'all' });
      expect(workbook.getValue(primary)).toEqual({ kind: 'blank' });
      expect(workbook.getValue(extra)).toEqual({ kind: 'blank' });
      expect(commentAt(store.getState(), primary)).toBeNull();
      expect(commentAt(store.getState(), extra)).toBeNull();
      expect(store.getState().format.formats.get(key(primary))).toBeUndefined();
      expect(store.getState().format.formats.get(key(extra))).toBeUndefined();
      expect(store.getState().ui.pendingFormat).toBeNull();
      expect(store.getState().conditional.rules).toHaveLength(0);
      expect(history.undo()).toBe(true);
      expect(workbook.getValue(primary)).toEqual({ kind: 'number', value: 7 });
      expect(workbook.cellFormula(extra)).toBe('=1+2');
      expect(commentAt(store.getState(), primary)).toBe('primary');
      expect(commentAt(store.getState(), extra)).toBe('extra');
      expect(store.getState().format.formats.get(key(primary))).toMatchObject({
        hyperlink: 'https://primary.example',
        validation: { kind: 'whole' },
        bold: true,
      });
      expect(store.getState().conditional.rules).toHaveLength(1);
      expect(history.undo()).toBe(false);
      expect(history.redo()).toBe(true);
      expect(workbook.getValue(primary)).toEqual({ kind: 'blank' });
      expect(commentAt(store.getState(), primary)).toBeNull();
      expect(commentAt(store.getState(), extra)).toBeNull();
    } finally {
      workbook.dispose();
    }
  });

  it('clears All contents through a registered controller even without a policy', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const history = new History();
    const addr = { sheet: 0, row: 0, col: 0 };
    const locked = { sheet: 0, row: 0, col: 1 };
    workbook.setNumber(addr, 7);
    workbook.setNumber(locked, 8);
    mutators.replaceCells(store, workbook.cells(0));
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    mutators.setCellFormat(store, addr, { bold: true });
    setCellLocked(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, false);
    setProtectedSheet(store, 0, true);
    const onChanged = vi.fn();
    const controller = new InteractionController({
      store,
      getWb: () => workbook,
      history,
      onChanged,
    });
    const unregister = registerInteractionController(store, controller);
    try {
      executeRibbonClearAction({ store, workbook, history, action: 'all' });
      expect(onChanged).toHaveBeenCalledWith(expect.objectContaining({ status: 'applied' }));
      expect(workbook.getValue(addr)).toEqual({ kind: 'blank' });
      expect(workbook.getValue(locked)).toEqual({ kind: 'number', value: 8 });
      expect(store.getState().format.formats.get(key(addr))?.bold).toBeUndefined();
      expect(history.undo()).toBe(true);
      expect(workbook.getValue(addr)).toEqual({ kind: 'number', value: 7 });
      expect(store.getState().format.formats.get(key(addr))?.bold).toBe(true);
      expect(history.undo()).toBe(false);
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });

  it('rejects a malformed foreign-sheet extra range before any clear', () => {
    const store = createSpreadsheetStore();
    const workbook = makeWorkbook([
      { addr: { sheet: 0, row: 0, col: 0 }, value: { kind: 'number', value: 7 }, formula: null },
    ]);
    const history = new History();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    store.setState((s) => ({
      ...s,
      selection: {
        ...s.selection,
        extraRanges: [{ sheet: 1, r0: 0, c0: 0, r1: 0, c1: 0 }],
      },
    }));

    executeRibbonClearAction({ store, workbook, history, action: 'all' });
    expect([...workbook.cells(0)]).toHaveLength(1);
    expect(history.canUndo()).toBe(false);
  });

  it('rolls back content and physical comments when a later comment write fails', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault();
    const history = new History();
    const primary = { sheet: 0, row: 1, col: 1 };
    const extra = { sheet: 0, row: 4, col: 4 };
    try {
      expect(workbook.isStub).toBe(false);
      workbook.setNumber(primary, 7);
      workbook.setText(extra, 'extra');
      expect(workbook.setCommentEntry(0, primary.row, primary.col, 'Alice', 'primary')).toBe(true);
      expect(workbook.setCommentEntry(0, extra.row, extra.col, 'Bob', 'extra')).toBe(true);
      mutators.setCellFormat(store, primary, { comment: 'primary', commentAuthor: 'Alice' });
      mutators.setCellFormat(store, extra, { comment: 'extra', commentAuthor: 'Bob' });
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
        ui: { ...state.ui, pendingFormat: { addr: extra, format: { fill: '#ff0' } } },
      }));
      const beforePending = structuredClone(store.getState().ui.pendingFormat);
      const beforePrimary = workbook.getComment(0, primary.row, primary.col);
      const beforeExtra = workbook.getComment(0, extra.row, extra.col);
      const originalSetCommentEntry = workbook.setCommentEntry.bind(workbook);
      let clears = 0;
      vi.spyOn(workbook, 'setCommentEntry').mockImplementation((sheet, row, col, author, text) => {
        if (text.length === 0 && ++clears === 2) return false;
        return originalSetCommentEntry(sheet, row, col, author, text);
      });

      expect(() => executeRibbonClearAction({ store, workbook, history, action: 'all' })).toThrow(
        'comment engine write failed',
      );
      expect(workbook.getValue(primary)).toEqual({ kind: 'number', value: 7 });
      expect(workbook.getValue(extra)).toEqual({ kind: 'text', value: 'extra' });
      expect(workbook.getComment(0, primary.row, primary.col)).toEqual(beforePrimary);
      expect(workbook.getComment(0, extra.row, extra.col)).toEqual(beforeExtra);
      expect(commentAt(store.getState(), primary)).toBe('primary');
      expect(commentAt(store.getState(), extra)).toBe('extra');
      expect(store.getState().ui.pendingFormat).toEqual(beforePending);
      expect(history.canUndo()).toBe(false);
    } finally {
      workbook.dispose();
    }
  });

  it('preflights a registered policy before Clear All and keeps restricted Undo fail-closed', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const history = new History();
    const addr = { sheet: 0, row: 0, col: 0 };
    workbook.setFormula(addr, '=1+2');
    mutators.replaceCells(store, workbook.cells(0));
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    mutators.setCellFormat(store, addr, { bold: true });
    const controller = new InteractionController({ store, getWb: () => workbook, history });
    const unregister = registerInteractionController(store, controller);
    try {
      const base = fixedFormPolicy([{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }]);
      controller.setPolicy({
        ...base,
        operations: { ...base.operations, clear: true, format: true },
      });
      executeRibbonClearAction({ store, workbook, history, action: 'all' });
      expect(workbook.getValue(addr)).toEqual({ kind: 'blank' });
      expect(store.getState().format.formats.get('0:0:0')).toBeUndefined();
      expect(history.canUndo()).toBe(true);
      expect(history.undo()).toBe(false);
      expect(workbook.getValue(addr)).toEqual({ kind: 'blank' });
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });

  it('uses clear-all for every Clear All preflight while standalone contents uses clear-contents', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const history = new History();
    const addr = { sheet: 0, row: 0, col: 0 };
    workbook.setFormula(addr, '=1+2');
    mutators.replaceCells(store, workbook.cells(0));
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    mutators.setCellFormat(store, addr, { bold: true });
    const controller = new InteractionController({ store, getWb: () => workbook, history });
    const unregister = registerInteractionController(store, controller);
    try {
      const base = fixedFormPolicy([{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }]);
      controller.setPolicy({
        ...base,
        operations: { ...base.operations, clear: true, format: true },
        restrict: ({ origin, commandId }) => origin === 'ribbon' && commandId === 'clear-all',
      });
      executeRibbonClearAction({ store, workbook, history, action: 'all' });
      expect(workbook.cellFormula(addr)).toBeNull();
      expect(store.getState().format.formats.get(key(addr))).toBeUndefined();

      workbook.setFormula(addr, '=1+2');
      mutators.replaceCells(store, workbook.cells(0));
      mutators.setCellFormat(store, addr, { bold: true });
      executeRibbonClearAction({ store, workbook, history, action: 'contents' });
      expect(workbook.cellFormula(addr)).toBe('=1+2');
      expect(store.getState().format.formats.get(key(addr))?.bold).toBe(true);
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });

  it('does not let clear-contents permission authorize composite Clear All', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const history = new History();
    const addr = { sheet: 0, row: 0, col: 0 };
    workbook.setFormula(addr, '=1+2');
    mutators.replaceCells(store, workbook.cells(0));
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    mutators.setCellFormat(store, addr, { bold: true });
    const controller = new InteractionController({ store, getWb: () => workbook, history });
    const unregister = registerInteractionController(store, controller);
    try {
      const base = fixedFormPolicy([{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }]);
      controller.setPolicy({
        ...base,
        operations: { ...base.operations, clear: true, format: true },
        restrict: ({ origin, commandId }) => origin === 'ribbon' && commandId === 'clear-contents',
      });
      executeRibbonClearAction({ store, workbook, history, action: 'all' });
      expect(workbook.cellFormula(addr)).toBe('=1+2');
      expect(store.getState().format.formats.get(key(addr))?.bold).toBe(true);
      expect(history.canUndo()).toBe(false);

      executeRibbonClearAction({ store, workbook, history, action: 'contents' });
      expect(workbook.cellFormula(addr)).toBeNull();
      expect(store.getState().format.formats.get(key(addr))?.bold).toBe(true);
      expect(history.canUndo()).toBe(true);
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });

  it('rejects Clear All atomically when registered clear policy denies the selection', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const history = new History();
    const addr = { sheet: 0, row: 0, col: 0 };
    workbook.setNumber(addr, 7);
    mutators.replaceCells(store, workbook.cells(0));
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    mutators.setCellFormat(store, addr, { bold: true });
    const controller = new InteractionController({ store, getWb: () => workbook, history });
    const unregister = registerInteractionController(store, controller);
    try {
      const base = fixedFormPolicy([{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }]);
      controller.setPolicy({
        ...base,
        operations: { ...base.operations, clear: false, format: true },
      });
      executeRibbonClearAction({ store, workbook, history, action: 'all' });
      expect(workbook.getValue(addr)).toEqual({ kind: 'number', value: 7 });
      expect(store.getState().format.formats.get('0:0:0')?.bold).toBe(true);
      expect(history.canUndo()).toBe(false);
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });

  it('fails closed before content when Clear All contains unsupported comment metadata', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const history = new History();
    const addr = { sheet: 0, row: 0, col: 0 };
    workbook.setNumber(addr, 7);
    mutators.replaceCells(store, workbook.cells(0));
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setComment(store, addr, 'protected metadata');
    const controller = new InteractionController({ store, getWb: () => workbook, history });
    const unregister = registerInteractionController(store, controller);
    try {
      const base = fixedFormPolicy([{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }]);
      controller.setPolicy({
        ...base,
        operations: { ...base.operations, clear: true, format: true, comment: true },
      });
      executeRibbonClearAction({ store, workbook, history, action: 'all' });
      expect(workbook.getValue(addr)).toEqual({ kind: 'number', value: 7 });
      expect(commentAt(store.getState(), addr)).toBe('protected metadata');
      expect(history.canUndo()).toBe(false);
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });

  it('rejects a foreign no-policy command beside a registered policy and ignores a shared-policy executor', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const history = new History();
    const addr = { sheet: 0, row: 0, col: 0 };
    workbook.setNumber(addr, 7);
    mutators.replaceCells(store, workbook.cells(0));
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const controller = new InteractionController({ store, getWb: () => workbook, history });
    const unregister = registerInteractionController(store, controller);
    try {
      const base = fixedFormPolicy([{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }]);
      const policy = { ...base, operations: { ...base.operations, clear: true } };
      controller.setPolicy(policy);
      const foreign = {
        execute: vi.fn() as unknown as InteractionController['execute'],
        policy: undefined,
      } as Pick<InteractionController, 'execute' | 'policy'>;
      executeRibbonClearAction({ store, workbook, history, action: 'all', commands: foreign });
      expect(workbook.getValue(addr)).toEqual({ kind: 'number', value: 7 });
      const shared = {
        execute: vi.fn() as unknown as InteractionController['execute'],
        policy,
      } as Pick<InteractionController, 'execute' | 'policy'>;
      executeRibbonClearAction({ store, workbook, history, action: 'all', commands: shared });
      expect(shared.execute).not.toHaveBeenCalled();
      expect(workbook.getValue(addr)).toEqual({ kind: 'blank' });
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });

  it('uses the ribbon clear-formats origin and command id for format policy preflight', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const history = new History();
    const addr = { sheet: 0, row: 0, col: 0 };
    mutators.setCellFormat(store, addr, { bold: true });
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const controller = new InteractionController({ store, getWb: () => workbook, history });
    const unregister = registerInteractionController(store, controller);
    try {
      const base = fixedFormPolicy([{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }]);
      controller.setPolicy({
        ...base,
        operations: { ...base.operations, format: true },
        restrict: ({ origin, commandId }) => origin === 'ribbon' && commandId === 'clear-formats',
      });
      executeRibbonClearAction({ store, workbook, history, action: 'formats' });
      expect(store.getState().format.formats.get(key(addr))).toBeUndefined();
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });

  it('clears sparse full-column content and a remote extra without materializing the column', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const history = new History();
    const inColumn = { sheet: 0, row: 100_001, col: 0 };
    const remote = { sheet: 0, row: 200_000, col: 4 };
    workbook.setText(inColumn, 'column');
    workbook.setText(remote, 'remote');
    mutators.setCellFormat(store, remote, { hyperlink: 'https://remote.example', bold: true });
    mutators.replaceCells(store, workbook.physicalCells(0));
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1_048_575, c1: 0 });
    store.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        extraRanges: [{ sheet: 0, r0: remote.row, c0: remote.col, r1: remote.row, c1: remote.col }],
      },
    }));
    try {
      executeRibbonClearAction({ store, workbook, history, action: 'all' });
      expect(workbook.getValue(inColumn)).toEqual({ kind: 'blank' });
      expect(workbook.getValue(remote)).toEqual({ kind: 'blank' });
      expect(store.getState().format.formats.has('0:200000:4')).toBe(false);
    } finally {
      workbook.dispose();
    }
  });

  it('treats partial merges as a no-op and clears a fully covered merge once', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const history = new History();
    const anchor = { sheet: 0, row: 0, col: 0 };
    const extra = { sheet: 0, row: 2, col: 2 };
    workbook.setText(anchor, 'merged');
    workbook.setText(extra, 'extra');
    mutators.replaceCells(store, workbook.physicalCells(0));
    mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    try {
      executeRibbonClearAction({ store, workbook, history, action: 'all' });
      expect(workbook.getValue(anchor)).toEqual({ kind: 'text', value: 'merged' });
      expect(workbook.getValue(extra)).toEqual({ kind: 'text', value: 'extra' });

      mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
      store.setState((state) => ({
        ...state,
        selection: {
          ...state.selection,
          extraRanges: [{ sheet: 0, r0: extra.row, c0: extra.col, r1: extra.row, c1: extra.col }],
        },
      }));
      executeRibbonClearAction({ store, workbook, history, action: 'all' });
      expect(workbook.getValue(anchor)).toEqual({ kind: 'blank' });
      expect(workbook.getValue(extra)).toEqual({ kind: 'blank' });
      expect(store.getState().merges.byAnchor.has('0:0:0')).toBe(true);
    } finally {
      workbook.dispose();
    }
  });

  it('skips protected cells and retains an indivisible conditional rule spanning a lock', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const history = new History();
    const writable = { sheet: 0, row: 0, col: 0 };
    const locked = { sheet: 0, row: 0, col: 1 };
    workbook.setNumber(writable, 1);
    workbook.setNumber(locked, 2);
    mutators.replaceCells(store, workbook.cells(0));
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    setComment(store, writable, 'writable');
    setComment(store, locked, 'locked');
    mutators.setCellFormat(store, writable, { bold: true });
    mutators.setCellFormat(store, locked, { italic: true });
    setProtectedSheet(store, 0, true);
    setCellLocked(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, false);
    store.setState((state) => ({
      ...state,
      conditional: {
        rules: [
          {
            kind: 'duplicates',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
            apply: { fill: '#f00' },
          },
        ],
      },
    }));

    try {
      executeRibbonClearAction({ store, workbook, history, action: 'all' });
      expect(workbook.getValue(writable)).toEqual({ kind: 'blank' });
      expect(workbook.getValue(locked)).toEqual({ kind: 'number', value: 2 });
      expect(commentAt(store.getState(), writable)).toBeNull();
      expect(commentAt(store.getState(), locked)).toBe('locked');
      expect(store.getState().format.formats.get(key(locked))?.italic).toBe(true);
      expect(store.getState().conditional.rules).toHaveLength(1);
    } finally {
      workbook.dispose();
    }
  });

  it('clears a conditional rule on a protected cell that was unlocked before format clearing', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const history = new History();
    const addr = { sheet: 0, row: 0, col: 0 };
    workbook.setNumber(addr, 7);
    mutators.replaceCells(store, workbook.cells(0));
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    mutators.setCellFormat(store, addr, { bold: true });
    setProtectedSheet(store, 0, true);
    setCellLocked(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, false);
    store.setState((state) => ({
      ...state,
      conditional: {
        rules: [
          {
            kind: 'duplicates',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            apply: { fill: '#f00' },
          },
        ],
      },
    }));

    try {
      executeRibbonClearAction({ store, workbook, history, action: 'all' });
      expect(workbook.getValue(addr)).toEqual({ kind: 'blank' });
      expect(store.getState().format.formats.get(key(addr))).toBeUndefined();
      expect(store.getState().conditional.rules).toHaveLength(0);
      expect(history.canUndo()).toBe(true);

      expect(history.undo()).toBe(true);
      expect(workbook.getValue(addr)).toEqual({ kind: 'number', value: 7 });
      expect(store.getState().format.formats.get(key(addr))).toMatchObject({
        bold: true,
        locked: false,
      });
      expect(store.getState().conditional.rules).toHaveLength(1);
    } finally {
      workbook.dispose();
    }
  });

  it('restores only the target hyperlink and preserves an outside format edit on undo', () => {
    const store = createSpreadsheetStore();
    const workbook = makeWorkbook();
    const history = new History();
    const target = { sheet: 0, row: 0, col: 0 };
    const outside = { sheet: 0, row: 4, col: 4 };
    mutators.setCellFormat(store, target, {
      hyperlink: 'https://target.example',
      hyperlinkDisplay: 'Target',
      hyperlinkTooltip: 'target tip',
    });
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    executeRibbonClearAction({ store, workbook, history, action: 'hyperlinks' });
    mutators.setCellFormat(store, outside, { fill: '#00ff00' });
    expect(history.undo()).toBe(true);
    expect(store.getState().format.formats.get(key(target))).toMatchObject({
      hyperlink: 'https://target.example',
      hyperlinkDisplay: 'Target',
    });
    expect(store.getState().format.formats.get(key(outside))?.fill).toBe('#00ff00');
  });

  it('does not let a comment or hyperlink undo overwrite a later locked-cell edit', () => {
    const store = createSpreadsheetStore();
    const workbook = makeWorkbook();
    const commentHistory = new History();
    const hyperlinkHistory = new History();
    const target = { sheet: 0, row: 0, col: 0 };
    const locked = { sheet: 0, row: 0, col: 1 };
    setProtectedSheet(store, 0, true);
    setCellLocked(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, false);
    setComment(store, target, 'target', workbook);
    store.setState((state) => ({
      ...state,
      format: {
        ...state.format,
        formats: new Map([
          ...state.format.formats,
          [key(locked), { comment: 'locked-before', hyperlink: 'https://locked-before.example' }],
        ]),
      },
      selection: {
        ...state.selection,
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      },
    }));
    mutators.setCellFormat(store, target, { hyperlink: 'https://target-before.example' });

    executeRibbonClearAction({ store, workbook, history: commentHistory, action: 'comments' });
    mutators.setCellFormat(store, locked, { comment: 'locked-after' });
    expect(commentHistory.undo()).toBe(true);
    expect(commentAt(store.getState(), target)).toBe('target');
    expect(commentAt(store.getState(), locked)).toBe('locked-after');
    expect(commentHistory.redo()).toBe(true);
    expect(commentAt(store.getState(), locked)).toBe('locked-after');

    executeRibbonClearAction({ store, workbook, history: hyperlinkHistory, action: 'hyperlinks' });
    mutators.setCellFormat(store, locked, { hyperlink: 'https://locked-after.example' });
    expect(hyperlinkHistory.undo()).toBe(true);
    expect(store.getState().format.formats.get(key(locked))?.hyperlink).toBe(
      'https://locked-after.example',
    );
    expect(hyperlinkHistory.redo()).toBe(true);
    expect(store.getState().format.formats.get(key(locked))?.hyperlink).toBe(
      'https://locked-after.example',
    );
  });

  it('keeps metadata, pending format, and history unchanged when the content batch is rejected', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const history = new History();
    const addr = { sheet: 0, row: 0, col: 0 };
    workbook.setNumber(addr, 7);
    mutators.replaceCells(store, workbook.cells(0));
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setComment(store, addr, 'keep');
    mutators.setCellFormat(store, addr, { bold: true });
    mutators.setPendingFormat(store, { addr, format: { fill: '#ff0' } });
    vi.spyOn(workbook, 'applyCellPatchAtomic').mockImplementation(() => {
      throw new Error('injected content failure');
    });
    try {
      executeRibbonClearAction({ store, workbook, history, action: 'all' });
      expect(workbook.getValue(addr)).toEqual({ kind: 'number', value: 7 });
      expect(commentAt(store.getState(), addr)).toBe('keep');
      expect(store.getState().format.formats.get(key(addr))?.bold).toBe(true);
      expect(store.getState().ui.pendingFormat).toEqual({ addr, format: { fill: '#ff0' } });
      expect(history.canUndo()).toBe(false);
    } finally {
      workbook.dispose();
    }
  });

  it('restores pending and scoped formats when a later engine flush fails', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault();
    const history = new History();
    const target = { sheet: 0, row: 0, col: 0 };
    const outside = { sheet: 0, row: 2, col: 2 };
    expect(workbook.isStub).toBe(false);
    mutators.setCellFormat(store, target, { bold: true });
    mutators.setCellFormat(store, outside, { italic: true });
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    mutators.setPendingFormat(store, { addr: target, format: { fill: '#ff0' } });
    flushFormatToEngine(workbook, store, 0);
    const beforeTargetXf = workbook.getCellXfIndex(0, target.row, target.col);
    const beforeOutsideXf = workbook.getCellXfIndex(0, outside.row, outside.col);
    const fontFor = (addr: typeof target) => {
      const xfIndex = workbook.getCellXfIndex(addr.sheet, addr.row, addr.col);
      expect(xfIndex).not.toBeNull();
      const xf = workbook.getCellXf(xfIndex ?? 0);
      expect(xf).not.toBeNull();
      return workbook.getFontRecord(xf?.fontIndex ?? 0);
    };
    expect(fontFor(target)?.bold).toBe(true);
    expect(fontFor(outside)?.italic).toBe(true);
    const beforePending = structuredClone(store.getState().ui.pendingFormat);
    const originalAddXfRecord = workbook.addXfRecord.bind(workbook);
    let sawPendingCleared = false;
    let injected = false;
    vi.spyOn(workbook, 'addXfRecord').mockImplementation((record) => {
      if (!injected && store.getState().ui.pendingFormat === null) {
        injected = true;
        sawPendingCleared = true;
        throw new Error('injected format flush failure');
      }
      return originalAddXfRecord(record);
    });
    try {
      expect(() => executeRibbonClearAction({ store, workbook, history, action: 'all' })).toThrow(
        'injected format flush failure',
      );
      expect(sawPendingCleared).toBe(true);
      expect(workbook.getCellXfIndex(0, target.row, target.col)).toBe(beforeTargetXf);
      expect(workbook.getCellXfIndex(0, outside.row, outside.col)).toBe(beforeOutsideXf);
      expect(fontFor(target)?.bold).toBe(true);
      expect(fontFor(outside)?.italic).toBe(true);
      expect(store.getState().format.formats.get(key(target))?.bold).toBe(true);
      expect(store.getState().format.formats.get(key(outside))?.italic).toBe(true);
      expect(store.getState().ui.pendingFormat).toEqual(beforePending);
      expect(history.canUndo()).toBe(false);
    } finally {
      workbook.dispose();
    }
  });
});
