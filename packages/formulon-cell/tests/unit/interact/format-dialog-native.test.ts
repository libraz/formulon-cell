// @vitest-environment node

import { describe, expect, it, vi } from 'vitest';
import { applyCellStyle, applyCellStyleByName } from '../../../src/commands/cell-styles.js';
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

const xfAt = (workbook: WorkbookHandle, addr = target): ReturnType<WorkbookHandle['getCellXf']> => {
  const xfIndex = workbook.getCellXfIndex(addr.sheet, addr.row, addr.col);
  return xfIndex === null ? null : workbook.getCellXf(xfIndex);
};

const fillAt = (
  workbook: WorkbookHandle,
  addr = target,
): ReturnType<WorkbookHandle['getFillRecord']> => {
  const xf = xfAt(workbook, addr);
  return xf ? workbook.getFillRecord(xf.fillIndex) : null;
};

const effectiveFontAt = (
  workbook: WorkbookHandle,
  addr = target,
): ReturnType<WorkbookHandle['getFontRecord']> => {
  const effective = workbook.getEffectiveStyle(addr.sheet, addr.row, addr.col);
  const xf = effective ? workbook.getCellXf(effective.xfIndex) : null;
  return xf ? workbook.getFontRecord(xf.fontIndex) : null;
};

describe('Format Cells dialog native engine contract', () => {
  it('keeps recovered opaque XFs after a later unrelated flush when rollback flush also fails', async () => {
    const workbook = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    const history = new History();
    const second = { sheet: 0, row: 0, col: 1 };
    const normal = workbook.getCellXf(0);
    if (!normal) throw new Error('missing Normal XF');
    const opaqueIndex = workbook.addXfRecord({ ...normal, quotePrefix: true });
    for (const addr of [target, second]) {
      workbook.setNumber(addr, 1);
      expect(workbook.setCellXfIndex(0, addr.row, addr.col, opaqueIndex)).toBe(true);
    }
    workbook.attachStore(store);
    const before = xfAt(workbook);
    const originalSetter = workbook.setCellXfIndex.bind(workbook);
    let calls = 0;
    const setter = vi
      .spyOn(workbook, 'setCellXfIndex')
      .mockImplementation((sheet, row, col, index) => {
        calls += 1;
        return calls === 2 || calls === 3 ? false : originalSetter(sheet, row, col, index);
      });
    try {
      expect(() =>
        recordDialogFormatChange({
          history,
          store,
          workbook,
          sheet: 0,
          targets: [target, second],
          pendingBefore: null,
          mutate: () => {
            mutators.setCellFormat(store, target, { align: 'center' });
            mutators.setCellFormat(store, second, { align: 'right' });
            return true;
          },
        }),
      ).toThrow(AggregateError);
      expect(history.canUndo()).toBe(false);
      expect(store.getState().format.formats.size).toBe(0);
      for (const addr of [target, second]) {
        expect(workbook.getCellXfIndex(0, addr.row, addr.col)).toBe(opaqueIndex);
        expect(xfAt(workbook, addr)).toEqual(before);
      }
      setter.mockRestore();
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 2 }, { italic: true });
      expect(workbook.getCellXfIndex(0, 0, 2)).toBeGreaterThan(0);
      for (const addr of [target, second]) {
        expect(workbook.getCellXfIndex(0, addr.row, addr.col)).toBe(opaqueIndex);
        expect(xfAt(workbook, addr)).toEqual(before);
      }
    } finally {
      setter.mockRestore();
      workbook.detachStore(store);
      workbook.dispose();
    }
  });

  it.each(['id', 'name'] as const)(
    'restores native XF through the public style API by %s without a workbook argument',
    async (kind) => {
      const workbook = await WorkbookHandle.createDefault();
      const store = createSpreadsheetStore();
      const history = new History();
      workbook.attachStore(store);
      try {
        workbook.setNumber(target, 1);
        mutators.setCellFormat(store, target, { bold: true });
        const current = xfAt(workbook);
        if (!current) throw new Error('missing target XF');
        const opaqueIndex = workbook.addXfRecord({ ...current, quotePrefix: true });
        expect(workbook.setCellXfIndex(0, 0, 0, opaqueIndex)).toBe(true);
        const before = xfAt(workbook);
        const range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };
        if (kind === 'id') applyCellStyle(store, history, range, 'good');
        else expect(applyCellStyleByName(store, history, range, 'good')).toBe(true);
        const afterIndex = workbook.getCellXfIndex(0, 0, 0);
        const after = xfAt(workbook);
        expect(afterIndex).not.toBe(opaqueIndex);
        expect(after?.quotePrefix).toBe(true);
        expect(history.undo()).toBe(true);
        expect(workbook.getCellXfIndex(0, 0, 0)).toBe(opaqueIndex);
        expect(xfAt(workbook)).toEqual(before);
        expect(history.redo()).toBe(true);
        expect(workbook.getCellXfIndex(0, 0, 0)).toBe(afterIndex);
        expect(xfAt(workbook)).toEqual(after);
      } finally {
        workbook.detachStore(store);
        workbook.dispose();
      }
    },
  );

  it('keeps an opaque native XF when a muted mutation throws before publication', async () => {
    const workbook = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    workbook.attachStore(store);
    try {
      workbook.setNumber(target, 42);
      const normal = workbook.getCellStyleXf(0);
      expect(normal).not.toBeNull();
      if (!normal) return;
      const sourceIndex = workbook.addXfRecord({
        ...normal,
        applyNumberFormat: true,
        applyFont: false,
        applyFill: true,
        applyBorder: false,
        applyAlignment: false,
        applyProtection: true,
        quotePrefix: true,
        hasProtection: true,
        locked: false,
        hidden: true,
      });
      expect(workbook.setCellXfIndex(0, 0, 0, sourceIndex)).toBe(true);
      const beforeIndex = workbook.getCellXfIndex(0, 0, 0);
      const beforeXf = xfAt(workbook);

      expect(() =>
        recordDialogFormatChange({
          history: new History(),
          store,
          workbook,
          sheet: 0,
          targets: [target],
          pendingBefore: null,
          mutate: () => {
            mutators.setCellFormat(store, target, { align: 'center' });
            throw new Error('dialog mutation failed before native publication');
          },
        }),
      ).toThrow('dialog mutation failed before native publication');
      expect(store.getState().format.formats.get('0:0:0')).toBeUndefined();
      expect(workbook.getCellXfIndex(0, 0, 0)).toBe(beforeIndex);
      expect(xfAt(workbook)).toEqual(beforeXf);
    } finally {
      workbook.detachStore(store);
      workbook.dispose();
    }
  });

  it('restores the visible inherited row font after undoing a cell alignment edit', async () => {
    const workbook = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    const history = new History();
    workbook.attachStore(store);
    try {
      const normal = workbook.getCellStyleXf(0);
      const defaultFont = workbook.getFontRecord(0);
      expect(normal).not.toBeNull();
      expect(defaultFont).not.toBeNull();
      if (!normal || !defaultFont) return;
      const inheritedFont = workbook.addFontRecord({ ...defaultFont, size: 22 });
      const rowStyle = workbook.addXfRecord({ ...normal, fontIndex: inheritedFont });
      expect(rowStyle).toBeGreaterThanOrEqual(0);
      expect(workbook.setRowStyle(0, 0, rowStyle)).toBe(true);
      expect(workbook.getEffectiveStyle(0, 0, 0)?.source).toBe(1);
      expect(effectiveFontAt(workbook)?.size).toBe(22);

      expect(
        recordDialogFormatChange({
          history,
          store,
          workbook,
          sheet: 0,
          targets: [target],
          pendingBefore: null,
          mutate: () => {
            mutators.setCellFormat(store, target, { align: 'center' });
            return true;
          },
        }),
      ).toBe(true);
      expect(effectiveFontAt(workbook)?.size).toBe(22);

      expect(history.undo()).toBe(true);
      expect(effectiveFontAt(workbook)?.size).toBe(22);
      // Native clear now restores the blank cell's row inheritance instead
      // of copying the effective font into a permanent direct cell XF.
      expect(workbook.getEffectiveStyle(0, 0, 0)?.source).toBe(1);
      const laterFont = workbook.addFontRecord({ ...defaultFont, size: 26 });
      const laterRowStyle = workbook.addXfRecord({ ...normal, fontIndex: laterFont });
      expect(workbook.setRowStyle(0, 0, laterRowStyle)).toBe(true);
      expect(effectiveFontAt(workbook)?.size).toBe(26);
      expect(workbook.getEffectiveStyle(0, 0, 0)?.source).toBe(1);
      expect(history.redo()).toBe(true);
      expect(effectiveFontAt(workbook)?.size).toBe(22);
    } finally {
      workbook.detachStore(store);
      workbook.dispose();
    }
  });

  it('restores exact native alignment XF on undo and redo with a mounted store', async () => {
    const workbook = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    const history = new History();
    const outside = { sheet: 0, row: 0, col: 1 };
    workbook.attachStore(store);
    try {
      workbook.setNumber(target, 1);
      workbook.setNumber(outside, 2);
      mutators.setCellFormat(store, target, { bold: true });
      mutators.setCellFormat(store, outside, { italic: true });
      const beforeIndex = workbook.getCellXfIndex(0, 0, 0);
      const beforeXf = xfAt(workbook);
      const outsideIndex = workbook.getCellXfIndex(0, 0, 1);
      const outsideXf = xfAt(workbook, outside);

      expect(
        recordDialogFormatChange({
          history,
          store,
          workbook,
          sheet: 0,
          targets: [target],
          pendingBefore: null,
          mutate: () => {
            mutators.setCellFormat(store, target, { align: 'center' });
            return true;
          },
        }),
      ).toBe(true);
      const afterIndex = workbook.getCellXfIndex(0, 0, 0);
      const afterXf = xfAt(workbook);
      expect(afterXf?.horizontalAlign).toBe(2);
      expect(afterIndex).not.toBe(beforeIndex);

      expect(history.undo()).toBe(true);
      expect(workbook.getCellXfIndex(0, 0, 0)).toBe(beforeIndex);
      expect(xfAt(workbook)).toEqual(beforeXf);
      expect(workbook.getCellXfIndex(0, 0, 1)).toBe(outsideIndex);
      expect(xfAt(workbook, outside)).toEqual(outsideXf);

      expect(history.redo()).toBe(true);
      expect(workbook.getCellXfIndex(0, 0, 0)).toBe(afterIndex);
      expect(xfAt(workbook)).toEqual(afterXf);
      expect(workbook.getCellXfIndex(0, 0, 1)).toBe(outsideIndex);
      expect(xfAt(workbook, outside)).toEqual(outsideXf);
    } finally {
      workbook.detachStore(store);
      workbook.dispose();
    }
  });

  it('restores exact native fill XF when a dialog adds and removes a fill', async () => {
    const workbook = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    const history = new History();
    workbook.attachStore(store);
    try {
      workbook.setNumber(target, 1);
      mutators.setCellFormat(store, target, { cellStyle: 'good', bold: true });
      const beforeIndex = workbook.getCellXfIndex(0, 0, 0);
      const beforeXf = xfAt(workbook);
      expect(fillAt(workbook)?.pattern).toBe(0);

      expect(
        recordDialogFormatChange({
          history,
          store,
          workbook,
          sheet: 0,
          targets: [target],
          pendingBefore: null,
          mutate: () => {
            mutators.setCellFormat(store, target, { fill: '#ffeeaa' });
            return true;
          },
        }),
      ).toBe(true);
      const afterIndex = workbook.getCellXfIndex(0, 0, 0);
      const afterXf = xfAt(workbook);
      expect(fillAt(workbook)?.pattern).toBe(1);

      expect(history.undo()).toBe(true);
      expect(workbook.getCellXfIndex(0, 0, 0)).toBe(beforeIndex);
      expect(xfAt(workbook)).toEqual(beforeXf);
      expect(fillAt(workbook)?.pattern).toBe(0);

      expect(history.redo()).toBe(true);
      expect(workbook.getCellXfIndex(0, 0, 0)).toBe(afterIndex);
      expect(xfAt(workbook)).toEqual(afterXf);
      expect(fillAt(workbook)?.pattern).toBe(1);
    } finally {
      workbook.detachStore(store);
      workbook.dispose();
    }
  });

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
      expect(fontAt(workbook)?.size).toBe(9);

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
      expect(fontAt(workbook)?.size).toBe(5.25);
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
      expect(fontAt(workbook)?.size).toBe(9);
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
      expect(fontAt(workbook)?.size).toBe(5.25);
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
      const targetIndex = workbook.getCellXfIndex(0, target.row, target.col);
      const targetXf = xfAt(workbook, target);
      const otherIndex = workbook.getCellXfIndex(0, other.row, other.col);
      const otherXf = xfAt(workbook, other);

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
      expect(workbook.getCellXfIndex(0, target.row, target.col)).toBe(targetIndex);
      expect(xfAt(workbook, target)).toEqual(targetXf);
      expect(workbook.getCellXfIndex(0, other.row, other.col)).toBe(otherIndex);
      expect(xfAt(workbook, other)).toEqual(otherXf);
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
