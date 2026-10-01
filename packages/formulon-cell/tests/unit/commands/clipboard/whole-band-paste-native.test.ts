// @vitest-environment node

import { describe, expect, it } from 'vitest';
import { copy } from '../../../../src/commands/clipboard/copy.js';
import { pasteSpecial } from '../../../../src/commands/clipboard/paste-special.js';
import { captureSnapshotFromCopyResult } from '../../../../src/commands/clipboard/snapshot.js';
import { History } from '../../../../src/commands/history.js';
import { addrKey } from '../../../../src/engine/address.js';
import type { Addr } from '../../../../src/engine/types.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';

describe('whole-band paste native persistence', () => {
  it.each([
    { axis: 'column', custom: true },
    { axis: 'row', custom: true },
    { axis: 'column', custom: false },
    { axis: 'row', custom: false },
  ] as const)(
    '$axis paste (custom size: $custom) restores values, notes and sizes together',
    async ({ axis, custom }) => {
      const wb = await WorkbookHandle.createDefault();
      const store = createSpreadsheetStore();
      const column = axis === 'column';
      const addr = (offset: number, band: number): Addr => ({
        sheet: 0,
        row: column ? offset : band,
        col: column ? band : offset,
      });
      const source = addr(4, 0);
      const destination = addr(4, 3);
      const tail = addr(20, 3);
      const tailNote = addr(30, 3);
      const sourceSize = custom
        ? column
          ? 82
          : 38
        : column
          ? store.getState().layout.defaultColWidth
          : store.getState().layout.defaultRowHeight;
      const destinationSize = column ? 48 : 24;
      const sourceRange = column
        ? { sheet: 0, r0: 0, c0: 0, r1: 1_048_575, c1: 0 }
        : { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 16_383 };
      const destinationRange = column
        ? { sheet: 0, r0: 0, c0: 3, r1: 1_048_575, c1: 3 }
        : { sheet: 0, r0: 3, c0: 0, r1: 3, c1: 16_383 };
      const size = (handle: WorkbookHandle, band: number): number | undefined =>
        column
          ? handle.getColumnLayouts(0).find((item) => item.first <= band && item.last >= band)
              ?.width
          : handle.getRowLayouts(0).find((item) => item.row === band)?.height;
      try {
        expect(wb.isStub).toBe(false);
        expect(wb.capabilities.commentsEnumerable).toBe(true);
        wb.setNumber(source, 7);
        wb.setNumber(tail, 99);
        wb.setCommentEntry(0, source.row, source.col, 'Alice', 'source note');
        wb.setCommentEntry(0, tailNote.row, tailNote.col, 'Bob', 'engine-only tail');
        mutators.setCellFormat(store, source, {
          bold: true,
          comment: 'source note',
          commentAuthor: 'Alice',
        });
        mutators.setCellFormat(store, tail, { italic: true });
        if (column) {
          if (custom) mutators.setColWidth(store, 0, sourceSize);
          mutators.setColWidth(store, 3, destinationSize);
          if (custom) wb.setColumnWidth(0, 0, 0, sourceSize);
          wb.setColumnWidth(0, 3, 3, destinationSize);
        } else {
          if (custom) mutators.setRowHeight(store, 0, sourceSize);
          mutators.setRowHeight(store, 3, destinationSize);
          if (custom) wb.setRowHeight(0, 0, sourceSize);
          wb.setRowHeight(0, 3, destinationSize);
        }
        wb.recalc();
        mutators.replaceCells(store, wb.cells(0));
        mutators.setRange(store, sourceRange);
        const result = copy(store.getState());
        expect(result).not.toBeNull();
        if (!result) throw new Error('missing whole-band copy');
        const snapshot = captureSnapshotFromCopyResult(store.getState(), result);
        expect(snapshot).not.toBeNull();
        if (!snapshot) throw new Error('missing whole-band snapshot');
        mutators.setActive(store, { sheet: 0, row: destinationRange.r0, col: destinationRange.c0 });
        mutators.setRange(store, destinationRange);
        const history = new History();
        wb.attachHistory(history);
        history.begin();
        try {
          expect(
            pasteSpecial(
              store.getState(),
              store,
              wb,
              snapshot,
              {
                what: 'all',
                operation: 'none',
                skipBlanks: false,
                transpose: false,
              },
              history,
            )?.writtenRange,
          ).toEqual(destinationRange);
        } finally {
          history.end();
        }
        expect(wb.getValue(destination)).toEqual({ kind: 'number', value: 7 });
        expect(store.getState().selection.active).toEqual({
          sheet: 0,
          row: destinationRange.r0,
          col: destinationRange.c0,
        });
        expect(store.getState().selection.range).toEqual(destinationRange);
        expect(wb.getValue(tail)).toEqual({ kind: 'blank' });
        expect(wb.getComment(0, tailNote.row, tailNote.col)).toBeNull();
        expect(store.getState().format.formats.get(addrKey(tailNote))?.comment).toBeUndefined();
        expect(wb.getComment(0, destination.row, destination.col)).toEqual({
          author: 'Alice',
          text: 'source note',
        });
        expect(size(wb, 3)).toBe(sourceSize);
        expect(history.undo()).toBe(true);
        expect(wb.getValue(destination)).toEqual({ kind: 'blank' });
        expect(wb.getValue(tail)).toEqual({ kind: 'number', value: 99 });
        expect(wb.getComment(0, tailNote.row, tailNote.col)).toEqual({
          author: 'Bob',
          text: 'engine-only tail',
        });
        expect(store.getState().format.formats.get(addrKey(tailNote))).toMatchObject({
          commentAuthor: 'Bob',
          comment: 'engine-only tail',
        });
        expect(size(wb, 3)).toBe(destinationSize);
        expect(history.canUndo()).toBe(false);
        expect(history.redo()).toBe(true);
        expect(wb.getComment(0, tailNote.row, tailNote.col)).toBeNull();
        expect(store.getState().format.formats.get(addrKey(tailNote))?.comment).toBeUndefined();
        const reloaded = await WorkbookHandle.loadBytes(wb.save());
        try {
          expect(reloaded.getValue(destination)).toEqual({ kind: 'number', value: 7 });
          expect(reloaded.getValue(tail)).toEqual({ kind: 'blank' });
          expect(reloaded.getComment(0, tailNote.row, tailNote.col)).toBeNull();
          expect(reloaded.getComment(0, destination.row, destination.col)).toEqual({
            author: 'Alice',
            text: 'source note',
          });
          expect(size(reloaded, 3)).toBe(sourceSize);
        } finally {
          reloaded.dispose();
        }
      } finally {
        wb.dispose();
      }
    },
  );
});
