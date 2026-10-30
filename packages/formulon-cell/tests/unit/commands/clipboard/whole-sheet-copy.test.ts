import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { copy } from '../../../../src/commands/clipboard/copy.js';
import { cut } from '../../../../src/commands/clipboard/cut.js';
import { pasteSpecial } from '../../../../src/commands/clipboard/paste-special.js';
import {
  captureSnapshot,
  captureSnapshotFromCopyResult,
} from '../../../../src/commands/clipboard/snapshot.js';
import { parseTSV } from '../../../../src/commands/clipboard/tsv.js';
import { History } from '../../../../src/commands/history.js';
import { addrKey, MAX_COL, MAX_ROW } from '../../../../src/engine/address.js';
import type { Range } from '../../../../src/engine/types.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { attachClipboard } from '../../../../src/interact/clipboard.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import { assertSnap, seedAndMirror, setActive, setSelection } from './paste-special/fixtures.js';

const fullSheet: Range = {
  sheet: 0,
  r0: 0,
  c0: 0,
  r1: MAX_ROW,
  c1: MAX_COL,
};

const fullSheetSelection = (store: SpreadsheetStore): void => {
  setSelection(store, 0, 0, MAX_ROW, MAX_COL);
};

const fireClipboard = (host: HTMLElement, type: 'copy' | 'cut'): DataTransfer => {
  const transfer = new DataTransfer();
  const event = new ClipboardEvent(type, {
    clipboardData: transfer,
    bubbles: true,
    cancelable: true,
  });
  host.dispatchEvent(event);
  return transfer;
};

describe('whole-sheet copy', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await WorkbookHandle.createDefault();
  });

  afterEach(() => {
    wb.dispose();
    document.body.innerHTML = '';
  });

  it('compacts a sparse full-sheet copy to the values, formats, blank cells, and merges in use', () => {
    seedAndMirror(store, wb, [
      { row: 4, col: 3, value: 7, formula: '=A1' },
      { row: 5, col: 4, value: 'later' },
    ]);
    store.setState((state) => {
      const cells = new Map(state.data.cells);
      cells.set(addrKey({ sheet: 0, row: 2, col: 12 }), {
        value: { kind: 'blank' },
        formula: null,
      });
      const formats = new Map(state.format.formats);
      formats.set(addrKey({ sheet: 0, row: 8, col: 7 }), { bold: true, comment: 'note' });
      return {
        ...state,
        data: { ...state.data, cells },
        format: { ...state.format, formats },
        layout: {
          ...state.layout,
          rowHeights: new Map([[10, 31]]),
          colWidths: new Map([[12, 77]]),
        },
      };
    });
    mutators.mergeRange(store, { sheet: 0, r0: 10, c0: 9, r1: 11, c1: 10 });
    fullSheetSelection(store);

    const copied = copy(store.getState());

    expect(copied).not.toBeNull();
    expect(copied?.range).toEqual(fullSheet);
    expect(copied?.logicalRange).toEqual(fullSheet);
    expect(copied?.payloadRanges).toEqual([{ sheet: 0, r0: 2, c0: 3, r1: 11, c1: 12 }]);
    const matrix = parseTSV(copied?.tsv ?? '');
    expect(matrix).toHaveLength(10);
    expect(matrix[0]).toHaveLength(10);
    expect(matrix[0]?.[9]).toBe('');
    expect(matrix[2]?.[0]).toBe('=A1');
    expect(matrix[3]?.[1]).toBe('later');

    expect(captureSnapshot(store.getState(), fullSheet)).toBeNull();
    const snapshot = copied ? captureSnapshotFromCopyResult(store.getState(), copied) : null;
    assertSnap(snapshot);
    expect(snapshot.range).toEqual({ sheet: 0, r0: 2, c0: 3, r1: 11, c1: 12 });
    expect(snapshot.logicalRange).toEqual(fullSheet);
    expect(snapshot.rows).toBe(10);
    expect(snapshot.cols).toBe(10);
    expect(snapshot.cells[6]?.[4]?.format).toMatchObject({ bold: true, comment: 'note' });
    expect(snapshot.merges).toEqual([{ r0: 8, c0: 6, r1: 9, c1: 7 }]);
    expect(snapshot.rowHeights?.get(10)).toBe(31);
    expect(snapshot.colWidths?.get(12)).toBe(77);
  });

  it('uses A1 for an empty full-sheet payload and keeps the million-cell materialization cap', () => {
    fullSheetSelection(store);
    const empty = copy(store.getState());
    expect(empty?.range).toEqual(fullSheet);
    expect(empty?.payloadRanges).toEqual([{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }]);
    expect(empty?.tsv).toBe('');

    store.setState((state) => {
      const cells = new Map(state.data.cells);
      cells.set(addrKey({ sheet: 0, row: 0, col: 0 }), {
        value: { kind: 'number', value: 1 },
        formula: null,
      });
      cells.set(addrKey({ sheet: 0, row: 1_000, col: 1_000 }), {
        value: { kind: 'number', value: 2 },
        formula: null,
      });
      return { ...state, data: { ...state.data, cells } };
    });
    expect(copy(store.getState())).toBeNull();
  });

  it('pastes a full-sheet snapshot to A1 on another sheet with offsets, merges, layout, clearing, and one undo step', () => {
    seedAndMirror(store, wb, [
      { row: 4, col: 3, value: 7, formula: '=A1' },
      { row: 5, col: 4, value: 'source' },
    ]);
    store.setState((state) => {
      const cells = new Map(state.data.cells);
      cells.set(addrKey({ sheet: 0, row: 2, col: 12 }), {
        value: { kind: 'blank' },
        formula: null,
      });
      const formats = new Map(state.format.formats);
      formats.set(addrKey({ sheet: 0, row: 8, col: 7 }), {
        bold: true,
        comment: 'source note',
        commentAuthor: 'Source',
      });
      return {
        ...state,
        data: { ...state.data, cells },
        format: { ...state.format, formats },
        layout: {
          ...state.layout,
          rowHeights: new Map([[10, 31]]),
          colWidths: new Map([[12, 77]]),
        },
      };
    });
    const sourceMerge = { sheet: 0, r0: 10, c0: 9, r1: 11, c1: 10 };
    mutators.mergeRange(store, sourceMerge);
    expect(wb.engineAddMerge(0, sourceMerge)).toBe(true);
    fullSheetSelection(store);

    const copied = copy(store.getState());
    assertSnap(copied);
    const snapshot = captureSnapshotFromCopyResult(store.getState(), copied);
    assertSnap(snapshot);

    expect(wb.addSheet('Destination')).toBe(1);
    wb.setText({ sheet: 1, row: 100, col: 100 }, 'stale');
    wb.setCommentEntry(1, 100, 100, 'Old author', 'stale note');
    const staleMerge = { sheet: 1, r0: 99, c0: 99, r1: 100, c1: 100 };
    expect(wb.engineAddMerge(1, staleMerge)).toBe(true);
    mutators.mergeRange(store, staleMerge);
    mutators.setCellFormat(
      store,
      { sheet: 1, row: 100, col: 100 },
      {
        italic: true,
        comment: 'stale note',
        commentAuthor: 'Old author',
      },
    );
    store.setState((state) => ({
      ...state,
      data: {
        ...state.data,
        sheetIndex: 1,
        cells: new Map([
          [
            addrKey({ sheet: 1, row: 100, col: 100 }),
            { value: { kind: 'text', value: 'stale' }, formula: null },
          ],
        ]),
      },
      layout: {
        ...state.layout,
        rowHeights: new Map([[100, 22]]),
        colWidths: new Map([[100, 55]]),
      },
    }));
    setActive(store, 0, 0, 1);

    const history = new History();
    wb.attachHistory(history);
    history.begin();
    const result = pasteSpecial(
      store.getState(),
      store,
      wb,
      snapshot,
      { what: 'all', operation: 'none', skipBlanks: false, transpose: false },
      history,
    );
    history.end();
    wb.recalc();

    expect(result?.writtenRange).toEqual({ ...fullSheet, sheet: 1 });
    expect(wb.cellFormula({ sheet: 1, row: 4, col: 3 })).toBe('=A1');
    expect(wb.getValue({ sheet: 1, row: 5, col: 4 })).toEqual({ kind: 'text', value: 'source' });
    expect(wb.getValue({ sheet: 1, row: 100, col: 100 })).toEqual({ kind: 'blank' });
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 1, row: 8, col: 7 })),
    ).toMatchObject({
      bold: true,
      comment: 'source note',
    });
    const clearedFormat = store
      .getState()
      .format.formats.get(addrKey({ sheet: 1, row: 100, col: 100 }));
    expect(clearedFormat?.italic).toBeUndefined();
    expect(clearedFormat?.comment).toBeUndefined();
    expect(wb.getComments(1).some((comment) => comment.row === 100 && comment.col === 100)).toBe(
      false,
    );
    expect(wb.getMerges(1)).toContainEqual({ sheet: 1, r0: 10, c0: 9, r1: 11, c1: 10 });
    expect(wb.getMerges(1)).not.toContainEqual(staleMerge);
    expect(store.getState().layout.rowHeights.get(10)).toBe(31);
    expect(store.getState().layout.rowHeights.has(100)).toBe(false);
    expect(store.getState().layout.colWidths.get(12)).toBe(77);
    expect(store.getState().layout.colWidths.has(100)).toBe(false);

    expect(history.undo()).toBe(true);
    expect(wb.getValue({ sheet: 1, row: 100, col: 100 })).toEqual({ kind: 'text', value: 'stale' });
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 1, row: 100, col: 100 })),
    ).toMatchObject({
      italic: true,
      comment: 'stale note',
    });
    expect(wb.getMerges(1)).toContainEqual(staleMerge);
    expect(store.getState().layout.rowHeights.get(100)).toBe(22);
    expect(store.getState().layout.colWidths.get(100)).toBe(55);

    expect(history.redo()).toBe(true);
    expect(wb.cellFormula({ sheet: 1, row: 4, col: 3 })).toBe('=A1');
    expect(wb.getValue({ sheet: 1, row: 100, col: 100 })).toEqual({ kind: 'blank' });
    expect(wb.getMerges(1)).toContainEqual({ sheet: 1, r0: 10, c0: 9, r1: 11, c1: 10 });
    expect(store.getState().layout.rowHeights.get(10)).toBe(31);
    expect(store.getState().layout.colWidths.get(12)).toBe(77);
  });

  it('refreshes the full-sheet source with one sparse range query and never enumerates the grid', () => {
    seedAndMirror(store, wb, [{ row: 4, col: 3, value: 'first' }]);
    fullSheetSelection(store);
    const host = document.createElement('div');
    document.body.appendChild(host);
    const handle = attachClipboard({ host, store, wb, onAfterCommit: () => {} });
    fireClipboard(host, 'copy');
    // Extend beyond the captured payload without creating a huge blank matrix.
    wb.setText({ sheet: 0, row: 200, col: 20 }, 'later');

    const ranges: Range[] = [];
    const original = wb.cellsInRange.bind(wb);
    vi.spyOn(wb, 'cellsInRange').mockImplementation(function* (range, pageSize) {
      ranges.push({ ...range });
      yield* original(range, pageSize);
    });
    const fullSheetScans = vi.spyOn(wb, 'cells').mockImplementation(() => {
      throw new Error('unexpected full-sheet cell scan');
    });

    const refreshed = handle.getSnapshot();

    expect(ranges).toEqual([fullSheet]);
    expect(fullSheetScans).not.toHaveBeenCalled();
    expect(refreshed?.range).toEqual({ sheet: 0, r0: 4, c0: 3, r1: 200, c1: 20 });
    expect(refreshed?.cells.at(-1)?.at(-1)?.value).toEqual({ kind: 'text', value: 'later' });
    handle.detach();
  });

  it('rejects direct, event-driven, and shortcut full-sheet cuts', async () => {
    fullSheetSelection(store);
    expect(cut(store.getState(), wb)).toBeNull();

    const host = document.createElement('div');
    document.body.appendChild(host);
    const handle = attachClipboard({ host, store, wb, onAfterCommit: () => {} });

    fireClipboard(host, 'cut');
    expect(handle.getSnapshot()).toBeNull();
    expect(store.getState().ui.copyMode).not.toBe('cut');

    await handle.runShortcut('cut');
    expect(handle.getSnapshot()).toBeNull();
    expect(store.getState().ui.copyMode).not.toBe('cut');
    handle.detach();
  });
});
