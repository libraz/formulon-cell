import { describe, expect, it } from 'vitest';
import { shiftHyperlinkSpans } from '../../../../src/commands/axis-shift.js';
import { History } from '../../../../src/commands/history.js';
import { deleteCols, deleteRows, insertRows } from '../../../../src/commands/structure.js';
import { flushFormatToEngine } from '../../../../src/engine/cell-format-sync.js';
import { hydrateCommentsAndHyperlinksFromEngine } from '../../../../src/engine/format-sync.js';
import { addrKey, WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import type { CellFormat } from '../../../../src/store/store.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';

const link = (lastRow: number, lastCol: number): CellFormat => ({
  hyperlink: 'https://example.com/',
  hyperlinkDisplay: 'Go',
  hyperlinkSpan: { lastRow, lastCol },
});

describe('shiftHyperlinkSpans', () => {
  const src = new Map([['0:2:1', link(5, 3)]]);

  it('grows a span for an insert inside it and moves it for one above it', () => {
    const grown = shiftHyperlinkSpans(src, new Map([['0:2:1', link(5, 3)]]), 0, 'row', 4, 2);
    expect(grown.get('0:2:1')?.hyperlinkSpan).toEqual({ lastRow: 7, lastCol: 3 });

    const moved = shiftHyperlinkSpans(src, new Map([['0:4:1', link(5, 3)]]), 0, 'row', 0, 2);
    expect(moved.get('0:4:1')?.hyperlinkSpan).toEqual({ lastRow: 7, lastCol: 3 });
  });

  it('shrinks a span for a delete inside it and leaves other sheets alone', () => {
    const shrunk = shiftHyperlinkSpans(src, new Map([['0:2:1', link(5, 3)]]), 0, 'col', 2, -2);
    expect(shrunk.get('0:2:1')?.hyperlinkSpan).toEqual({ lastRow: 5, lastCol: 1 });

    const other = new Map([['0:2:1', link(5, 3)]]);
    expect(shiftHyperlinkSpans(src, other, 1, 'row', 3, -1)).toBe(other);
  });

  it('collapses a span that shrinks to its anchor into a plain link', () => {
    const column = new Map([['0:2:1', link(5, 1)]]);
    const out = shiftHyperlinkSpans(column, new Map([['0:2:1', link(5, 1)]]), 0, 'row', 3, -3);
    expect(out.get('0:2:1')?.hyperlinkSpan).toBeUndefined();
    expect(out.get('0:2:1')?.hyperlink).toBe('https://example.com/');
  });

  it('moves the link to the first surviving cell when its anchor is deleted', () => {
    // Rows 1..3 deleted: the anchor at row 2 goes, rows 4..5 survive as rows 1..2.
    const shifted = new Map<string, CellFormat>([['0:1:1', { bold: true }]]);
    const out = shiftHyperlinkSpans(src, shifted, 0, 'row', 1, -3);
    expect(out.get('0:1:1')).toMatchObject({
      bold: true,
      hyperlink: 'https://example.com/',
      hyperlinkDisplay: 'Go',
      hyperlinkSpan: { lastRow: 2, lastCol: 3 },
    });
  });

  it('drops the link when its whole span is deleted', () => {
    const out = shiftHyperlinkSpans(src, new Map(), 0, 'row', 2, -4);
    expect(out.size).toBe(0);
  });
});

describe('hyperlink spans through row and column edits', () => {
  const setup = async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    expect(wb.addHyperlinkRange(0, 2, 1, 5, 3, 'https://example.com/', 'Go', 'Tip', '')).toBe(true);
    hydrateCommentsAndHyperlinksFromEngine(wb, store, 0);
    return { wb, store };
  };
  const anchor = () => addrKey({ sheet: 0, row: 2, col: 1 });

  it('grows with an insert inside the span and reaches the saved file', async () => {
    const { wb, store } = await setup();
    try {
      expect(insertRows(store, wb, null, 4, 2)).toBe(true);
      expect(store.getState().format.formats.get(anchor())?.hyperlinkSpan).toEqual({
        lastRow: 7,
        lastCol: 3,
      });
      flushFormatToEngine(wb, store, 0);
      const saved = await WorkbookHandle.loadBytes(wb.save());
      try {
        expect(saved.getHyperlinksFull(0)).toMatchObject([
          { row: 2, col: 1, lastRow: 7, lastCol: 3, target: 'https://example.com/' },
        ]);
      } finally {
        saved.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('shrinks with a column delete and moves with a row insert above it', async () => {
    const { wb, store } = await setup();
    try {
      expect(deleteCols(store, wb, null, 2, 1)).toBe(true);
      expect(store.getState().format.formats.get(anchor())?.hyperlinkSpan).toEqual({
        lastRow: 5,
        lastCol: 2,
      });
      expect(insertRows(store, wb, null, 0, 1)).toBe(true);
      const moved = addrKey({ sheet: 0, row: 3, col: 1 });
      expect(store.getState().format.formats.get(moved)?.hyperlinkSpan).toEqual({
        lastRow: 6,
        lastCol: 2,
      });
      expect(store.getState().format.formats.get(anchor())?.hyperlink).toBeUndefined();
    } finally {
      wb.dispose();
    }
  });

  it('keeps the link on the surviving rows when the anchor row is deleted, and undoes', async () => {
    const { wb, store } = await setup();
    const history = new History();
    try {
      expect(deleteRows(store, wb, history, 1, 2)).toBe(true);
      expect(
        store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 1 })),
      ).toMatchObject({
        hyperlink: 'https://example.com/',
        hyperlinkSpan: { lastRow: 3, lastCol: 3 },
      });
      history.undo();
      expect(store.getState().format.formats.get(anchor())?.hyperlinkSpan).toEqual({
        lastRow: 5,
        lastCol: 3,
      });
    } finally {
      wb.dispose();
    }
  });
});
