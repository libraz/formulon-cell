import { describe, expect, it } from 'vitest';
import { clearHyperlink, setHyperlink } from '../../../src/commands/hyperlinks.js';
import { flushFormatToEngine } from '../../../src/engine/cell-format-sync.js';
import { hydrateCommentsAndHyperlinksFromEngine } from '../../../src/engine/format-sync.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';

const reload = async (wb: WorkbookHandle) => {
  const restored = await WorkbookHandle.loadBytes(wb.save());
  const store = createSpreadsheetStore();
  hydrateCommentsAndHyperlinksFromEngine(restored, store, 0);
  return { restored, store };
};

describe.skipIf(!canLoadWasm())('hyperlink round-trip', () => {
  it('keeps an in-document location through flush and reload', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      wb.addSheet?.('Sheet2');
      setHyperlink(store, { sheet: 0, row: 1, col: 1 }, '#Sheet2!A1', wb, {
        display: 'Go',
        tooltip: 'Jump',
      });
      flushFormatToEngine(wb, store, 0);
      const links = wb.getHyperlinksFull(0);
      expect(links).toHaveLength(1);
      expect(links?.[0]).toMatchObject({ target: '', location: 'Sheet2!A1', display: 'Go' });

      const { restored, store: next } = await reload(wb);
      try {
        const fmt = next.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 1 }));
        expect(fmt).toMatchObject({
          hyperlinkLocation: 'Sheet2!A1',
          hyperlinkDisplay: 'Go',
          hyperlinkTooltip: 'Jump',
        });
        flushFormatToEngine(restored, next, 0);
        expect(restored.getHyperlinksFull(0)).toEqual(links);
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('keeps a multi-cell hyperlink as one rectangle', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.addHyperlinkRange(0, 0, 0, 2, 2, 'https://example.com/', 'Link', 'Tip', '')).toBe(
        true,
      );
      const { restored, store } = await reload(wb);
      try {
        const formats = store.getState().format.formats;
        expect(formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))).toMatchObject({
          hyperlink: 'https://example.com/',
          hyperlinkSpan: { lastRow: 2, lastCol: 2 },
        });
        flushFormatToEngine(restored, store, 0);
        expect(restored.getHyperlinksFull(0)).toEqual(wb.getHyperlinksFull(0));
        expect(restored.getHyperlinksFull(0)).toHaveLength(1);

        clearHyperlink(store, { sheet: 0, row: 0, col: 0 }, restored);
        expect(restored.getHyperlinksFull(0)).toEqual([]);
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('replacing the target of a span link drops the stale location', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      setHyperlink(store, { sheet: 0, row: 0, col: 0 }, '#Sheet1!B2', wb);
      setHyperlink(store, { sheet: 0, row: 0, col: 0 }, 'https://example.com/', wb);
      const links = wb.getHyperlinksFull(0);
      expect(links).toHaveLength(1);
      expect(links?.[0]).toMatchObject({ target: 'https://example.com/', location: '' });
    } finally {
      wb.dispose();
    }
  });
});
