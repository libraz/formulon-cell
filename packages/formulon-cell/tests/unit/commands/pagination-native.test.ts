import { describe, expect, it } from 'vitest';
import {
  PAGE_GUTTER_PX,
  pageBoxPx,
  pageLayoutGaps,
  paginationFor,
} from '../../../src/commands/pagination.js';
import { hydrateLayoutFromEngine } from '../../../src/engine/layout-sync.js';
import { syncPageSetupToEngine } from '../../../src/engine/print-sync.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, getPageSetup, mutators } from '../../../src/store/store.js';

describe('native pagination integration', () => {
  it('uses native page bands, break provenance, scale, and point margins', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    try {
      expect(wb.capabilities.detailedPagination).toBe(true);
      for (let row = 0; row < 160; row += 1) {
        wb.setNumber({ sheet: 0, row, col: 0 }, row + 1);
        store.setState((s) => {
          const cells = new Map(s.data.cells);
          cells.set(addrKey({ sheet: 0, row, col: 0 }), {
            value: { kind: 'number', value: row + 1 },
            formula: null,
          });
          return { ...s, data: { ...s.data, cells } };
        });
      }
      mutators.setPageSetup(store, 0, {
        printArea: 'A1:A160',
        printTitleRows: '1:1',
        manualPageBreakRows: [60],
        scale: 0.5,
      });
      syncPageSetupToEngine(wb, store, 0);

      const pagination = paginationFor(store.getState(), 0, { engine: wb });
      expect(pagination.native).toBeDefined();
      expect(pagination.native?.pages.length).toBe(pagination.native?.pageCount);
      expect(pagination.rowBands[0]?.start).toBe(0);
      expect(pagination.rowBands.some((band) => band.start === 60 && band.manual)).toBe(true);
      expect(pagination.scale).toBeCloseTo(0.5);
      // Native margins are points; the view consumes CSS pixels at 96/72.
      expect(pagination.marginPx.top).toBeCloseTo((0.75 * 96) / pagination.scale, 5);
      expect(pagination.pagePx.width).toBeCloseTo(
        ((pagination.native?.printable.width ?? 0) * (96 / 72)) / pagination.scale,
        5,
      );
      const nextRowBand = pagination.rowBands[1];
      expect(nextRowBand).toBeDefined();
      const nativePaper = pageBoxPx(pagination, getPageSetup(store.getState(), 0));
      const gaps = pageLayoutGaps(store.getState(), 0, { engine: wb }).pageGapRows;
      const rowGap = gaps.get(nextRowBand?.start ?? -1);
      const previousHeight = pagination.rowBands[0]
        ? pagination.rowBands[0].end - pagination.rowBands[0].start + 1
        : 0;
      const firstPage = pagination.native?.pages.find(
        (page) =>
          page.firstRow === pagination.rowBands[0]?.start &&
          page.lastRow === pagination.rowBands[0]?.end,
      );
      const secondPage = pagination.native?.pages.find(
        (page) => page.firstRow === nextRowBand?.start && page.lastRow === nextRowBand?.end,
      );
      const originDelta =
        firstPage && secondPage
          ? ((secondPage.originYPt - firstPage.originYPt) * (96 / 72)) / pagination.scale
          : 0;
      expect(rowGap).toBeCloseTo(
        Math.max(
          0,
          nativePaper.height - previousHeight * store.getState().layout.defaultRowHeight,
        ) +
          originDelta +
          PAGE_GUTTER_PX,
        5,
      );

      // The physical paper starts stay one paper plus gutter apart even when
      // the next page reserves repeated title rows inside its origin.
      const zoom = store.getState().viewport.zoom;
      const firstGap = gaps.get(pagination.rowBands[0]?.start ?? -1) ?? 0;
      const firstBodyStart = firstGap;
      const secondBodyStart =
        firstBodyStart +
        previousHeight * store.getState().layout.defaultRowHeight * zoom +
        (rowGap ?? 0);
      const factor = ((96 / 72) * zoom) / pagination.scale;
      const firstPaperStart = firstBodyStart - (firstPage?.originYPt ?? 0) * factor;
      const secondPaperStart = secondBodyStart - (secondPage?.originYPt ?? 0) * factor;
      expect(secondPaperStart - firstPaperStart).toBeCloseTo(
        nativePaper.height * zoom + PAGE_GUTTER_PX,
        5,
      );

      const tailBand = pagination.rowBands.at(-1);
      const previousTailBand = pagination.rowBands.at(-2);
      const previousTailPage = pagination.native?.pages.find(
        (page) =>
          page.firstRow === previousTailBand?.start && page.lastRow === previousTailBand?.end,
      );
      if (tailBand && previousTailBand && previousTailPage) {
        const tailGap = gaps.get(tailBand.start);
        // Blank bands carry the previous native cell origin, so their
        // physical paper advances by one page without another title offset.
        const previousTailHeight =
          (previousTailBand.end - previousTailBand.start + 1) *
          store.getState().layout.defaultRowHeight;
        expect(tailGap).toBeCloseTo(nativePaper.height - previousTailHeight + PAGE_GUTTER_PX, 5);
      }
    } finally {
      wb.dispose();
    }
  });

  it('keeps multi-area preview rectangular while print consumes native area order', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    try {
      wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
      wb.setNumber({ sheet: 0, row: 0, col: 2 }, 2);
      mutators.setPageSetup(store, 0, { printArea: 'A1:A2,C1:C2' });
      syncPageSetupToEngine(wb, store, 0);

      const native = wb.paginateSheet(0);
      expect(native?.printArea).toHaveLength(2);
      const preview = paginationFor(store.getState(), 0, { engine: wb });
      expect(preview.native).toBeUndefined();
      const print = paginationFor(store.getState(), 0, { engine: wb, purpose: 'print' });
      expect(print.native?.pages.length).toBe(native?.pages.length);
      expect(print.native?.pages.map((page) => page.areaIndex)).toEqual(
        native?.pages.map((page) => page.areaIndex),
      );
    } finally {
      wb.dispose();
    }
  });

  it('places centered native paper from the centered cell-block origin', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    try {
      for (let row = 0; row < 160; row += 1) {
        wb.setNumber({ sheet: 0, row, col: 0 }, row + 1);
        store.setState((s) => {
          const cells = new Map(s.data.cells);
          cells.set(addrKey({ sheet: 0, row, col: 0 }), {
            value: { kind: 'number', value: row + 1 },
            formula: null,
          });
          return { ...s, data: { ...s.data, cells } };
        });
      }
      mutators.setPageSetup(store, 0, {
        printArea: 'A1:A160',
        manualPageBreakRows: [60],
        centerHorizontally: true,
        centerVertically: true,
        scale: 0.5,
      });
      syncPageSetupToEngine(wb, store, 0);

      const state = store.getState();
      const pagination = paginationFor(state, 0, { engine: wb });
      const firstBand = pagination.rowBands[0];
      const firstPage = pagination.native?.pages.find(
        (page) => page.firstRow === firstBand?.start && page.lastRow === firstBand?.end,
      );
      expect(firstPage).toBeDefined();
      const native = pagination.native;
      if (!native || !firstPage || !firstBand) return;
      const titleOffset = Math.max(0, firstPage.originYPt - native.printable.y);
      const centerOffset = Math.max(
        0,
        (native.printable.height - titleOffset - firstPage.heightPt) / 2,
      );
      const firstGap = pageLayoutGaps(state, 0, { engine: wb }).pageGapRows.get(firstBand.start);
      expect(firstGap).toBeCloseTo(
        (pagination.marginPx.top + (centerOffset * (96 / 72)) / pagination.scale) *
          state.viewport.zoom,
        5,
      );
    } finally {
      wb.dispose();
    }
  });
});

describe('Mac pagination', () => {
  it('breaks columns on the store widths instead of the Windows print model', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'ja', platform: 'mac' });
    const store = createSpreadsheetStore();
    try {
      hydrateLayoutFromEngine(wb, store, 0);
      store.setState((s) => {
        const cells = new Map(s.data.cells);
        for (let col = 0; col < 30; col += 1) {
          wb.setNumber({ sheet: 0, row: 0, col }, col);
          cells.set(addrKey({ sheet: 0, row: 0, col }), {
            value: { kind: 'number', value: col },
            formula: null,
          });
        }
        return { ...s, data: { ...s.data, cells } };
      });

      const pagination = paginationFor(store.getState(), 0, { engine: wb });
      expect(pagination.native).toBeUndefined();
      const { colWidths, defaultColWidth } = store.getState().layout;
      for (const band of pagination.colBands) {
        let width = 0;
        for (let col = band.start; col <= band.end; col += 1) {
          width += colWidths.get(col) ?? defaultColWidth;
        }
        expect(width).toBeLessThanOrEqual(pagination.pagePx.width);
      }
    } finally {
      wb.dispose();
    }
  });
});
