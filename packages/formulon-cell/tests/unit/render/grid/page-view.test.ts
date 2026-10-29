import { beforeEach, describe, expect, it } from 'vitest';
import {
  associatePaginationEngine,
  PAGE_GUTTER_PX,
  pageBoxPx,
  paginationFor,
  resetPaginationCache,
} from '../../../../src/commands/pagination.js';
import { setWorkbookView } from '../../../../src/commands/view.js';
import { addrKey } from '../../../../src/engine/address.js';
import { syncPageSetupToEngine } from '../../../../src/engine/print-sync.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { en } from '../../../../src/i18n/strings/en.js';
import {
  buildColLayout,
  buildRowLayout,
  gridOriginY,
  layoutForView,
} from '../../../../src/render/geometry.js';
import {
  paintPageBreakPreview,
  paintPageLayoutBackground,
  paintPageLayoutChrome,
  paintPageRulers,
} from '../../../../src/render/grid/page-view.js';
import {
  createSpreadsheetStore,
  getPageSetup,
  mutators,
  type SpreadsheetStore,
  type State,
} from '../../../../src/store/store.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

const theme = {
  bg: '#fff',
  bgRail: '#faf9f8',
  fg: '#201f1e',
  rule: '#d9d9d9',
  headerFg: '#605e5c',
  fontUi: 'Aptos, sans-serif',
  pageBackdrop: '#a19c90',
  pagePaper: '#ffffff',
  pageEdge: '#8a8579',
  pageBandFg: '#8d8779',
  pageBreakAuto: '#2f6fd0',
  pageBreakManual: '#1a4fa0',
  pageNumberFg: 'rgba(47,111,208,0.16)',
  pageOutside: 'rgba(120,120,120,0.26)',
} as unknown as ResolvedTheme;

interface FillRecord {
  style: string;
  rect: [number, number, number, number];
}

interface TextRecord {
  text: string;
  x: number;
  y: number;
}

function makeCtxSpy(): {
  ctx: CanvasRenderingContext2D;
  fills: FillRecord[];
  texts: string[];
  textRecords: TextRecord[];
} {
  const fills: FillRecord[] = [];
  const texts: string[] = [];
  const textRecords: TextRecord[] = [];
  let fillStyle = '';
  const ctx = {
    get fillStyle(): string {
      return fillStyle;
    },
    set fillStyle(v: string) {
      fillStyle = v;
    },
    strokeStyle: '',
    lineWidth: 1,
    font: '',
    globalAlpha: 1,
    textBaseline: 'alphabetic',
    textAlign: 'left',
    save(): void {},
    restore(): void {},
    beginPath(): void {},
    closePath(): void {},
    rect(): void {},
    clip(): void {},
    setLineDash(): void {},
    moveTo(): void {},
    lineTo(): void {},
    stroke(): void {},
    strokeRect(): void {},
    fill(): void {},
    fillRect(x: number, y: number, w: number, h: number): void {
      fills.push({ style: fillStyle, rect: [x, y, w, h] });
    },
    fillText(text: string, x = 0, y = 0): void {
      texts.push(text);
      textRecords.push({ text, x, y });
    },
  } as unknown as CanvasRenderingContext2D;
  return { ctx, fills, texts, textRecords };
}

const seed = (store: SpreadsheetStore, row: number, col: number): void => {
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    cells.set(addrKey({ sheet: 0, row, col }), {
      value: { kind: 'number', value: 1 },
      formula: null,
    });
    return {
      ...s,
      data: { ...s.data, cells },
      viewport: { ...s.viewport, rowCount: 120, colCount: 40, widthPx: 1200 },
    };
  });
};

const viewFor = (state: State) => {
  const layout = layoutForView(state);
  return {
    view: { ...state, layout },
    cols: buildColLayout(layout, state.viewport),
    rows: buildRowLayout(layout, state.viewport),
  };
};

const pc = (ctx: CanvasRenderingContext2D) => ({ ctx, dpr: 1, cssWidth: 1200, cssHeight: 900 });

describe('Page Layout painting', () => {
  beforeEach(() => resetPaginationCache());

  it('lays paper over the desk before the cells are drawn', () => {
    const store = createSpreadsheetStore();
    seed(store, 200, 20);
    setWorkbookView(store, 'pageLayout');
    const state = store.getState();
    const { view, cols, rows } = viewFor(state);
    const { ctx, fills } = makeCtxSpy();

    paintPageLayoutBackground(pc(ctx), view, theme, paginationFor(state, 0), cols, rows);

    expect(fills[0]?.style).toBe(theme.pageBackdrop);
    expect(fills.slice(1).every((f) => f.style === theme.pagePaper)).toBe(true);
    expect(fills.length).toBeGreaterThan(1);
  });

  it('offers a header and footer slot per page and prompts only in the centre', () => {
    const store = createSpreadsheetStore();
    seed(store, 200, 20);
    setWorkbookView(store, 'pageLayout');
    const state = store.getState();
    const { view, cols, rows } = viewFor(state);
    const { ctx, texts } = makeCtxSpy();

    const bands = paintPageLayoutChrome(
      pc(ctx),
      view,
      theme,
      paginationFor(state, 0),
      getPageSetup(state, 0),
      { addHeader: en.pageView.addHeader, addFooter: en.pageView.addFooter },
      cols,
      rows,
    );

    const headers = bands.filter((h) => h.kind === 'header');
    expect(headers.length).toBeGreaterThan(0);
    expect(headers.length % 3).toBe(0);
    expect(headers.slice(0, 3).map((h) => h.slot)).toEqual(['left', 'center', 'right']);
    // One prompt per header band, in its centre slot only.
    expect(texts.filter((t) => t === en.pageView.addHeader)).toHaveLength(headers.length / 3);
  });

  it('paints the header text a page carries instead of the prompt', () => {
    const store = createSpreadsheetStore();
    seed(store, 200, 20);
    mutators.setPageSetup(store, 0, { headerCenter: 'Quarterly report' });
    setWorkbookView(store, 'pageLayout');
    const state = store.getState();
    const { view, cols, rows } = viewFor(state);
    const { ctx, texts } = makeCtxSpy();

    paintPageLayoutChrome(
      pc(ctx),
      view,
      theme,
      paginationFor(state, 0),
      getPageSetup(state, 0),
      { addHeader: en.pageView.addHeader, addFooter: en.pageView.addFooter },
      cols,
      rows,
    );

    expect(texts).toContain('Quarterly report');
    expect(texts).not.toContain(en.pageView.addHeader);
  });

  it('puts a margin grip on each ruler and measures it in the locale unit', () => {
    const store = createSpreadsheetStore();
    seed(store, 200, 20);
    setWorkbookView(store, 'pageLayout');
    const state = store.getState();
    const { view, cols, rows } = viewFor(state);
    const pagination = paginationFor(state, 0);
    const { ctx } = makeCtxSpy();

    const handles = paintPageRulers(pc(ctx), view, theme, pagination, 'cm', cols, rows);
    expect(handles.map((h) => h.side)).toContain('left');
    expect(handles.map((h) => h.side)).toContain('top');
    const left = handles.find((h) => h.side === 'left');
    // The paper edge is one left margin ahead of the printable origin.
    expect((left?.position ?? 0) - (left?.paperOrigin ?? 0)).toBeCloseTo(
      pagination.marginPx.left,
      5,
    );
  });
});

describe('Page Break Preview painting', () => {
  beforeEach(() => resetPaginationCache());

  it('numbers the printed pages and washes out the paper past them', () => {
    const store = createSpreadsheetStore();
    seed(store, 4, 3);
    setWorkbookView(store, 'pageBreakPreview');
    const state = store.getState();
    const { view, cols, rows } = viewFor(state);
    const { ctx, texts, fills } = makeCtxSpy();

    paintPageBreakPreview(
      pc(ctx),
      view,
      theme,
      paginationFor(state, 0),
      getPageSetup(state, 0),
      (page) => `Page ${page}`,
      cols,
      rows,
    );

    expect(texts).toContain('Page 1');
    expect(texts).not.toContain('Page 2');
    expect(fills.some((f) => f.style === theme.pageOutside)).toBe(true);
  });

  it('exposes the print-area frame as a drag handle', () => {
    const store = createSpreadsheetStore();
    seed(store, 4, 3);
    setWorkbookView(store, 'pageBreakPreview');
    const state = store.getState();
    const { view, cols, rows } = viewFor(state);
    const { ctx } = makeCtxSpy();

    const handles = paintPageBreakPreview(
      pc(ctx),
      view,
      theme,
      paginationFor(state, 0),
      getPageSetup(state, 0),
      (page) => `Page ${page}`,
      cols,
      rows,
    );
    const frame = handles.filter((h) => h.kind === 'printArea');
    expect(frame.map((h) => h.axis).sort()).toEqual(['col', 'row']);
    expect(frame.find((h) => h.axis === 'row')?.index).toBe(4);
    expect(frame.find((h) => h.axis === 'col')?.index).toBe(3);
  });

  it('draws a break line only where a page actually splits the printed range', () => {
    const store = createSpreadsheetStore();
    seed(store, 200, 3);
    mutators.setPageSetup(store, 0, { manualPageBreakRows: [20] });
    setWorkbookView(store, 'pageBreakPreview');
    const state = store.getState();
    const { view, cols, rows } = viewFor(state);
    const { ctx } = makeCtxSpy();

    const handles = paintPageBreakPreview(
      pc(ctx),
      view,
      theme,
      paginationFor(state, 0),
      getPageSetup(state, 0),
      (page) => `Page ${page}`,
      cols,
      rows,
    );

    const breaks = handles.filter((h) => h.kind === 'break');
    expect(breaks.some((h) => h.axis === 'row' && h.index === 20 && h.manual)).toBe(true);
    // The sheet is one column band wide, so no vertical split exists.
    expect(breaks.some((h) => h.axis === 'col')).toBe(false);
  });

  it('keeps native page watermarks on the continuous cell band', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    try {
      for (let row = 0; row < 3; row += 1) {
        wb.setNumber({ sheet: 0, row, col: 0 }, row + 1);
        store.setState((s) => {
          const cells = new Map(s.data.cells);
          cells.set(addrKey({ sheet: 0, row, col: 0 }), {
            value: { kind: 'number', value: row + 1 },
            formula: null,
          });
          return {
            ...s,
            data: { ...s.data, cells },
            viewport: { ...s.viewport, rowCount: 10, colCount: 5, widthPx: 800 },
          };
        });
      }
      mutators.setPageSetup(store, 0, { printArea: 'A1:B3' });
      syncPageSetupToEngine(wb, store, 0);
      setWorkbookView(store, 'pageBreakPreview');
      const state = store.getState();
      const { view, cols, rows } = viewFor(state);
      const { ctx, textRecords } = makeCtxSpy();

      paintPageBreakPreview(
        pc(ctx),
        view,
        theme,
        paginationFor(state, 0, { engine: wb }),
        getPageSetup(state, 0),
        (page) => `Page ${page}`,
        cols,
        rows,
      );

      const pageOne = textRecords.find((record) => record.text === 'Page 1');
      expect(pageOne).toBeDefined();
      // The two-row cell band is near the grid origin. A physical A4 paper
      // rect would put the old watermark hundreds of pixels below it.
      expect(pageOne?.y).toBeLessThan(100);
    } finally {
      wb.dispose();
    }
  });

  it('continues native blank tail paper by one physical page', async () => {
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
          return {
            ...s,
            data: { ...s.data, cells },
            viewport: { ...s.viewport, rowCount: 180, colCount: 1, widthPx: 800 },
          };
        });
      }
      mutators.setPageSetup(store, 0, {
        printArea: 'A1:A160',
        printTitleRows: '1:1',
        manualPageBreakRows: [60],
        centerVertically: true,
        scale: 0.5,
      });
      syncPageSetupToEngine(wb, store, 0);
      setWorkbookView(store, 'pageLayout');
      const state = store.getState();
      associatePaginationEngine(state, wb);
      const { view, cols, rows } = viewFor(state);
      const pagination = paginationFor(state, 0, { engine: wb });
      const { ctx, fills } = makeCtxSpy();

      paintPageLayoutBackground(pc(ctx), view, theme, pagination, cols, rows);

      const pageHeight = pageBoxPx(pagination, getPageSetup(state, 0)).height;
      const pageStarts = [
        ...new Set(
          fills
            .filter((fill) => fill.style === theme.pagePaper && fill.rect[2] > 500)
            .map((fill) => Number(fill.rect[1].toFixed(3))),
        ),
      ].sort((a, b) => a - b);
      expect(pageStarts.length).toBeGreaterThanOrEqual(4);
      const previousPageStart = pageStarts[2];
      const tailPageStart = pageStarts[3];
      expect(previousPageStart).toBeDefined();
      expect(tailPageStart).toBeDefined();
      if (previousPageStart === undefined || tailPageStart === undefined) return;
      expect(tailPageStart - previousPageStart).toBeCloseTo(pageHeight + PAGE_GUTTER_PX, 2);

      store.setState((s) => ({
        ...s,
        viewport: { ...s.viewport, rowStart: 160, rowCount: 20 },
      }));
      const tailState = store.getState();
      associatePaginationEngine(tailState, wb);
      const { view: tailView, cols: tailCols, rows: tailRows } = viewFor(tailState);
      const tailPagination = paginationFor(tailState, 0, { engine: wb });
      const { ctx: tailCtx, fills: tailFills } = makeCtxSpy();
      paintPageLayoutBackground(pc(tailCtx), tailView, theme, tailPagination, tailCols, tailRows);
      const lastNativePage = tailPagination.native?.pages.at(-1);
      const tailGap = tailView.layout.pageGapRows?.get(160);
      const printable = tailPagination.native?.printable;
      const titleOffset =
        lastNativePage && printable ? Math.max(0, lastNativePage.originYPt - printable.y) : 0;
      const centerOffset =
        lastNativePage && printable
          ? Math.max(0, (printable.height - titleOffset - lastNativePage.heightPt) / 2)
          : 0;
      const carriedOriginPt = (lastNativePage?.originYPt ?? 0) + centerOffset;
      const factor = 96 / 72 / tailPagination.scale;
      const expectedTailStart =
        gridOriginY(tailView.layout) + (tailGap ?? 0) - carriedOriginPt * factor;
      const tailPaper = tailFills.find(
        (fill) => fill.style === theme.pagePaper && fill.rect[2] > 500,
      );
      expect(tailPaper?.rect[1]).toBeCloseTo(expectedTailStart, 2);
    } finally {
      associatePaginationEngine(store.getState(), null);
      wb.dispose();
    }
  });
});
