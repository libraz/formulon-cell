import { beforeEach, describe, expect, it } from 'vitest';
import { paginationFor, resetPaginationCache } from '../../../../src/commands/pagination.js';
import { setWorkbookView } from '../../../../src/commands/view.js';
import { addrKey } from '../../../../src/engine/address.js';
import { en } from '../../../../src/i18n/strings/en.js';
import { buildColLayout, buildRowLayout, layoutForView } from '../../../../src/render/geometry.js';
import {
  getPageBandHits,
  getPageBreakHandles,
  getRulerHandles,
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

function makeCtxSpy(): {
  ctx: CanvasRenderingContext2D;
  fills: FillRecord[];
  texts: string[];
} {
  const fills: FillRecord[] = [];
  const texts: string[] = [];
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
    fillText(text: string): void {
      texts.push(text);
    },
  } as unknown as CanvasRenderingContext2D;
  return { ctx, fills, texts };
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

    const headers = getPageBandHits().filter((h) => h.kind === 'header');
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

    paintPageRulers(pc(ctx), view, theme, pagination, 'cm', cols, rows);

    const handles = getRulerHandles();
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

    const handles = getPageBreakHandles();
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

    const breaks = getPageBreakHandles().filter((h) => h.kind === 'break');
    expect(breaks.some((h) => h.axis === 'row' && h.index === 20 && h.manual)).toBe(true);
    // The sheet is one column band wide, so no vertical split exists.
    expect(breaks.some((h) => h.axis === 'col')).toBe(false);
  });
});
