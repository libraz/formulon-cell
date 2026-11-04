import { describe, expect, it } from 'vitest';
import { buildColLayout, buildRowLayout, layoutForView } from '../../../../src/render/geometry.js';
import { paintHeaders } from '../../../../src/render/grid/headers.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

const theme = {
  bg: '#fff',
  bgHeader: '#e6e6e6',
  fg: '#201f1e',
  rule: '#d9d9d9',
  headerBg: '#faf9f8',
  headerRule: '#d9d9d9',
  headerRuleOuter: '#d9d9d9',
  headerFg: '#605e5c',
  headerFgActive: '#107c41',
  headerFgFiltered: '#1c00ff',
  headerFgActiveWeight: 600,
  headerActiveBar: '#107c41',
  headerStripOuter: '#faf9f8',
  headerStripInner: '#faf9f8',
  outlineGutterBg: '#faf9f8',
  accent: '#107c41',
  fontUi: 'Arial',
  textHeader: 11,
} as unknown as ResolvedTheme;

describe('paintHeaders draw state', () => {
  it('sets each label font once per change and strokes separators once per axis', () => {
    const state = createSpreadsheetStore().getState();
    state.viewport = {
      rowStart: 0,
      rowCount: 20,
      colStart: 0,
      colCount: 10,
      zoom: 1,
      widthPx: 700,
    };
    const layout = layoutForView(state);
    const cols = buildColLayout(layout, state.viewport);
    const rows = buildRowLayout(layout, state.viewport);
    let fontSets = 0;
    let strokes = 0;
    const store: Record<string, unknown> = {};
    const ctx = new Proxy(store, {
      get(t, k: string) {
        if (k in t) return t[k];
        if (k === 'stroke') return () => (strokes += 1);
        if (k === 'measureText') {
          return () => ({ width: 0, fontBoundingBoxAscent: 9, fontBoundingBoxDescent: 2 });
        }
        return () => ({ addColorStop() {} });
      },
      set(t, k: string, v) {
        if (k === 'font') fontSets += 1;
        t[k] = v;
        return true;
      },
    }) as unknown as CanvasRenderingContext2D;

    paintHeaders(
      { ctx, dpr: 1, cssWidth: 700, cssHeight: 500 },
      { ...state, layout },
      theme,
      cols,
      rows,
    );

    // Initial font, and at most one switch to the active weight and back per axis.
    expect(fontSets).toBeLessThanOrEqual(6);
    // Border rule, column separators, row separators, plus outline gutters.
    expect(strokes).toBeLessThanOrEqual(6);
  });

  it('paints the row numbers of the rows a filter lets through in the filtered colour', () => {
    const base = createSpreadsheetStore().getState();
    const range = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 1 };
    const state = {
      ...base,
      viewport: { rowStart: 0, rowCount: 8, colStart: 0, colCount: 4, zoom: 1, widthPx: 700 },
      ui: {
        ...base.ui,
        filterRanges: [range],
        filterCriteria: [{ range, byCol: 1, hiddenValues: ['Y'] }],
      },
    };
    const layout = layoutForView({
      ...state,
      layout: { ...base.layout, hiddenRows: new Set([4]) },
    });
    const cols = buildColLayout(layout, state.viewport);
    const rows = buildRowLayout(layout, state.viewport);
    const colourAt = new Map<string, string>();
    let fill = '';
    const ctx = new Proxy({} as Record<string, unknown>, {
      get(t, k: string) {
        if (k in t) return t[k];
        if (k === 'fillText') return (text: string) => colourAt.set(text, fill);
        if (k === 'measureText') {
          return () => ({ width: 0, fontBoundingBoxAscent: 9, fontBoundingBoxDescent: 2 });
        }
        return () => ({ addColorStop() {} });
      },
      set(t, k: string, v) {
        if (k === 'fillStyle' && typeof v === 'string') fill = v;
        t[k] = v;
        return true;
      },
    }) as unknown as CanvasRenderingContext2D;

    paintHeaders(
      { ctx, dpr: 1, cssWidth: 700, cssHeight: 500 },
      { ...state, layout },
      theme,
      cols,
      rows,
    );

    expect(colourAt.get('1')).not.toBe('#1c00ff');
    expect(colourAt.get('2')).toBe('#1c00ff');
    expect(colourAt.get('4')).toBe('#1c00ff');
    expect(colourAt.get('6')).not.toBe('#1c00ff');
  });
});
