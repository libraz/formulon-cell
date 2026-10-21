import { describe, expect, it } from 'vitest';
import { setOutlineSummaryPosition } from '../../../src/commands/outline.js';
import { type AxisLayout, layoutForView, type ViewState } from '../../../src/render/geometry.js';
import { paintOutlineGutters } from '../../../src/render/painters/controls.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import type { ResolvedTheme } from '../../../src/theme/resolve.js';

const theme = new Proxy({}, { get: () => '#000' }) as unknown as ResolvedTheme;

const ctx = {
  save() {},
  restore() {},
  beginPath() {},
  rect() {},
  clip() {},
  moveTo() {},
  lineTo() {},
  stroke() {},
  fillRect() {},
  strokeRect() {},
} as unknown as CanvasRenderingContext2D;

const axis = (count: number, size: number): AxisLayout => ({
  visible: Array.from({ length: count }, (_, i) => i),
  positionAt: new Map(Array.from({ length: count }, (_, i) => [i, i * size])),
  sizeAt: new Map(Array.from({ length: count }, (_, i) => [i, size])),
  frozenTotal: 0,
  gapAt: new Map(),
});

const paint = (summary: { below?: boolean; right?: boolean }) => {
  const store = createSpreadsheetStore();
  // Rows 3..5 and columns 3..5 form a group; row/column 2 and 6 are the neighbours.
  store.setState((s) => ({
    ...s,
    layout: {
      ...s.layout,
      outlineRows: new Map([3, 4, 5].map((r) => [r, 1])),
      outlineCols: new Map([3, 4, 5].map((c) => [c, 1])),
      outlineRowGutter: 14,
      outlineColGutter: 14,
    },
  }));
  setOutlineSummaryPosition(store, summary);
  const state = store.getState();
  const view = { ...state, layout: layoutForView(state) } as ViewState;
  const hits = paintOutlineGutters(ctx, view, theme, axis(10, 50), axis(10, 20), 800, 600);
  return {
    row: hits.find((h) => h.axis === 'row')?.rect,
    col: hits.find((h) => h.axis === 'col')?.rect,
  };
};

describe('outline summary position', () => {
  it('puts the toggles on the row below and the column to the right by default', () => {
    const { row, col } = paint({});
    const rowCenter = (row?.y ?? 0) + (row?.h ?? 0) / 2;
    const colCenter = (col?.x ?? 0) + (col?.w ?? 0) / 2;
    const base = paint({ below: false, right: false });
    const rowCenterAbove = (base.row?.y ?? 0) + (base.row?.h ?? 0) / 2;
    const colCenterLeft = (base.col?.x ?? 0) + (base.col?.w ?? 0) / 2;
    expect(rowCenter).toBeGreaterThan(rowCenterAbove);
    expect(colCenter).toBeGreaterThan(colCenterLeft);
    // One row (20px) / column (50px) above or left of the group start versus
    // one past the group end.
    expect(rowCenter - rowCenterAbove).toBe(20 * 4);
    expect(colCenter - colCenterLeft).toBe(50 * 4);
  });

  it('keeps the default for a sheet whose entry returns to below / right', () => {
    const store = createSpreadsheetStore();
    setOutlineSummaryPosition(store, { below: false });
    expect(store.getState().layout.outlineSummary?.get(0)).toEqual({ below: false, right: true });
    setOutlineSummaryPosition(store, { below: true });
    expect(store.getState().layout.outlineSummary?.has(0)).toBe(false);
  });

  it('refuses to change a protected sheet', () => {
    const store = createSpreadsheetStore();
    mutators.setSheetProtected(store, 0, true);
    setOutlineSummaryPosition(store, { below: false });
    expect(store.getState().layout.outlineSummary?.get(0)).toBeUndefined();
  });
});
