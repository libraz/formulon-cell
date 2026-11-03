import { describe, expect, it } from 'vitest';
import {
  buildColLayout,
  buildRowLayout,
  cellRectIn,
  layoutForView,
} from '../../../src/render/geometry.js';
import { paintBorders } from '../../../src/render/grid/borders.js';
import type { CellBorders } from '../../../src/store/store.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import type { ResolvedTheme } from '../../../src/theme/resolve.js';

const run = (borders: CellBorders, dpr: number) => {
  const store = createSpreadsheetStore();
  mutators.setCellFormat(store, { sheet: 0, row: 1, col: 1 }, { borders });
  store.setState((state) => ({
    ...state,
    viewport: { ...state.viewport, rowCount: 4, colCount: 4, widthPx: 400 },
  }));
  const view = { ...store.getState(), layout: layoutForView(store.getState()) };
  const cols = buildColLayout(view.layout, view.viewport);
  const rows = buildRowLayout(view.layout, view.viewport);
  const fills: Array<{ style: string; rect: [number, number, number, number]; alpha: number }> = [];
  const target: Record<string, unknown> = {
    fillStyle: '',
    globalAlpha: 1,
    fillRect: (x: number, y: number, w: number, h: number) =>
      fills.push({
        style: String(target.fillStyle),
        rect: [x, y, w, h],
        alpha: Number(target.globalAlpha),
      }),
  };
  const ctx = new Proxy(target, {
    get: (o, k) => (k in o ? o[k as string] : () => undefined),
    set: (o, k, v) => {
      o[k as string] = v;
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
  const theme = { bg: '#fefefe', cellAutoFg: '#111111' } as ResolvedTheme;
  paintBorders(ctx, dpr, view, theme, cols, rows);
  return { fills, rect: cellRectIn(view.layout, cols, rows, 1, 1) };
};

describe('paintBorders gridline gaps', () => {
  for (const dpr of [1, 2]) {
    it(`clears the whole medium band under a dashed bottom edge (dpr ${dpr})`, () => {
      const { fills, rect } = run({ bottom: { style: 'mediumDashed' } }, dpr);
      const bottom = Math.round((rect.y + rect.h) * dpr);
      const gaps = fills.filter((f) => f.style === '#fefefe');
      expect(gaps).toHaveLength(1);
      expect(gaps[0]?.rect[1]).toBe(bottom - 2 * dpr);
      expect(gaps[0]?.rect[3]).toBe(2 * dpr);
    });
  }

  it('clears one gridline width under a thin dashed edge', () => {
    const { fills, rect } = run({ bottom: { style: 'dashed' } }, 2);
    const gaps = fills.filter((f) => f.style === '#fefefe');
    expect(gaps[0]?.rect[1]).toBe(Math.round((rect.y + rect.h) * 2) - 2);
    expect(gaps[0]?.rect[3]).toBe(2);
  });

  it('clears the whole medium band left of a dashed right edge', () => {
    const { fills, rect } = run({ right: { style: 'mediumDashDot' } }, 2);
    const right = Math.round((rect.x + rect.w) * 2);
    const gaps = fills.filter((f) => f.style === '#fefefe');
    expect(gaps[0]?.rect[0]).toBe(right - 4);
    expect(gaps[0]?.rect[2]).toBe(4);
  });
});

describe('paintBorders slant dash-dot ends', () => {
  const inkOf = (fills: ReturnType<typeof run>['fills'], color: string) => {
    const ink: { x: number; y: number }[] = [];
    for (const f of fills) {
      if (f.style !== color || f.alpha <= 0) continue;
      const [x, y, w, h] = f.rect;
      for (let yy = y; yy < y + h; yy++)
        for (let xx = x; xx < x + w; xx++) ink.push({ x: xx, y: yy });
    }
    return ink;
  };

  it('ends a vertical edge on a slanted cut 3 device px below the cell bottom', () => {
    const { fills, rect } = run({ right: { style: 'slantDashDot', color: '#c00000' } }, 2);
    const bottom = Math.round((rect.y + rect.h) * 2);
    const left = Math.round((rect.x + rect.w) * 2) - 4;
    const ink = inkOf(fills, '#c00000');
    expect(ink.length).toBeGreaterThan(0);
    expect(ink.every(({ x, y }) => y + 2 * (x - left) < bottom + 3)).toBe(true);
    expect(ink.some(({ y }) => y >= bottom)).toBe(true);
  });

  it('runs a horizontal edge past a corner on a slanted cut, and ends a lone one flat', () => {
    const color = '#c00000';
    const slant = { style: 'slantDashDot', color } as const;
    const corner = run({ top: slant, right: slant }, 2);
    const right = Math.round((corner.rect.x + corner.rect.w) * 2);
    const top = Math.round(corner.rect.y * 2) - 4;
    const past = inkOf(corner.fills, color).filter(({ x }) => x >= right);
    expect(past.length).toBeGreaterThan(0);
    expect(past.every(({ x, y }) => x + 2 * (y - top) < right + 3)).toBe(true);

    const lone = run({ top: slant }, 2);
    const loneRight = Math.round((lone.rect.x + lone.rect.w) * 2);
    expect(inkOf(lone.fills, color).every(({ x }) => x < loneRight)).toBe(true);
  });

  it('leaves the corner band to the horizontal edge that owns it', () => {
    const color = '#c00000';
    const { fills, rect } = run(
      { top: { style: 'slantDashDot', color }, right: { style: 'slantDashDot', color } },
      2,
    );
    const right = Math.round((rect.x + rect.w) * 2);
    const top = Math.round(rect.y * 2);
    const corner = inkOf(fills, color).filter(({ y }) => y >= top - 4 && y < top);
    expect(corner.length).toBeGreaterThan(0);
    expect(corner.every(({ x, y }) => x + 2 * (y - (top - 4)) < right + 3)).toBe(true);
  });
});
