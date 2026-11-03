import { describe, expect, it } from 'vitest';
import { type AxisLayout, layoutForView, type ViewState } from '../../../src/render/geometry.js';
import { paintOutlineGutters } from '../../../src/render/painters/controls.js';
import { outlineGutterWidth } from '../../../src/store/outline-summary.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import type { ResolvedTheme } from '../../../src/theme/resolve.js';

// Excel for Mac at DPR 2: 24 device px per level lane (12 CSS px), a 4 px spine
// (2 CSS px) 6 px inside the lane, 24x24 toggles (12x12 CSS) filling the lane
// and centred on the summary row, 4x4 detail dots (2x2 CSS) in the next lane,
// and one more lane plus 6 device px (3 CSS) for the level buttons.
const theme = new Proxy({}, { get: () => '#000' }) as unknown as ResolvedTheme;

const recorder = () => {
  const fills: [number, number, number, number][] = [];
  const texts: { text: string; x: number; y: number }[] = [];
  const strokes: { width: number; from: [number, number]; to: [number, number] }[] = [];
  let from: [number, number] = [0, 0];
  let width = 1;
  const ctx = {
    save() {},
    restore() {},
    beginPath() {},
    rect() {},
    clip() {},
    moveTo(x: number, y: number) {
      from = [x, y];
    },
    lineTo(x: number, y: number) {
      strokes.push({ width, from, to: [x, y] });
    },
    stroke() {},
    fillRect(x: number, y: number, w: number, h: number) {
      fills.push([x, y, w, h]);
    },
    strokeRect() {},
    fillText(text: string, x: number, y: number) {
      texts.push({ text, x, y });
    },
    set lineWidth(v: number) {
      width = v;
    },
    get lineWidth() {
      return width;
    },
  } as unknown as CanvasRenderingContext2D;
  return { ctx, fills, strokes, texts };
};

const axis = (count: number, size: number): AxisLayout => ({
  visible: Array.from({ length: count }, (_, i) => i),
  positionAt: new Map(Array.from({ length: count }, (_, i) => [i, i * size])),
  sizeAt: new Map(Array.from({ length: count }, (_, i) => [i, size])),
  frozenTotal: 0,
  gapAt: new Map(),
});

describe('outline gutter on the Mac point model', () => {
  it('reserves a lane per level, one for the level buttons, and 3 px', () => {
    const levels = new Map([
      [2, 1],
      [3, 2],
    ]);
    expect(outlineGutterWidth(levels, { screenScale: 0.75 })).toBe(39);
    expect(outlineGutterWidth(levels, {})).toBe(28);
  });

  it('draws Excel-sized spines, toggles and detail dots', () => {
    // Rows 2..7 at level 1 with rows 3..4 at level 2, 24 px rows.
    const levels = new Map([
      [2, 1],
      [3, 2],
      [4, 2],
      [5, 1],
      [6, 1],
      [7, 1],
    ]);
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        screenScale: 0.75,
        outlineRows: levels,
        outlineRowGutter: outlineGutterWidth(levels, { screenScale: 0.75 }),
      },
    }));
    const state = store.getState();
    const view = { ...state, layout: layoutForView(state) } as ViewState;
    const { ctx, fills, strokes } = recorder();
    const hits = paintOutlineGutters(ctx, view, theme, axis(10, 50), axis(12, 24), 800, 600);
    const oy = view.layout.headerRowHeight;

    const level1 = hits.find((h) => h.level === 1);
    const level2 = hits.find((h) => h.level === 2);
    // Toggles fill their lane and centre on the row after the group.
    expect(level1?.rect).toEqual({ x: 0, y: oy + 8 * 24 + 6, w: 12, h: 12 });
    expect(level2?.rect).toEqual({ x: 12, y: oy + 5 * 24 + 6, w: 12, h: 12 });
    // The spine is 2 px wide, 3 px inside its lane.
    const spines = strokes.filter((s) => s.width === 2 && s.from[0] === s.to[0]);
    expect(spines.map((s) => s.from[0])).toEqual([4, 16]);
    // Detail dots: level-1-only rows in lane 2, level-2 rows in lane 3.
    const dots = fills.filter(([, , w, h]) => w === 2 && h === 2);
    expect(dots).toEqual(
      expect.arrayContaining([
        [15, oy + 2 * 24 + 11, 2, 2],
        [27, oy + 3 * 24 + 11, 2, 2],
        [27, oy + 4 * 24 + 11, 2, 2],
        [15, oy + 7 * 24 + 11, 2, 2],
      ]),
    );
  });

  // Excel for Mac at 2x halved: 12 px level buttons. The row buttons run along
  // the row lanes, 5 px below the top of the column-header band; the column
  // buttons stack down the column lanes, 5 px inside the row-header band.
  const outlinedView = (screenScale: number) => {
    const rowLevels = new Map([
      [2, 1],
      [3, 2],
      [4, 2],
      [5, 1],
      [6, 1],
      [7, 1],
    ]);
    const colLevels = new Map([
      [2, 1],
      [3, 1],
      [4, 1],
    ]);
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        screenScale,
        outlineRows: rowLevels,
        outlineCols: colLevels,
        outlineRowGutter: outlineGutterWidth(rowLevels, { screenScale }),
        outlineColGutter: outlineGutterWidth(colLevels, { screenScale }),
      },
    }));
    const state = store.getState();
    return { ...state, layout: layoutForView(state) } as ViewState;
  };

  it('paints the level buttons in the gutter corner', () => {
    const view = outlinedView(0.75);
    const { ctx, texts } = recorder();
    const hits = paintOutlineGutters(ctx, view, theme, axis(10, 50), axis(12, 24), 800, 600);
    const levelHits = hits.filter((h) => h.kind === 'level');
    const rowTop = view.layout.outlineColGutter + 5;
    const colLeft = view.layout.outlineRowGutter + 5;
    expect(levelHits.map((h) => [h.axis, h.level, h.rect])).toEqual([
      ['row', 1, { x: 0, y: rowTop, w: 12, h: 12 }],
      ['row', 2, { x: 12, y: rowTop, w: 12, h: 12 }],
      ['row', 3, { x: 24, y: rowTop, w: 12, h: 12 }],
      ['col', 1, { x: colLeft, y: 0, w: 12, h: 12 }],
      ['col', 2, { x: colLeft, y: 12, w: 12, h: 12 }],
    ]);
    expect(texts.map((t) => t.text)).toEqual(['1', '2', '3', '1', '2']);
  });

  it('keeps the level buttons off the default profile', () => {
    const view = outlinedView(1);
    const { ctx, texts } = recorder();
    const hits = paintOutlineGutters(ctx, view, theme, axis(10, 50), axis(12, 24), 800, 600);
    expect(hits.some((h) => h.kind === 'level')).toBe(false);
    expect(texts).toEqual([]);
  });
});
