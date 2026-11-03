import { describe, expect, it } from 'vitest';
import { type AxisLayout, layoutForView, type ViewState } from '../../../src/render/geometry.js';
import { paintOutlineGutters } from '../../../src/render/painters/controls.js';
import { outlineGutterWidth } from '../../../src/store/outline-summary.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import type { ResolvedTheme } from '../../../src/theme/resolve.js';

// Excel for Mac at DPR 2: 24 device px per level lane (12 CSS px), a 4 px spine
// (2 CSS px) 6 px inside the lane, 24x24 toggles (12x12 CSS) filling the lane
// and centred on the summary row, 4x4 detail dots (2x2 CSS) in the next lane,
// one more lane for the level buttons, a 4 device px (2 CSS) outer pad and a
// 6 device px (3 CSS) margin.
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
    roundRect() {},
    fill() {},
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
  it('reserves a lane per level, one for the level buttons, a 2 px pad and 3 px', () => {
    const levels = new Map([
      [2, 1],
      [3, 2],
    ]);
    expect(outlineGutterWidth(levels, { screenScale: 0.75 })).toBe(41);
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
    expect(level1?.rect).toEqual({ x: 2, y: oy + 8 * 24 + 6, w: 12, h: 12 });
    expect(level2?.rect).toEqual({ x: 14, y: oy + 5 * 24 + 6, w: 12, h: 12 });
    // The spine is 2 px wide, 3 px inside its lane.
    const spines = strokes.filter((s) => s.width === 2 && s.from[0] === s.to[0]);
    expect(spines.map((s) => s.from[0])).toEqual([6, 18]);
    // Detail dots: level-1-only rows in lane 2, level-2 rows in lane 3.
    const dots = fills.filter(([, , w, h]) => w === 2 && h === 2);
    expect(dots).toEqual(
      expect.arrayContaining([
        [17, oy + 2 * 24 + 11, 2, 2],
        [29, oy + 3 * 24 + 11, 2, 2],
        [29, oy + 4 * 24 + 11, 2, 2],
        [17, oy + 7 * 24 + 11, 2, 2],
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
      ['row', 1, { x: 2, y: rowTop, w: 12, h: 12 }],
      ['row', 2, { x: 14, y: rowTop, w: 12, h: 12 }],
      ['row', 3, { x: 26, y: rowTop, w: 12, h: 12 }],
      ['col', 1, { x: colLeft, y: 2, w: 12, h: 12 }],
      ['col', 2, { x: colLeft, y: 14, w: 12, h: 12 }],
    ]);
    expect(texts.map((t) => t.text)).toEqual(['1', '2', '3', '1', '2']);
  });

  it('runs each Mac spine on to its toggle', () => {
    // Excel (Mac, DPR 2): the level-1 bracket of rows 3:8 runs unbroken from
    // y 98 to the toggle border at 396, past the group's end at 384.
    const view = outlinedView(0.75);
    const { ctx, strokes } = recorder();
    const hits = paintOutlineGutters(ctx, view, theme, axis(10, 50), axis(12, 24), 800, 600);
    const rowToggles = hits.filter((h) => h.kind === 'toggle' && h.axis === 'row');
    const rowSpines = strokes.filter((s) => s.width === 2 && s.from[0] === s.to[0]);
    expect(rowSpines.map((s) => s.to[1])).toEqual(rowToggles.map((h) => h.rect.y));
    const colToggle = hits.find((h) => h.kind === 'toggle' && h.axis === 'col');
    const colSpine = strokes.find((s) => s.width === 2 && s.from[1] === s.to[1]);
    expect(colSpine?.to[0]).toBe(colToggle?.rect.x);
  });

  it('paints the gutters and level buttons in the outline tokens', () => {
    // Excel (Mac light, DPR 2): gutter 230; each 24 device px button has a
    // 1 px 198 border, a 239 row under it, a face shading 251 -> 234 with a 234 right bevel, the
    // 102 digit, and a 1 px 247 shadow on its right and bottom.
    const view = outlinedView(0.75);
    const tokens: Record<string, string> = {
      outlineGutterBg: '#e6e6e6',
      outlineButtonFaceTop: '#fbfbfb',
      outlineButtonFaceBottom: '#eaeaee',
      outlineButtonBorder: '#c6c6c8',
      outlineButtonFg: '#666666',
      outlineButtonShadow: '#f7f7f7',
      outlineButtonHighlight: '#efefef',
    };
    const painted = new Proxy(tokens, {
      get: (t, k: string) => t[k] ?? '#000',
    }) as unknown as ResolvedTheme;
    const fills: Array<{ style: unknown; rect: [number, number, number, number] }> = [];
    const texts: Array<{ text: string; style: unknown }> = [];
    const gradients: Array<{ from: [number, number]; to: [number, number]; stops: string[] }> = [];
    const ctx = {
      fillStyle: '' as unknown,
      strokeStyle: '' as unknown,
      lineWidth: 1,
      font: '',
      textAlign: 'left',
      textBaseline: 'alphabetic',
      save() {},
      restore() {},
      beginPath() {},
      rect() {},
      clip() {},
      moveTo() {},
      lineTo() {},
      stroke() {},
      strokeRect() {},
      roundRect() {},
      fill() {},
      fillRect(this: { fillStyle: unknown }, x: number, y: number, w: number, h: number) {
        fills.push({ style: this.fillStyle, rect: [x, y, w, h] });
      },
      fillText(this: { fillStyle: unknown }, text: string) {
        texts.push({ text, style: this.fillStyle });
      },
      createLinearGradient(x0: number, y0: number, x1: number, y1: number) {
        const g = {
          from: [x0, y0] as [number, number],
          to: [x1, y1] as [number, number],
          stops: [] as string[],
        };
        gradients.push(g);
        return { addColorStop: (_: number, c: string) => g.stops.push(c) };
      },
    } as unknown as CanvasRenderingContext2D;

    const hits = paintOutlineGutters(ctx, view, painted, axis(10, 50), axis(12, 24), 800, 600);

    const { outlineRowGutter, outlineColGutter } = view.layout;
    expect(fills).toEqual(
      expect.arrayContaining([
        { style: '#e6e6e6', rect: [0, 0, outlineRowGutter, 600] },
        { style: '#e6e6e6', rect: [0, 0, 800, outlineColGutter] },
      ]),
    );
    const button = hits.find((h) => h.kind === 'level' && h.axis === 'row' && h.level === 1);
    const { x, y } = button?.rect ?? { x: 0, y: 0 };
    const at = (rx: number, ry: number, w: number, h: number) =>
      fills.find((f) => f.rect.join() === [rx, ry, w, h].join())?.style;
    expect(at(x, y, 11.5, 11.5)).toBe('#c6c6c8');
    expect(gradients).toContainEqual({
      from: [0, y + 1],
      to: [0, y + 11],
      stops: ['#fbfbfb', '#eaeaee'],
    });
    expect(at(x + 10.5, y + 0.5, 0.5, 10.5)).toBe('#eaeaee');
    expect(at(x + 0.5, y + 0.5, 10.5, 0.5)).toBe('#efefef');
    expect(at(x + 11.5, y, 0.5, 12)).toBe('#f7f7f7');
    expect(at(x, y + 11.5, 12, 0.5)).toBe('#f7f7f7');
    expect(texts.filter((t) => t.style === '#666666').map((t) => t.text)).toEqual([
      '1',
      '2',
      '3',
      '1',
      '2',
    ]);
  });

  it('keeps the level buttons off the default profile', () => {
    const view = outlinedView(1);
    const { ctx, texts } = recorder();
    const hits = paintOutlineGutters(ctx, view, theme, axis(10, 50), axis(12, 24), 800, 600);
    expect(hits.some((h) => h.kind === 'level')).toBe(false);
    expect(texts).toEqual([]);
  });
});

describe('outline brackets and toggles on the Mac point model', () => {
  // Excel (Mac light, DPR 2): spines, caps and dots 153; toggle border 138 on a
  // rounded 24 px square; minus glyph 51.
  const tokens: Record<string, string> = {
    outlineRule: '#999999',
    outlineToggleBorder: '#8a8a8a',
    outlineToggleFg: '#333333',
    bg: '#ffffff',
  };
  const painted = new Proxy(tokens, {
    get: (t, k: string) => t[k] ?? '#000',
  }) as unknown as ResolvedTheme;

  const styled = () => {
    const fills: Array<{ style: unknown; rect: [number, number, number, number] }> = [];
    const strokes: Array<{
      style: unknown;
      width: number;
      from: [number, number];
      to: [number, number];
    }> = [];
    const rounds: Array<{ style: unknown; rect: [number, number, number, number]; r: number }> = [];
    const texts: Array<{ text: string; x: number }> = [];
    let from: [number, number] = [0, 0];
    let path: Array<{ rect: [number, number, number, number]; r: number }> = [];
    const segs: Array<{ width: number; from: [number, number]; to: [number, number] }> = [];
    const ctx = {
      fillStyle: '' as unknown,
      strokeStyle: '' as unknown,
      lineWidth: 1,
      font: '',
      textAlign: 'left',
      textBaseline: 'alphabetic',
      save() {},
      restore() {},
      beginPath() {
        path = [];
        segs.length = 0;
      },
      rect() {},
      clip() {},
      moveTo(x: number, y: number) {
        from = [x, y];
      },
      lineTo(this: { lineWidth: number }, x: number, y: number) {
        segs.push({ width: this.lineWidth, from, to: [x, y] });
      },
      stroke(this: { strokeStyle: unknown }) {
        for (const s of segs) strokes.push({ style: this.strokeStyle, ...s });
      },
      roundRect(x: number, y: number, w: number, h: number, r: number) {
        path.push({ rect: [x, y, w, h], r });
      },
      fill(this: { fillStyle: unknown }) {
        for (const p of path) rounds.push({ style: this.fillStyle, ...p });
      },
      strokeRect() {},
      fillRect(this: { fillStyle: unknown }, x: number, y: number, w: number, h: number) {
        fills.push({ style: this.fillStyle, rect: [x, y, w, h] });
      },
      fillText(text: string, x: number) {
        texts.push({ text, x });
      },
      createLinearGradient() {
        return { addColorStop() {} };
      },
    } as unknown as CanvasRenderingContext2D;
    return { ctx, fills, strokes, rounds, texts };
  };

  const view = () => {
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
        screenScale: 0.75,
        outlineRows: rowLevels,
        outlineCols: colLevels,
        outlineRowGutter: outlineGutterWidth(rowLevels, { screenScale: 0.75 }),
        outlineColGutter: outlineGutterWidth(colLevels, { screenScale: 0.75 }),
      },
    }));
    const state = store.getState();
    return { ...state, layout: layoutForView(state) } as ViewState;
  };

  it('caps each bracket with a 6x2 stub at its open end', () => {
    // Excel: level-1 spine x -118..-115 from y 98, cap 12x4 at x -118..-107,
    // y 98..101; the column bracket caps 4x12 down from its start.
    const v = view();
    const { ctx, fills, strokes } = styled();
    paintOutlineGutters(ctx, v, painted, axis(10, 50), axis(12, 24), 800, 600);
    const oy = v.layout.outlineColGutter + v.layout.headerRowHeight;
    const ox = v.layout.outlineRowGutter + v.layout.headerColWidth;
    const caps = fills.filter((f) => f.style === '#999999' && f.rect[2] * f.rect[3] === 12);
    const rowSpines = strokes.filter((s) => s.width === 2 && s.from[0] === s.to[0]);
    const colSpine = strokes.find((s) => s.width === 2 && s.from[1] === s.to[1]);
    expect(rowSpines.map((s) => s.from)).toEqual([
      [6, oy + 2 * 24 + 1],
      [18, oy + 3 * 24 + 1],
    ]);
    expect(caps.map((c) => c.rect)).toEqual(
      expect.arrayContaining([
        [5, oy + 2 * 24 + 1, 6, 2],
        [17, oy + 3 * 24 + 1, 6, 2],
        [ox + 2 * 50 + 1, (colSpine?.from[1] ?? 0) - 1, 2, 6],
      ]),
    );
  });

  it('leaves no detail dot under a toggle', () => {
    // Excel: row 6 holds the level-2 toggle; its minus runs unbroken where the
    // row's lane-2 dot would sit.
    const v = view();
    const { ctx, fills } = styled();
    const hits = paintOutlineGutters(ctx, v, painted, axis(10, 50), axis(12, 24), 800, 600);
    const toggles = hits.filter((h) => h.kind === 'toggle').map((h) => h.rect);
    const dots = fills.filter((f) => f.rect[2] === 2 && f.rect[3] === 2);
    expect(dots.length).toBeGreaterThan(0);
    const inside = dots.filter(({ rect: [x, y] }) =>
      toggles.some((t) => x >= t.x && x < t.x + t.w && y >= t.y && y < t.y + t.h),
    );
    expect(inside).toEqual([]);
  });

  it('paints brackets, dots and toggles in the outline tokens', () => {
    const v = view();
    const { ctx, fills, strokes, rounds } = styled();
    const hits = paintOutlineGutters(ctx, v, painted, axis(10, 50), axis(12, 24), 800, 600);
    const spines = strokes.filter((s) => s.width === 2);
    expect(spines.length).toBe(3);
    expect(spines.every((s) => s.style === '#999999')).toBe(true);
    const dots = fills.filter((f) => f.rect[2] === 2 && f.rect[3] === 2);
    expect(dots.every((d) => d.style === '#999999')).toBe(true);
    // Rounded toggle: a radius-2.5 border square, a radius-1.5 face 1 px inside.
    const t = hits.find((h) => h.kind === 'toggle' && h.axis === 'row' && h.level === 2)?.rect;
    if (!t) throw new Error('no toggle');
    expect(rounds).toEqual(
      expect.arrayContaining([
        { style: '#8a8a8a', rect: [t.x, t.y, 12, 12], r: 2.5 },
        { style: '#ffffff', rect: [t.x + 1, t.y + 1, 10, 10], r: 1.5 },
      ]),
    );
    expect(fills).toContainEqual({ style: '#333333', rect: [t.x + 3, t.y + 5.5, 6, 1] });
  });

  it('centres level digits on the face inside the bevel', () => {
    // Excel: the '1' stem sits 1 device px left of the button's half width.
    const v = view();
    const { ctx, texts } = styled();
    const hits = paintOutlineGutters(ctx, v, painted, axis(10, 50), axis(12, 24), 800, 600);
    const buttons = hits.filter((h) => h.kind === 'level');
    expect(texts.map((t) => t.x)).toEqual(buttons.map((b) => b.rect.x + 5.5));
  });
});
