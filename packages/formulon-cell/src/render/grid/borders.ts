import { addrKey } from '../../engine/address.js';
import type { CellBorders } from '../../store/store.js';
import type { ResolvedTheme } from '../../theme/resolve.js';
import {
  type AxisLayout,
  cellRectIn,
  gridOriginX,
  gridOriginY,
  rangeRects,
  type ViewState,
} from '../geometry.js';
import { paintCellDiagonals } from '../painters.js';
import {
  type BorderEdge,
  borderEdge,
  borderStrokes,
  dashCornerPixels,
  type EdgeRect,
  edgeRects,
  fillEdgeRect,
  isDashed,
  resolveBorderEdge,
  slantTailLines,
  strokeSpans,
} from './border-edges.js';
import { gridLinesShown } from './lines.js';
import { mergedDiagonalBorders, mergeRangeAt } from './merged-cells.js';
import { connectedRects } from './range-rects.js';

/** Borders a cell inherits when it has no direct format, e.g. from a row or
 *  column style. Only consulted for cells without their own `borders`. */
export type InheritedBorders = (row: number, col: number) => CellBorders | undefined;

type Side = 'top' | 'bottom' | 'left' | 'right';

/** Height of the clip below the column headers; any canvas is shorter. */
const CLIP_EXTENT = 1e7;

export function paintBorders(
  ctx: CanvasRenderingContext2D,
  dpr: number,
  state: ViewState,
  theme: ResolvedTheme,
  cols: AxisLayout,
  rows: AxisLayout,
  inherited?: InheritedBorders,
  coveredRules?: ReadonlySet<string>,
): void {
  const { layout, data, format } = state;
  if (format.formats.size === 0 && !inherited) return;
  const sheet = data.sheetIndex;
  const g = Math.max(1, Math.round(dpr));
  const rtl = layout.rtl === true;

  const sidesCache = new Map<string, CellBorders | null>();
  /** Sides of a cell after dropping the ones a merge buries inside itself. */
  const sidesAt = (row: number, col: number): CellBorders | null => {
    if (row < 0 || col < 0) return null;
    const key = addrKey({ sheet, row, col });
    const cached = sidesCache.get(key);
    if (cached !== undefined) return cached;
    let borders = format.formats.get(key)?.borders ?? inherited?.(row, col) ?? null;
    const merge = borders ? mergeRangeAt(state, { sheet, row, col }) : null;
    if (borders && merge) {
      borders = { ...borders };
      if (row > merge.r0) borders.top = undefined;
      if (row < merge.r1) borders.bottom = undefined;
      if (rtl ? col < merge.c1 : col > merge.c0) borders.left = undefined;
      if (rtl ? col > merge.c0 : col < merge.c1) borders.right = undefined;
    }
    sidesCache.set(key, borders);
    return borders;
  };

  /** Overflowing text hides the vertical rule between two columns it spans. */
  const covered = (row: number, a: number, b: number): boolean =>
    coveredRules?.has(`${row}:${Math.min(a, b)}`) === true;

  /** The edge shown on `side` of a cell once its neighbour's opposite side
   *  is taken into account. Screen-left in RTL is the next column. */
  const edgeAt = (row: number, col: number, side: Side): BorderEdge | null => {
    const own = borderEdge(sidesAt(row, col)?.[side]);
    switch (side) {
      case 'top':
        return resolveBorderEdge(borderEdge(sidesAt(row - 1, col)?.bottom), own);
      case 'bottom':
        return resolveBorderEdge(own, borderEdge(sidesAt(row + 1, col)?.top));
      case 'left': {
        if (covered(row, col, rtl ? col + 1 : col - 1)) return null;
        const neighbour = borderEdge(sidesAt(row, rtl ? col + 1 : col - 1)?.right);
        return rtl ? resolveBorderEdge(own, neighbour) : resolveBorderEdge(neighbour, own);
      }
      case 'right': {
        if (covered(row, col, rtl ? col - 1 : col + 1)) return null;
        const neighbour = borderEdge(sidesAt(row, rtl ? col - 1 : col + 1)?.left);
        return rtl ? resolveBorderEdge(neighbour, own) : resolveBorderEdge(own, neighbour);
      }
    }
  };

  const fillAt = (row: number, col: number): string | undefined =>
    row < 0 || col < 0 ? undefined : format.formats.get(addrKey({ sheet, row, col }))?.fill;
  // Edges are laid out in logical order: in RTL the x axis is negated, so the
  // leading column is always `col - 1` and every rect is mirrored back below.
  const leadEdge = (row: number, col: number): BorderEdge | null =>
    edgeAt(row, col, rtl ? 'right' : 'left');
  const trailEdge = (row: number, col: number): BorderEdge | null =>
    edgeAt(row, col, rtl ? 'left' : 'right');
  const automatic = theme.cellAutoFg;

  ctx.save();
  const ox = gridOriginX(layout);
  const oy = gridOriginY(layout);
  ctx.beginPath();
  ctx.rect(rtl ? 0 : ox, oy, layout.viewWidth - ox, CLIP_EXTENT);
  ctx.clip();
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  const origin = {
    x: Math.round((rtl ? layout.viewWidth - ox : ox) * dpr),
    y: Math.round(oy * dpr),
  };
  const toScreen = (rect: EdgeRect): EdgeRect => (rtl ? { ...rect, x: -(rect.x + rect.w) } : rect);

  // Gridline gaps are cleared for every edge before any stroke is drawn, so
  // one edge's clearing cannot erase a neighbouring edge's stroke.
  const strokes: (() => void)[] = [];
  // Two cells share each edge; draw it once so translucent pixels of a slant
  // cut are not composited twice.
  const drawn = new Set<string>();
  const drawEdge = (
    key: string,
    continued: boolean,
    edge: BorderEdge,
    horizontal: boolean,
    e: number,
    a0: number,
    a1: number,
    start: { before: BorderEdge | null; after: BorderEdge | null },
    end: { before: BorderEdge | null; after: BorderEdge | null },
    under: string | undefined,
    tail = false,
  ): void => {
    if (drawn.has(key)) return;
    drawn.add(key);
    if (gridLinesShown(state)) {
      // A medium band is wider than the gridline, so the gridlines crossing
      // its upper half go too.
      const band = -Math.min(-g, ...borderStrokes(edge.style, g).map((s) => s[0]));
      const gap = toScreen(
        horizontal
          ? { x: a0 - g, y: e - band, w: a1 - a0 + g, h: band }
          : { x: e - band, y: a0 - g, w: band, h: a1 - a0 + g },
      );
      ctx.fillStyle = under ?? theme.bg;
      ctx.fillRect(gap.x, gap.y, gap.w, gap.h);
    }
    let spans = strokeSpans(edge, g, a0, a1, start, end, continued);
    // At a slant dash-dot corner the horizontal edges run through it; a
    // vertical edge starts flush with its boundary.
    if (!horizontal && edge.style === 'slantDashDot' && (start.before || start.after)) {
      spans = spans.map(([from, to]) => [Math.max(from, a0), to] as const);
    }
    const color = edge.color || automatic;
    strokes.push(() => {
      for (const rect of edgeRects(edge, horizontal, e, spans, g)) {
        const screen = toScreen(rect);
        if (tail && !rtl && edge.style === 'slantDashDot') {
          for (const line of slantTailLines(screen, horizontal, g)) {
            fillEdgeRect(ctx, edge, horizontal, line, color, g, origin, screen);
          }
        } else {
          fillEdgeRect(ctx, edge, horizontal, screen, color, g, origin);
        }
      }
    });
    if (!horizontal || !isDashed(edge.style)) return;
    // Two dashed lines stop short of each other; the corner square between
    // them is inked where both patterns are on.
    const [y0, y1] = [e + (borderStrokes(edge.style, g)[0]?.[0] ?? 0), e];
    for (const [at, perps] of [
      [a0, start],
      [a1, end],
    ] as const) {
      const perp = [perps.before, perps.after].find((p) => p && isDashed(p.style));
      if (!perp) continue;
      const [s0, s1] = borderStrokes(perp.style, g)[0] ?? [0, 0];
      const corner = toScreen({ x: at + s0, y: y0, w: s1 - s0, h: y1 - y0 });
      strokes.push(() => {
        ctx.fillStyle = color;
        const pixels = dashCornerPixels(
          edge.style,
          perp.style,
          g,
          corner.x - origin.x,
          corner.x + corner.w - origin.x,
          corner.y - origin.y,
          corner.y + corner.h - origin.y,
        );
        for (const p of pixels) ctx.fillRect(p.x + origin.x, p.y + origin.y, 1, 1);
      });
    }
  };

  const px = (v: number): number => Math.round(v * dpr);
  /** Whether `index - 1` is painted right before `index` on the axis, so a
   *  line through both is already drawn up to the shared boundary. */
  const follows = (axis: AxisLayout, index: number): boolean => {
    const prev = axis.positionAt.get(index - 1);
    const size = axis.sizeAt.get(index - 1);
    return prev !== undefined && size !== undefined && prev + size === axis.positionAt.get(index);
  };
  for (const r of rows.visible) {
    for (const c of cols.visible) {
      const top = edgeAt(r, c, 'top');
      const bottom = edgeAt(r, c, 'bottom');
      const lead = leadEdge(r, c);
      const trail = trailEdge(r, c);
      if (!top && !bottom && !lead && !trail) continue;
      const rect = cellRectIn(layout, cols, rows, r, c);
      const L = rtl ? -px(rect.x + rect.w) : px(rect.x);
      const R = rtl ? -px(rect.x) : px(rect.x + rect.w);
      const T = px(rect.y);
      const B = px(rect.y + rect.h);
      if (top) {
        const start = { before: leadEdge(r - 1, c), after: lead };
        const end = { before: trailEdge(r - 1, c), after: trail };
        const continued = follows(cols, c) && edgeAt(r, c - 1, 'top') !== null;
        const under = fillAt(r, c) ?? fillAt(r - 1, c);
        const tail = !!(end.before || end.after);
        drawEdge(`h${r}:${c}`, continued, top, true, T, L, R, start, end, under, tail);
      }
      if (bottom) {
        const start = { before: lead, after: leadEdge(r + 1, c) };
        const end = { before: trail, after: trailEdge(r + 1, c) };
        const continued = follows(cols, c) && edgeAt(r, c - 1, 'bottom') !== null;
        const under = fillAt(r + 1, c) ?? fillAt(r, c);
        const tail = !!(end.before || end.after);
        drawEdge(`h${r + 1}:${c}`, continued, bottom, true, B, L, R, start, end, under, tail);
      }
      if (lead) {
        const start = { before: edgeAt(r, c - 1, 'top'), after: top };
        const end = { before: edgeAt(r, c - 1, 'bottom'), after: bottom };
        const continued = follows(rows, r) && leadEdge(r - 1, c) !== null;
        const under = fillAt(r, c) ?? fillAt(r, c - 1);
        const tail = !(follows(rows, r + 1) && leadEdge(r + 1, c) !== null);
        drawEdge(`v${r}:${c}`, continued, lead, false, L, T, B, start, end, under, tail);
      }
      if (trail) {
        const start = { before: top, after: edgeAt(r, c + 1, 'top') };
        const end = { before: bottom, after: edgeAt(r, c + 1, 'bottom') };
        const continued = follows(rows, r) && trailEdge(r - 1, c) !== null;
        const under = fillAt(r, c + 1) ?? fillAt(r, c);
        const tail = !(follows(rows, r + 1) && trailEdge(r + 1, c) !== null);
        drawEdge(`v${r}:${c + 1}`, continued, trail, false, R, T, B, start, end, under, tail);
      }
    }
  }
  for (const stroke of strokes) stroke();
  ctx.restore();

  // Diagonals belong to the cell surface, not to an edge.
  for (const r of rows.visible) {
    for (const c of cols.visible) {
      const key = addrKey({ sheet, row: r, col: c });
      const borders = format.formats.get(key)?.borders ?? inherited?.(r, c);
      if (!borders?.diagonalDown && !borders?.diagonalUp) continue;
      // A uniform diagonal on a merge is painted once below.
      if (mergeRangeAt(state, { sheet, row: r, col: c })) continue;
      paintCellDiagonals({
        ctx,
        theme,
        dpr,
        bounds: cellRectIn(layout, cols, rows, r, c),
        value: { kind: 'blank' },
        formula: null,
        isActive: false,
        isInRange: false,
        format: { borders },
      });
    }
  }

  // A uniform diagonal border survives a merge in Excel, but is rendered as
  // one line over the visible merged rectangle. A per-cell pass would draw
  // one diagonal per body cell and leave crossings in the merged surface;
  // rangeRects keeps the same freeze/RTL projection as the selection path.
  for (const merge of state.merges.byAnchor.values()) {
    if (merge.sheet !== sheet) continue;
    const rects = rangeRects(layout, state.viewport, merge);
    const clips = connectedRects(rects);
    if (clips.length === 0) continue;
    const borders = mergedDiagonalBorders(format.formats, merge);
    if (!borders) continue;
    const globalBounds = {
      x: Math.min(...rects.map((rect) => rect.x)),
      y: Math.min(...rects.map((rect) => rect.y)),
      w:
        Math.max(...rects.map((rect) => rect.x + rect.w)) -
        Math.min(...rects.map((rect) => rect.x)),
      h:
        Math.max(...rects.map((rect) => rect.y + rect.h)) -
        Math.min(...rects.map((rect) => rect.y)),
    };
    for (const clip of clips) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(clip.x, clip.y, clip.w, clip.h);
      ctx.clip();
      paintCellDiagonals({
        ctx,
        theme,
        dpr,
        bounds: globalBounds,
        value: { kind: 'blank' },
        formula: null,
        isActive: false,
        isInRange: false,
        format: { borders },
      });
      ctx.restore();
    }
  }
}
