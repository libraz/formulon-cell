import { SVG_NS } from '../svg-namespace.js';

/** A part of a chart that can be selected on its own. */
export type ChartPart =
  | { kind: 'series'; seriesIndex: number }
  | { kind: 'point'; seriesIndex: number; pointIndex: number }
  | { kind: 'title' }
  | { kind: 'legend' }
  | { kind: 'axisLabels'; role: 'axis-label' | 'category-label' }
  | { kind: 'gridlines' };

const SERIES_ROLES = new Set(['bar', 'point', 'slice', 'area', 'line']);
const POINT_ROLES = new Set(['bar', 'point', 'slice']);

interface Box {
  x0: number;
  y0: number;
  x1: number;
  y1: number;
}

const FRAME_PAD = 2;
const CIRCLE_R = 4;
const SQUARE = 5;

/** The part under a press. A press on a bar of the series already selected picks that single point. */
export function chartPartAt(
  target: EventTarget | null,
  current: ChartPart | null,
): ChartPart | null {
  if (!(target instanceof Element)) return null;
  const el = target.closest('[data-chart-role]');
  if (!el) return null;
  const role = el.getAttribute('data-chart-role') ?? '';
  const num = (name: string): number | null => {
    const raw = el.getAttribute(name);
    return raw === null || raw === '' ? null : Number(raw);
  };
  if (SERIES_ROLES.has(role)) {
    const seriesIndex = num('data-series-index');
    if (seriesIndex === null) return null;
    const pointIndex = num('data-point-index');
    const onSelectedSeries =
      (current?.kind === 'series' || current?.kind === 'point') &&
      current.seriesIndex === seriesIndex;
    if (onSelectedSeries && pointIndex !== null && POINT_ROLES.has(role)) {
      return { kind: 'point', seriesIndex, pointIndex };
    }
    return { kind: 'series', seriesIndex };
  }
  if (role === 'title') return { kind: 'title' };
  if (role === 'legend') return { kind: 'legend' };
  if (role === 'axis-label' || role === 'category-label') return { kind: 'axisLabels', role };
  if (role === 'axis-tick') return { kind: 'gridlines' };
  return null;
}

const attrNum = (el: Element, name: string): number => Number(el.getAttribute(name) ?? 0);

/** The element's box in scene units: measured when the engine can, otherwise from its attributes. */
function boxOf(el: Element): Box | null {
  const measured = (el as SVGGraphicsElement).getBBox;
  if (typeof measured === 'function') {
    try {
      const b = (el as SVGGraphicsElement).getBBox();
      if (b.width > 0 || b.height > 0)
        return { x0: b.x, y0: b.y, x1: b.x + b.width, y1: b.y + b.height };
    } catch {
      // Not rendered; use the attributes.
    }
  }
  switch (el.tagName.toLowerCase()) {
    case 'rect': {
      const x = attrNum(el, 'x');
      const y = attrNum(el, 'y');
      return { x0: x, y0: y, x1: x + attrNum(el, 'width'), y1: y + attrNum(el, 'height') };
    }
    case 'circle': {
      const r = attrNum(el, 'r');
      const cx = attrNum(el, 'cx');
      const cy = attrNum(el, 'cy');
      return { x0: cx - r, y0: cy - r, x1: cx + r, y1: cy + r };
    }
    case 'line': {
      const xs = [attrNum(el, 'x1'), attrNum(el, 'x2')];
      const ys = [attrNum(el, 'y1'), attrNum(el, 'y2')];
      return { x0: Math.min(...xs), y0: Math.min(...ys), x1: Math.max(...xs), y1: Math.max(...ys) };
    }
    case 'text': {
      const size = attrNum(el, 'font-size') || 10;
      const w = (el.textContent?.length ?? 0) * size * 0.55;
      const x = attrNum(el, 'x');
      const y = attrNum(el, 'y');
      const anchor = el.getAttribute('text-anchor');
      const x0 = anchor === 'middle' ? x - w / 2 : anchor === 'end' ? x - w : x;
      return { x0, y0: y - size, x1: x0 + w, y1: y + size * 0.25 };
    }
    case 'polyline':
    case 'polygon': {
      const n = (el.getAttribute('points') ?? '')
        .split(/[\s,]+/)
        .filter(Boolean)
        .map(Number);
      const xs = n.filter((_, i) => i % 2 === 0);
      const ys = n.filter((_, i) => i % 2 === 1);
      if (xs.length === 0) return null;
      return { x0: Math.min(...xs), y0: Math.min(...ys), x1: Math.max(...xs), y1: Math.max(...ys) };
    }
    default:
      return null;
  }
}

const union = (boxes: readonly Box[]): Box | null =>
  boxes.length === 0
    ? null
    : {
        x0: Math.min(...boxes.map((b) => b.x0)),
        y0: Math.min(...boxes.map((b) => b.y0)),
        x1: Math.max(...boxes.map((b) => b.x1)),
        y1: Math.max(...boxes.map((b) => b.y1)),
      };

const corners = (b: Box): [number, number][] => [
  [b.x0, b.y0],
  [b.x1, b.y0],
  [b.x0, b.y1],
  [b.x1, b.y1],
];

const edges = (b: Box): [number, number][] => {
  const mx = (b.x0 + b.x1) / 2;
  const my = (b.y0 + b.y1) / 2;
  return [...corners(b), [mx, b.y0], [mx, b.y1], [b.x0, my], [b.x1, my]];
};

const padded = (b: Box): Box => ({
  x0: b.x0 - FRAME_PAD,
  y0: b.y0 - FRAME_PAD,
  x1: b.x1 + FRAME_PAD,
  y1: b.y1 + FRAME_PAD,
});

const query = (svg: Element, selector: string): Element[] =>
  Array.from(svg.querySelectorAll(selector));

/** Whether the outer chart handles stay while this part is selected. */
export const partKeepsChartHandles = (part: ChartPart | null): boolean => part?.kind !== 'title';

/** Draw the selection marks of `part` into `svg`, replacing any earlier ones. */
export function paintChartPart(svg: SVGSVGElement, part: ChartPart | null): void {
  for (const old of query(svg, '.fc-chart__parts')) old.remove();
  if (!part) return;
  const g = document.createElementNS(SVG_NS, 'g');
  g.setAttribute('class', 'fc-chart__parts');
  g.setAttribute('pointer-events', 'none');
  const circle = (x: number, y: number): void => {
    const c = document.createElementNS(SVG_NS, 'circle');
    c.setAttribute('class', 'fc-chart__part-handle');
    c.setAttribute('cx', String(x));
    c.setAttribute('cy', String(y));
    c.setAttribute('r', String(CIRCLE_R));
    g.appendChild(c);
  };
  const square = (x: number, y: number): void => {
    const r = document.createElementNS(SVG_NS, 'rect');
    r.setAttribute('class', 'fc-chart__part-handle fc-chart__part-handle--square');
    r.setAttribute('x', String(x - SQUARE / 2));
    r.setAttribute('y', String(y - SQUARE / 2));
    r.setAttribute('width', String(SQUARE));
    r.setAttribute('height', String(SQUARE));
    g.appendChild(r);
  };
  const frame = (b: Box): void => {
    const r = document.createElementNS(SVG_NS, 'rect');
    r.setAttribute('class', 'fc-chart__part-frame');
    r.setAttribute('x', String(b.x0));
    r.setAttribute('y', String(b.y0));
    r.setAttribute('width', String(b.x1 - b.x0));
    r.setAttribute('height', String(b.y1 - b.y0));
    g.appendChild(r);
  };
  const boxes = (selector: string): Box[] => query(svg, selector).flatMap((el) => boxOf(el) ?? []);
  const seriesSel = (i: number): string =>
    [...SERIES_ROLES].map((r) => `[data-chart-role="${r}"][data-series-index="${i}"]`).join(',');

  switch (part.kind) {
    case 'series':
      for (const box of boxes(seriesSel(part.seriesIndex)))
        for (const [x, y] of corners(box)) circle(x, y);
      break;
    case 'point':
      for (const box of boxes(
        [...POINT_ROLES]
          .map(
            (r) =>
              `[data-chart-role="${r}"][data-series-index="${part.seriesIndex}"][data-point-index="${part.pointIndex}"]`,
          )
          .join(','),
      ))
        for (const [x, y] of corners(box)) circle(x, y);
      break;
    case 'title': {
      const box = union(boxes('[data-chart-role="title"]'));
      if (box) {
        const f = padded(box);
        frame(f);
        for (const [x, y] of corners(f)) circle(x, y);
      }
      break;
    }
    case 'legend': {
      const box = union(boxes('[data-chart-role="legend"]'));
      if (box) {
        const f = padded(box);
        frame(f);
        for (const [x, y] of edges(f)) square(x, y);
      }
      break;
    }
    case 'axisLabels': {
      const box = union(boxes(`[data-chart-role="${part.role}"]`));
      if (box) {
        const f = padded(box);
        frame(f);
        for (const [x, y] of corners(f)) circle(x, y);
      }
      break;
    }
    case 'gridlines':
      for (const line of query(svg, 'line[data-chart-role="axis-tick"]')) {
        circle(attrNum(line, 'x1'), attrNum(line, 'y1'));
        circle(attrNum(line, 'x2'), attrNum(line, 'y2'));
      }
      break;
  }
  svg.appendChild(g);
}
