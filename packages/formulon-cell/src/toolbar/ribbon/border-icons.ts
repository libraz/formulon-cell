// Pure SVG icon factories for the Borders dropdown. The data (presets,
// preview shapes, line-style sample swatches) is framework-agnostic and has no
// dependency on runtime state; the menu factory in menus/borders.ts wires these
// into the DOM.

import type { CellBorderStyle } from '../../store/store.js';
import { SVG_NS } from '../../svg-namespace.js';

export type BorderPreviewSide = 'thin' | 'thick' | 'double' | null;

export interface BorderPreviewSpec {
  top?: BorderPreviewSide;
  right?: BorderPreviewSide;
  bottom?: BorderPreviewSide;
  left?: BorderPreviewSide;
  /** When true draws an inner cross dividing the icon into 2×2 cells
   *  (used for the "格子 / All borders" icon). */
  innerGrid?: boolean;
  /** When false suppresses the faint dashed cell-outline base. */
  showBase?: boolean;
}

const svgRoot = (viewBox: string, size: readonly [number, number], ...classes: string[]) => {
  const svg = document.createElementNS(SVG_NS, 'svg');
  svg.setAttribute('viewBox', viewBox);
  svg.setAttribute('width', String(size[0]));
  svg.setAttribute('height', String(size[1]));
  svg.setAttribute('focusable', 'false');
  svg.setAttribute('aria-hidden', 'true');
  svg.classList.add(...classes);
  return svg;
};

const appendLine = (
  svg: SVGSVGElement,
  [x1, y1, x2, y2]: readonly [number, number, number, number],
  opts: { width: number | string; cap?: string; dash?: string | null },
): void => {
  const line = document.createElementNS(SVG_NS, 'line');
  line.setAttribute('x1', String(x1));
  line.setAttribute('y1', String(y1));
  line.setAttribute('x2', String(x2));
  line.setAttribute('y2', String(y2));
  line.setAttribute('stroke', 'currentColor');
  line.setAttribute('stroke-width', String(opts.width));
  if (opts.cap) line.setAttribute('stroke-linecap', opts.cap);
  if (opts.dash) line.setAttribute('stroke-dasharray', opts.dash);
  svg.appendChild(line);
};

const BORDER_ICON_BOX = { x: 2, y: 2, w: 12, h: 12 } as const;
const BORDER_BASE_COLOR = '#c7c7c7';

const drawBorderEdge = (
  svg: SVGSVGElement,
  side: BorderPreviewSide,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
): void => {
  if (!side) return;
  const isHorizontal = y1 === y2;
  const mk = (offset = 0, width = 1): void =>
    appendLine(
      svg,
      [
        isHorizontal ? x1 : x1 + offset,
        isHorizontal ? y1 + offset : y1,
        isHorizontal ? x2 : x2 + offset,
        isHorizontal ? y2 + offset : y2,
      ],
      { width, cap: 'square' },
    );
  if (side === 'thin') mk(0, 1);
  else if (side === 'thick') mk(0, 2);
  else if (side === 'double') {
    mk(-1, 0.75);
    mk(1, 0.75);
  }
};

export const createBorderPreview = (spec: BorderPreviewSpec): SVGSVGElement => {
  const svg = svgRoot('0 0 16 16', [16, 16], 'fc-tb__border-preview');
  const { x, y, w, h } = BORDER_ICON_BOX;
  if (spec.showBase !== false) {
    const base = document.createElementNS(SVG_NS, 'rect');
    base.setAttribute('x', String(x + 0.5));
    base.setAttribute('y', String(y + 0.5));
    base.setAttribute('width', String(w - 1));
    base.setAttribute('height', String(h - 1));
    base.setAttribute('fill', 'none');
    base.setAttribute('stroke', BORDER_BASE_COLOR);
    base.setAttribute('stroke-width', '1');
    base.setAttribute('stroke-dasharray', '1.5 1.5');
    svg.appendChild(base);
  }
  if (spec.innerGrid) {
    appendLine(svg, [x + w / 2, y, x + w / 2, y + h], { width: 1 });
    appendLine(svg, [x, y + h / 2, x + w, y + h / 2], { width: 1 });
  }
  drawBorderEdge(svg, spec.top ?? null, x, y, x + w, y);
  drawBorderEdge(svg, spec.bottom ?? null, x, y + h, x + w, y + h);
  drawBorderEdge(svg, spec.left ?? null, x, y, x, y + h);
  drawBorderEdge(svg, spec.right ?? null, x + w, y, x + w, y + h);
  return svg;
};

const createBorderToolIcon = (className: string): SVGSVGElement => {
  return svgRoot('0 0 16 16', [16, 16], 'fc-tb__border-preview', className);
};

/** Append a `<path>` to `svg`; `fill` defaults to `none` and each stroke attribute is written only when given. */
export const appendPath = (
  svg: SVGSVGElement,
  d: string,
  opts: {
    fill?: string;
    stroke?: string;
    strokeWidth?: string;
    strokeLinecap?: string;
    strokeLinejoin?: string;
    transform?: string;
  } = {},
): void => {
  const path = document.createElementNS(SVG_NS, 'path');
  path.setAttribute('d', d);
  path.setAttribute('fill', opts.fill ?? 'none');
  if (opts.stroke) path.setAttribute('stroke', opts.stroke);
  if (opts.strokeWidth) path.setAttribute('stroke-width', opts.strokeWidth);
  if (opts.strokeLinecap) path.setAttribute('stroke-linecap', opts.strokeLinecap);
  if (opts.strokeLinejoin) path.setAttribute('stroke-linejoin', opts.strokeLinejoin);
  if (opts.transform) path.setAttribute('transform', opts.transform);
  svg.appendChild(path);
};

export const createBorderEraserPreview = (): SVGSVGElement => {
  const svg = createBorderToolIcon('fc-tb__border-preview--eraser');
  appendPath(svg, 'M3 10 9.5 3.5l3 3L6 13H3v-3Z', {
    fill: '#f4b6d2',
    stroke: '#8a1f5a',
    strokeWidth: '1',
  });
  appendPath(svg, 'M8.2 4.8l3 3', { stroke: '#ffffff', strokeWidth: '1' });
  return svg;
};

const PEN_COLOR_BAR = 'fc-tb__pen-color-bar';

/** A pencil over a bar in the colour it applies: the icon of an entry that picks a line or point
 *  colour. `setPenColorBar` repaints the bar. */
export const createPenColorIcon = (className: string, color = '#ed7d31'): SVGSVGElement => {
  const svg = createBorderToolIcon(className);
  appendPath(svg, 'M3 12h10', { stroke: '#1f1f1f', strokeWidth: '1.5', strokeLinecap: 'square' });
  appendPath(svg, 'M4 8.5 9.8 2.7l2.1 2.1L6.1 10.6 3.5 11.2 4 8.5Z', {
    fill: '#ffffff',
    stroke: '#2f75b5',
    strokeWidth: '1',
  });
  appendPath(svg, 'M2 14h12', { stroke: color, strokeWidth: '2', strokeLinecap: 'square' });
  svg.lastElementChild?.classList.add(PEN_COLOR_BAR);
  return svg;
};

export const setPenColorBar = (icon: ParentNode, color: string): void => {
  icon.querySelector(`.${PEN_COLOR_BAR}`)?.setAttribute('stroke', color);
};

export const createBorderLineColorPreview = (): SVGSVGElement =>
  createPenColorIcon('fc-tb__border-preview--line-color');

export const createBorderLineStylePreview = (): SVGSVGElement => {
  const svg = createBorderToolIcon('fc-tb__border-preview--line-style');
  const styles: [number, string | null, string][] = [
    [4, null, '1'],
    [8, '3 2', '1'],
    [12, null, '2'],
  ];
  for (const [y, dash, width] of styles) {
    appendLine(svg, [2, y, 14, y], { width, cap: 'square', dash });
  }
  return svg;
};

const ALL_BORDERS: BorderPreviewSpec = {
  top: 'thin',
  right: 'thin',
  bottom: 'thin',
  left: 'thin',
  innerGrid: true,
  showBase: false,
};

export const BORDER_PRESETS: Record<string, BorderPreviewSpec> = {
  bottom: { bottom: 'thin' },
  top: { top: 'thin' },
  left: { left: 'thin' },
  right: { right: 'thin' },
  clear: {},
  all: ALL_BORDERS,
  outline: {
    top: 'thin',
    right: 'thin',
    bottom: 'thin',
    left: 'thin',
    showBase: false,
  },
  thickOutline: {
    top: 'thick',
    right: 'thick',
    bottom: 'thick',
    left: 'thick',
    showBase: false,
  },
  doubleBottom: { bottom: 'double' },
  thickBottom: { bottom: 'thick' },
  topAndBottom: { top: 'thin', bottom: 'thin' },
  topAndThickBottom: { top: 'thin', bottom: 'thick' },
  topAndDoubleBottom: { top: 'thin', bottom: 'double' },
  format: ALL_BORDERS,
};

/** Wide horizontal sample for the "線のスタイル / Line style" submenu so
 *  users can compare stroke widths and dash patterns side-by-side. */
export const createLineSamplePreview = (style: CellBorderStyle | 'none'): SVGSVGElement => {
  const svg = svgRoot('0 0 80 12', [80, 12], 'fc-tb__line-sample');
  if (style === 'none') return svg;
  const draw = (yOffset: number, w: number, d: string | null): void => {
    appendLine(svg, [4, yOffset, 76, yOffset], { width: w, dash: d });
  };
  if (style === 'double') {
    draw(4, 1, null);
    draw(8, 1, null);
    return svg;
  }
  if (style === 'thin') draw(6, 1, null);
  else if (style === 'hair') draw(6, 0.5, null);
  else if (style === 'medium') draw(6, 1.75, null);
  else if (style === 'thick') draw(6, 2.5, null);
  else if (style === 'dotted') draw(6, 1, '1.2 1.6');
  else if (style === 'dashed') draw(6, 1, '3 2');
  else if (style === 'mediumDashed') draw(6, 1.75, '4 2');
  else if (style === 'dashDot') draw(6, 1, '4 2 1 2');
  else if (style === 'mediumDashDot') draw(6, 1.75, '4 2 1 2');
  else if (style === 'dashDotDot') draw(6, 1, '4 2 1 2 1 2');
  else if (style === 'mediumDashDotDot') draw(6, 1.75, '4 2 1 2 1 2');
  else if (style === 'slantDashDot') draw(6, 1, '4 2 2 2');
  return svg;
};

export const LINE_STYLES_ALL: (CellBorderStyle | 'none')[] = [
  'none',
  'thin',
  'hair',
  'dotted',
  'dashed',
  'dashDot',
  'dashDotDot',
  'double',
  'medium',
  'mediumDashed',
  'mediumDashDot',
  'mediumDashDotDot',
  'slantDashDot',
  'thick',
];
