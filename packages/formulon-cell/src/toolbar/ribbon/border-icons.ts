// Pure SVG icon factories for the Borders dropdown. The data (presets,
// preview shapes, line-style sample swatches) is framework-agnostic and
// has no dependency on the playground's runtime state — only the
// surrounding menu factory in main.ts wires these into the actual DOM.

import type { CellBorderStyle } from '@libraz/formulon-cell';

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

export const SVG_NS = 'http://www.w3.org/2000/svg';
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
  const mk = (offset = 0, width = 1): void => {
    const line = document.createElementNS(SVG_NS, 'line');
    line.setAttribute('x1', String(isHorizontal ? x1 : x1 + offset));
    line.setAttribute('y1', String(isHorizontal ? y1 + offset : y1));
    line.setAttribute('x2', String(isHorizontal ? x2 : x2 + offset));
    line.setAttribute('y2', String(isHorizontal ? y2 + offset : y2));
    line.setAttribute('stroke', 'currentColor');
    line.setAttribute('stroke-width', String(width));
    line.setAttribute('stroke-linecap', 'square');
    svg.appendChild(line);
  };
  if (side === 'thin') mk(0, 1);
  else if (side === 'thick') mk(0, 2);
  else if (side === 'double') {
    mk(-1, 0.75);
    mk(1, 0.75);
  }
};

export const createBorderPreview = (spec: BorderPreviewSpec): SVGSVGElement => {
  const svg = document.createElementNS(SVG_NS, 'svg');
  svg.setAttribute('viewBox', '0 0 16 16');
  svg.setAttribute('width', '16');
  svg.setAttribute('height', '16');
  svg.setAttribute('focusable', 'false');
  svg.setAttribute('aria-hidden', 'true');
  svg.classList.add('app__border-preview');
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
    const v = document.createElementNS(SVG_NS, 'line');
    v.setAttribute('x1', String(x + w / 2));
    v.setAttribute('y1', String(y));
    v.setAttribute('x2', String(x + w / 2));
    v.setAttribute('y2', String(y + h));
    v.setAttribute('stroke', 'currentColor');
    v.setAttribute('stroke-width', '1');
    svg.appendChild(v);
    const hLine = document.createElementNS(SVG_NS, 'line');
    hLine.setAttribute('x1', String(x));
    hLine.setAttribute('y1', String(y + h / 2));
    hLine.setAttribute('x2', String(x + w));
    hLine.setAttribute('y2', String(y + h / 2));
    hLine.setAttribute('stroke', 'currentColor');
    hLine.setAttribute('stroke-width', '1');
    svg.appendChild(hLine);
  }
  drawBorderEdge(svg, spec.top ?? null, x, y, x + w, y);
  drawBorderEdge(svg, spec.bottom ?? null, x, y + h, x + w, y + h);
  drawBorderEdge(svg, spec.left ?? null, x, y, x, y + h);
  drawBorderEdge(svg, spec.right ?? null, x + w, y, x + w, y + h);
  return svg;
};

const createBorderToolIcon = (className: string): SVGSVGElement => {
  const svg = document.createElementNS(SVG_NS, 'svg');
  svg.setAttribute('viewBox', '0 0 16 16');
  svg.setAttribute('width', '16');
  svg.setAttribute('height', '16');
  svg.setAttribute('focusable', 'false');
  svg.setAttribute('aria-hidden', 'true');
  svg.classList.add('app__border-preview', className);
  return svg;
};

const appendPath = (
  svg: SVGSVGElement,
  d: string,
  opts: { fill?: string; stroke?: string; strokeWidth?: string; strokeLinecap?: string } = {},
): void => {
  const path = document.createElementNS(SVG_NS, 'path');
  path.setAttribute('d', d);
  path.setAttribute('fill', opts.fill ?? 'none');
  if (opts.stroke) path.setAttribute('stroke', opts.stroke);
  if (opts.strokeWidth) path.setAttribute('stroke-width', opts.strokeWidth);
  if (opts.strokeLinecap) path.setAttribute('stroke-linecap', opts.strokeLinecap);
  svg.appendChild(path);
};

export const createBorderEraserPreview = (): SVGSVGElement => {
  const svg = createBorderToolIcon('app__border-preview--eraser');
  appendPath(svg, 'M3 10 9.5 3.5l3 3L6 13H3v-3Z', {
    fill: '#f4b6d2',
    stroke: '#8a1f5a',
    strokeWidth: '1',
  });
  appendPath(svg, 'M8.2 4.8l3 3', { stroke: '#ffffff', strokeWidth: '1' });
  return svg;
};

export const createBorderLineColorPreview = (): SVGSVGElement => {
  const svg = createBorderToolIcon('app__border-preview--line-color');
  appendPath(svg, 'M3 12h10', { stroke: '#1f1f1f', strokeWidth: '1.5', strokeLinecap: 'square' });
  appendPath(svg, 'M4 8.5 9.8 2.7l2.1 2.1L6.1 10.6 3.5 11.2 4 8.5Z', {
    fill: '#ffffff',
    stroke: '#2f75b5',
    strokeWidth: '1',
  });
  appendPath(svg, 'M2 14h12', { stroke: '#ed7d31', strokeWidth: '2', strokeLinecap: 'square' });
  return svg;
};

export const createBorderLineStylePreview = (): SVGSVGElement => {
  const svg = createBorderToolIcon('app__border-preview--line-style');
  const styles: [number, string | null, string][] = [
    [4, null, '1'],
    [8, '3 2', '1'],
    [12, null, '2'],
  ];
  for (const [y, dash, width] of styles) {
    const line = document.createElementNS(SVG_NS, 'line');
    line.setAttribute('x1', '2');
    line.setAttribute('y1', String(y));
    line.setAttribute('x2', '14');
    line.setAttribute('y2', String(y));
    line.setAttribute('stroke', 'currentColor');
    line.setAttribute('stroke-width', width);
    line.setAttribute('stroke-linecap', 'square');
    if (dash) line.setAttribute('stroke-dasharray', dash);
    svg.appendChild(line);
  }
  return svg;
};

export const BORDER_PRESETS: Record<string, BorderPreviewSpec> = {
  bottom: { bottom: 'thin' },
  top: { top: 'thin' },
  left: { left: 'thin' },
  right: { right: 'thin' },
  clear: {},
  all: {
    top: 'thin',
    right: 'thin',
    bottom: 'thin',
    left: 'thin',
    innerGrid: true,
    showBase: false,
  },
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
  format: {
    top: 'thin',
    right: 'thin',
    bottom: 'thin',
    left: 'thin',
    innerGrid: true,
    showBase: false,
  },
};

/** Wide horizontal sample for the "線のスタイル / Line style" submenu so
 *  users can compare stroke widths and dash patterns side-by-side. */
export const createLineSamplePreview = (style: CellBorderStyle | 'none'): SVGSVGElement => {
  const svg = document.createElementNS(SVG_NS, 'svg');
  svg.setAttribute('viewBox', '0 0 80 12');
  svg.setAttribute('width', '80');
  svg.setAttribute('height', '12');
  svg.setAttribute('focusable', 'false');
  svg.setAttribute('aria-hidden', 'true');
  svg.classList.add('app__line-sample');
  if (style === 'none') return svg;
  const draw = (yOffset: number, w: number, d: string | null): void => {
    const line = document.createElementNS(SVG_NS, 'line');
    line.setAttribute('x1', '4');
    line.setAttribute('y1', String(yOffset));
    line.setAttribute('x2', '76');
    line.setAttribute('y2', String(yOffset));
    line.setAttribute('stroke', 'currentColor');
    line.setAttribute('stroke-width', String(w));
    if (d) line.setAttribute('stroke-dasharray', d);
    svg.appendChild(line);
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
  'mediumDashed',
  'mediumDashDot',
  'medium',
  'double',
  'thick',
];
