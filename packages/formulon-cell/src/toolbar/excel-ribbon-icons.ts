/**
 * Ribbon icon set.
 *
 * The definitions are composed from `icon-kit`, not written as literal path
 * data: one canvas, one optical size, a five-step stroke scale and a closed
 * palette, with shared structure (arrow heads, sheet frames, corner badges)
 * computed from a single definition. See `icon-kit/tokens.ts` for the rules and
 * `scripts/icons/render.mjs` for the contact sheet used to review changes.
 */

import { RIBBON_GLYPHS } from './icon-kit/glyphs/index.js';
import { PALETTE } from './icon-kit/tokens.js';
import type { IconDefinition, IconSegment } from './icon-kit/types.js';

export type ExcelRibbonIconSegment = IconSegment;
export type ExcelRibbonIconDefinition = IconDefinition;

export const EXCEL_RIBBON_ICON_PATHS: Record<string, ExcelRibbonIconDefinition> = RIBBON_GLYPHS;

export type ExcelRibbonIconName = keyof typeof RIBBON_GLYPHS;

export const excelRibbonIconPaths = (name: string): ExcelRibbonIconDefinition | null =>
  EXCEL_RIBBON_ICON_PATHS[name] ?? null;

/** The neutral roles follow the theme: ink takes the control's text colour,
 *  so it also greys out when the control is disabled, and the paper surfaces
 *  take the toolbar's icon surface tokens. Every other role stays fixed. */
const THEMED_ROLES: Readonly<Record<string, string>> = {
  [PALETTE.ink]: 'currentColor',
  [PALETTE.paper]: 'var(--fc-tb-icon-paper, #ffffff)',
  [PALETTE.mute]: 'var(--fc-tb-icon-mute, #f3f2f1)',
};

const paint = (path: SVGPathElement, property: 'fill' | 'stroke', color: string): void => {
  const themed = THEMED_ROLES[color];
  if (themed?.startsWith('var(')) path.style.setProperty(property, themed);
  else path.setAttribute(property, themed ?? color);
};

export const createExcelRibbonSvg = (name: string, className: string): SVGSVGElement | null => {
  const segments = excelRibbonIconPaths(name);
  if (!segments) return null;

  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.classList.add(className);
  svg.setAttribute('viewBox', '0 0 24 24');
  svg.setAttribute('focusable', 'false');
  svg.setAttribute('aria-hidden', 'true');

  for (const segment of segments) {
    const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    path.setAttribute('d', segment.d);
    paint(path, 'fill', segment.fill ?? 'none');
    if (segment.stroke) paint(path, 'stroke', segment.stroke);
    if (segment.strokeWidth) path.setAttribute('stroke-width', segment.strokeWidth);
    if (segment.strokeLinecap) path.setAttribute('stroke-linecap', segment.strokeLinecap);
    if (segment.strokeLinejoin) path.setAttribute('stroke-linejoin', segment.strokeLinejoin);
    if (segment.strokeDasharray) path.setAttribute('stroke-dasharray', segment.strokeDasharray);
    if (segment.fillRule) path.setAttribute('fill-rule', segment.fillRule);
    if (segment.transform) path.setAttribute('transform', segment.transform);
    svg.appendChild(path);
  }

  return svg;
};
