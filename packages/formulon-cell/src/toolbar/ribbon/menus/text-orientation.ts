// Text-orientation dropdown for the Home tab. The menu items render a small
// inline SVG glyph showing the orientation effect (counter-clockwise, vertical,
// rotate up/down, etc.) plus the localized label.

import type { ToolbarMenuText } from '../../../index.js';

import { SVG_NS } from '../border-icons.js';
import { createMenu, menuPresetButton, menuSeparator } from './general.js';

export type TextOrientationGlyph = 'ccw' | 'cw' | 'vertical' | 'up' | 'down' | 'format';

const ink = '#1f1f1f';
const grid = '#8a8f98';
const gridLight = '#d9d9d9';
const excelGreen = '#107c41';
const excelBlue = '#2f75b5';

type TextOrientationPathAttrs = {
  fill?: string;
  stroke?: string;
  strokeWidth?: string;
  strokeLinecap?: 'butt' | 'round' | 'square';
  strokeLinejoin?: 'bevel' | 'miter' | 'round';
  transform?: string;
};

const appendPath = (
  svg: SVGSVGElement,
  d: string,
  {
    fill = 'none',
    stroke,
    strokeWidth = '1.25',
    strokeLinecap = 'round',
    strokeLinejoin = 'round',
    transform,
  }: TextOrientationPathAttrs = {},
): void => {
  const path = document.createElementNS(SVG_NS, 'path');
  path.setAttribute('d', d);
  path.setAttribute('fill', fill);
  if (stroke) {
    path.setAttribute('stroke', stroke);
    path.setAttribute('stroke-width', strokeWidth);
    path.setAttribute('stroke-linecap', strokeLinecap);
    path.setAttribute('stroke-linejoin', strokeLinejoin);
  }
  if (transform) path.setAttribute('transform', transform);
  svg.appendChild(path);
};

const createTextOrientationIcon = (glyph: TextOrientationGlyph): SVGSVGElement => {
  const svg = document.createElementNS(SVG_NS, 'svg');
  svg.setAttribute('viewBox', '0 0 16 16');
  svg.setAttribute('width', '16');
  svg.setAttribute('height', '16');
  svg.classList.add('fc-tb__text-orientation-preview');
  svg.setAttribute('focusable', 'false');
  svg.setAttribute('aria-hidden', 'true');
  appendPath(svg, 'M2 13h12', { stroke: grid, strokeWidth: '1.25' });
  if (glyph === 'ccw' || glyph === 'cw') {
    const angle = glyph === 'ccw' ? -35 : 35;
    appendPath(svg, 'M4.2 9.8h5.6M4.2 7.7h4.2M10.5 7.1v5.1', {
      stroke: ink,
      strokeWidth: '1.45',
      transform: `rotate(${angle} 8 10)`,
    });
    appendPath(
      svg,
      glyph === 'ccw'
        ? 'M5.1 5.4 3.3 7.2l1.8 1.8M3.4 7.2h5.4'
        : 'M10.9 5.4l1.8 1.8-1.8 1.8M7.2 7.2h5.4',
      { stroke: excelGreen, strokeWidth: '1.45' },
    );
  } else if (glyph === 'vertical') {
    appendPath(svg, 'M7.2 3.2h1.6v1.6H7.2zM7.2 6.2h1.6v1.6H7.2zM7.2 9.2h1.6v1.6H7.2z', {
      fill: ink,
    });
    appendPath(svg, 'M5.2 3v8M10.8 3v8', { stroke: excelGreen, strokeWidth: '1.15' });
  } else if (glyph === 'up' || glyph === 'down') {
    const rotate = glyph === 'up' ? -90 : 90;
    appendPath(svg, 'M4.3 10h5.9M4.3 7.9h4.4M10.8 7.4v5', {
      stroke: ink,
      strokeWidth: '1.45',
      transform: `translate(8 10) rotate(${rotate}) translate(-8 -10)`,
    });
    appendPath(
      svg,
      glyph === 'up'
        ? 'M12 10.8V4.2M9.9 6.3 12 4.2l2.1 2.1'
        : 'M12 4.2v6.6M9.9 8.7l2.1 2.1 2.1-2.1',
      {
        stroke: excelGreen,
        strokeWidth: '1.45',
      },
    );
  } else if (glyph === 'format') {
    appendPath(svg, 'M2.5 3.5h11v7h-11z', { fill: '#ffffff', stroke: ink, strokeWidth: '1.05' });
    appendPath(svg, 'M2.5 7h11M8 3.5v7', { stroke: gridLight, strokeWidth: '1' });
    appendPath(svg, 'M4.2 5.3h2.2M9.7 8.7h2.1', { stroke: excelBlue, strokeWidth: '1.2' });
    appendPath(svg, 'M11.4 4.3 13.2 2.5v3.4', { stroke: excelGreen, strokeWidth: '1.25' });
  }
  return svg;
};

const textOrientationMenuItem = (
  glyph: TextOrientationGlyph,
  label: string,
  value: string,
): HTMLButtonElement => {
  return menuPresetButton(label, 'textOrientation', value, createTextOrientationIcon(glyph));
};

export const createTextOrientationMenu = (t: ToolbarMenuText): HTMLDivElement => {
  const menu = createMenu('menu-text-orientation');
  menu.append(
    textOrientationMenuItem('ccw', t.orientationAngleCounterclockwise, 'ccw'),
    textOrientationMenuItem('cw', t.orientationAngleClockwise, 'cw'),
    textOrientationMenuItem('vertical', t.orientationVerticalText, 'vertical'),
    textOrientationMenuItem('up', t.orientationRotateTextUp, 'up'),
    textOrientationMenuItem('down', t.orientationRotateTextDown, 'down'),
    menuSeparator(),
    textOrientationMenuItem('format', t.orientationFormatAlignment, 'format'),
  );
  return menu;
};
